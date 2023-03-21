using ATTAS_CORE;
using Excel = Microsoft.Office.Interop.Excel;

/*
################################
||           ATTAS            ||
################################
*/
ATTAS_ORTOOLS attas = new ATTAS_ORTOOLS();

attas.objOption = new int[6] { 0, 1, 1, 0, 1, 1 };
attas.objWeight = new int[6] { 1, 1, 1, 1, 1, 1 };
attas.maxSearchingTimeOption = 600.0;
attas.debugLoggerOption = true;
attas.strategyOption = 2;


const string inputExcelPath = @"D:\FPT\SEP490_G14\ATTAS_NSGA2_CDP\inputs\inputCF_SU23.xlsx";
const string outputExcelPath = @"D:\FPT\SEP490_G14\rawprocess\result.xlsx";

try
{
    /*
    ################################
    ||       READING EXCEL        ||
    ################################
    */
    Excel.Application oXL;
    Excel._Workbook oWB;
    Excel._Worksheet oWS;
    //Start Excel and get Application object.
    oXL = new Excel.Application();
    oWB = oXL.Workbooks.Open(inputExcelPath);

    Console.WriteLine($"ATTAS - Reading Data From Excel {inputExcelPath}");

    attas.numTasks = (int)oWB.Sheets[1].Cells[1, 2].Value2;
    attas.numInstructors = (int)oWB.Sheets[1].Cells[2, 2].Value2;
    attas.numSlots = (int)oWB.Sheets[1].Cells[3, 2].Value2;
    attas.numSubjects = (int)oWB.Sheets[1].Cells[4, 2].Value2;
    attas.numAreas = (int)oWB.Sheets[1].Cells[5, 2].Value2;
    attas.numBackupInstructors = (int)oWB.Sheets[1].Cells[6, 2].Value2; ;

    string[] classNames = excelToNameArray((Excel._Worksheet)oWB.Sheets[2], attas.numTasks, true,2,1);
    string[] slotNames = excelToNameArray((Excel._Worksheet)oWB.Sheets[3],attas.numSlots,true , 2,1);
    string[] instructorNames = excelToNameArray((Excel._Worksheet)oWB.Sheets[5], attas.numInstructors, true,2,1);
    string[] subjectNames = excelToNameArray((Excel._Worksheet)oWB.Sheets[5], attas.numSubjects, false ,1, 2);
    // SLOT
    attas.slotConflict = excelToArray((Excel._Worksheet)oWB.Sheets[3], 2, 2, attas.numSlots, attas.numSlots);
    attas.slotCompatibilityCost = excelToArray((Excel._Worksheet)oWB.Sheets[4], 2, 2, attas.numSlots, attas.numSlots);
    // INSTRUCTOR
    attas.instructorSubjectPreference = excelToArray((Excel._Worksheet)oWB.Sheets[5], 2, 2, attas.numInstructors, attas.numSubjects);
    attas.instructorSubject = toBinaryArray(attas.instructorSubjectPreference);
    attas.instructorSlotPreference = excelToArray((Excel._Worksheet)oWB.Sheets[6], 2, 2, attas.numInstructors, attas.numSlots);
    attas.instructorSlot = toBinaryArray(attas.instructorSlotPreference);
    attas.instructorQuota = flattenArray(excelToArray((Excel._Worksheet)oWB.Sheets[7], 2, 2, attas.numInstructors, 1));

    attas.instructorPreassign = new List<(int, int, int)>();
    for (int i=0; i<attas.numInstructors; i++)
        for(int j=0; j < attas.numSlots; j++)
        {
            var content = oWB.Sheets[8].Cells[i + 2, j + 2].Value2;
            if (content != null)
            {
                attas.instructorPreassign.Add((i, (int)content-1, 1));
            }
        }
    //attas.instructorPreassign = new List<(int, int, int)> { (32, 0, 1), (32, 1, 1), (32, 2, 1) };

    // AREA
    attas.areaDistance = excelToArray((Excel._Worksheet)oWB.Sheets[9], 2, 2, attas.numAreas, attas.numAreas);
    attas.areaSlotWeight = excelToArray((Excel._Worksheet)oWB.Sheets[10], 2, 2, attas.numSlots, attas.numSlots);
    // TASK
    attas.taskSubjectMapping = excelToMapping((Excel._Worksheet)oWB.Sheets[2], attas.numTasks, 2, subjectNames);
    attas.taskSlotMapping = excelToMapping((Excel._Worksheet)oWB.Sheets[2], attas.numTasks, 4, slotNames);
    attas.taskAreaMapping = new int[attas.numTasks];
    for(int i = 0;i < attas.numTasks;i++)
        attas.taskAreaMapping[i] = 1;

    
    oWB.Close();
    oXL.Quit();
    Console.WriteLine("Done!");

    /*
    ################################
    ||          SOLVING           ||
    ################################
    */
    Console.WriteLine("ATTAS - Start Solving");
    List<List<(int, int)>>? results = attas.solve();
    /*
    ################################
    ||       EXPORT RESULT        ||
    ################################
    */
    if (results != null)
    {
        Console.WriteLine($"ATTAS - Start Export Result To Excel {outputExcelPath}");
        oXL = new Excel.Application();
        oWB = oXL.Workbooks.Open(outputExcelPath);
        oWS = oWB.Sheets.Add();

        DateTime currentTime = DateTime.Now;
        string currentTimeString = currentTime.ToString("yyyy-MM-dd_HH-mm-ss");
        oWS.Name = $"result_{currentTimeString}";

        for (int i = 0; i < attas.numInstructors; i++)
        {
            oWS.Cells[i + 2, 1] = instructorNames[i];
            oWS.Cells[i + 2, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkOrange);
            alignMiddle(oWS.Cells[i + 2, 1]);
        }

        oWS.Cells[attas.numInstructors + 2, 1] = "UNASSIGNED";
        oWS.Cells[attas.numInstructors + 2, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkOrange);
        alignMiddle(oWS.Cells[attas.numInstructors + 2, 1]);

        for (int i = 0;i < attas.numSlots; i++)
        {
            oWS.Cells[1,i+2] = slotNames[i];
            oWS.Cells[1,i+2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SteelBlue);
            alignMiddle(oWS.Cells[1, i + 2]);
        }

        for (int i = 0; i <= attas.numInstructors + 1; i++)
            for (int j = 0; j <= attas.numSlots; j++)
            {
                fullBorder(oWS.Cells[i + 1, j + 1]);
            }
        List<(int, int)> tmp = results[0];
        foreach ((int, int) result in tmp)
            if (result.Item2 >= 0)
            {
                oWS.Cells[result.Item2 + 2, attas.taskSlotMapping[result.Item1] + 2] = $"{result.Item1 + 1}.{classNames[result.Item1]}.{subjectNames[attas.taskSubjectMapping[result.Item1]]}";
                //Console.WriteLine($"Tasks {result.Item1} assigned to instructor {result.Item2}");
            }
            else
            {
                oWS.Cells[attas.numInstructors + 2, attas.taskSlotMapping[result.Item1] + 2] = oWS.Cells[attas.numInstructors + 2, attas.taskSlotMapping[result.Item1] + 2].Value + $"{result.Item1 + 1}.{classNames[result.Item1]}.{subjectNames[attas.taskSubjectMapping[result.Item1]]}\n";
                //Console.WriteLine($"Tasks {result.Item1} need backup instructor!");
            }

        oWS.Columns.AutoFit();
        oWB.Save();
        oWB.Close();
        oXL.Quit();
        Console.WriteLine("Done");
    }
    else
    {
        Console.WriteLine("No Solution!");
    }
}
catch (Exception ex)
{
    Console.WriteLine($"An exception occurred: {ex.Message}");
}

/*
################################
||       Excel Utility        ||
################################
*/
static int[] excelToMapping(Excel._Worksheet oSheet,int numRows,int col, string[] namesArray)
{
    int[] mapping = new int[numRows];
    Excel.Range oRng;
    for (int i = 2; i<=numRows+1; i++)
    {
        oRng = oSheet.Cells[i, col];
        mapping[i - 2]= Array.IndexOf(namesArray, oRng.Value2);
    }
    return mapping;
}
static int[,] excelToArray(Excel._Worksheet oSheet, int startRow, int startCol, int numRows, int numCols)
{
    Excel.Range oRng;
    oRng = oSheet.Cells[startRow, startCol].Resize[numRows, numCols];
    object[,] values = (object[,])oRng.Value;
    int[,] data = new int[numRows, numCols];
    for (int i = 1; i <= numRows; i++)
    {
        for (int j = 1; j <= numCols; j++)
        {
            data[i - 1, j - 1] = (int)(double)values[i, j];
        }
    }
    return data;
}
static string[] excelToNameArray(Excel._Worksheet oSheet, int count, bool isColumn, int posrow, int poscol)
{
    string[] data = new string[count];
    Excel.Range oRng;
    if (isColumn)
    {
        oRng = oSheet.Cells[posrow, poscol].Resize[count, 1];
        object[,] values = (object[,])oRng.Value;
        for (int i = 1; i <= count; i++)
        {
            data[i - 1] = (string)values[i, 1];
        }
    }
    else
    {
        oRng = oSheet.Cells[posrow, poscol].Resize[1, count];
        object[,] values = (object[,])oRng.Value;
        for (int i = 1; i <= count; i++)
        {
            data[i - 1] = (string)values[1, i];
        }
    }
    return data;
}
static void alignMiddle(Excel.Range range)
{
    range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
}
static void fullBorder(Excel.Range range)
{
    // Set the border style, weight, and color
    Excel.XlLineStyle lineStyle = Excel.XlLineStyle.xlContinuous;
    Excel.XlBorderWeight lineWeight = Excel.XlBorderWeight.xlThin;
    object lineColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

    // Add the border to the top edge of the range
    Excel.Border topBorder = range.Borders[Excel.XlBordersIndex.xlEdgeTop];
    topBorder.LineStyle = lineStyle;
    topBorder.Weight = lineWeight;
    topBorder.Color = lineColor;

    // Add the border to the bottom edge of the range
    Excel.Border bottomBorder = range.Borders[Excel.XlBordersIndex.xlEdgeBottom];
    bottomBorder.LineStyle = lineStyle;
    bottomBorder.Weight = lineWeight;
    bottomBorder.Color = lineColor;

    // Add the border to the left edge of the range
    Excel.Border leftBorder = range.Borders[Excel.XlBordersIndex.xlEdgeLeft];
    leftBorder.LineStyle = lineStyle;
    leftBorder.Weight = lineWeight;
    leftBorder.Color = lineColor;

    // Add the border to the right edge of the range
    Excel.Border rightBorder = range.Borders[Excel.XlBordersIndex.xlEdgeRight];
    rightBorder.LineStyle = lineStyle;
    rightBorder.Weight = lineWeight;
    rightBorder.Color = lineColor;
}
/*
################################
||          Utility           ||
################################
*/
static int[,] toBinaryArray(int[,] data)
{
    int numRows = data.GetLength(0);
    int numColumns = data.GetLength(1);
    int[,] result = new int[numRows, numColumns];
    for (int i = 0; i < numRows; i++)
        for (int j = 0; j < numColumns; j++)
            if (data[i, j] > 0)
            {
                result[i, j] = 1;
            }
            else
            {
                result[i, j] = 0;
            }
    return result;
}
static int[] flattenArray(int[,] data)
{
    int numRows = data.GetLength(0);
    int numColumns = data.GetLength(1);
    int[] flattened = new int[numRows * numColumns];
    int k = 0;
    for (int i = 0; i < numRows; i++)
    {
        for (int j = 0; j < numColumns; j++)
        {
            flattened[k++] = data[i, j];
        }
    }
    return flattened;
}

