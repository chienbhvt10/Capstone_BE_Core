using ATTAS_CORE;
using Excel = Microsoft.Office.Interop.Excel;

/*
################################
||           ATTAS            ||
################################
*/
ATTAS attas = new ATTAS();

attas.objOption = new int[6] { 0, 0, 0, 0, 0, 0 };
attas.maxSearchingTimeOption = 120.0;
attas.debugLoggerOption = true;
attas.solverOption = 1;
attas.strategyOption = 2;

attas.numSubjects = 14;
attas.numTasks = 162;
attas.numSlots = 22;
attas.numInstructors = 34;
attas.numBackupInstructors = 5;
attas.numAreas = 3;

const string inputExcelPath = @"D:\FPT\SEP490_G14\input.xlsx";
const string outputExcelPath = @"D:\FPT\SEP490_G14\result.xlsx";

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
  
    string[] classNames = excelToNameArray((Excel._Worksheet)oWB.Sheets[1], attas.numTasks, true,2,1);
    string[] slotNames = excelToNameArray((Excel._Worksheet)oWB.Sheets[2],attas.numSlots,true , 2,1);
    string[] instructorNames = excelToNameArray((Excel._Worksheet)oWB.Sheets[4], attas.numInstructors, true,2,1);
    string[] subjectNames = excelToNameArray((Excel._Worksheet)oWB.Sheets[4], attas.numSubjects, false ,1, 2);
    // SLOT
    attas.slotConflict = excelToArray((Excel._Worksheet)oWB.Sheets[2], 2, 2, attas.numSlots, attas.numSlots);
    attas.slotCompatibilityCost = excelToArray((Excel._Worksheet)oWB.Sheets[3], 2, 2, attas.numSlots, attas.numSlots);
    // INSTRUCTOR
    attas.instructorSubjectPreference = excelToArray((Excel._Worksheet)oWB.Sheets[4], 2, 2, attas.numInstructors, attas.numSubjects);
    attas.instructorSubject = toBinaryArray(attas.instructorSubjectPreference);
    attas.instructorSlotPreference = excelToArray((Excel._Worksheet)oWB.Sheets[5], 2, 2, attas.numInstructors, attas.numSlots);
    attas.instructorSlot = toBinaryArray(attas.instructorSlotPreference);
    attas.instructorQuota = flattenArray(excelToArray((Excel._Worksheet)oWB.Sheets[6], 2, 2, attas.numInstructors, 1));
    //attas.instructorPreassign = new List<(int, int, int)> { (32, 0, 1), (32, 1, 1), (32, 2, 1) };
    // AREA
    attas.areaDistance = excelToArray((Excel._Worksheet)oWB.Sheets[8], 2, 2, attas.numAreas, attas.numAreas);
    attas.areaSlotWeight = excelToArray((Excel._Worksheet)oWB.Sheets[9], 2, 2, attas.numSlots, attas.numSlots);
    // TASK
    attas.taskSubjectMapping = new int[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 2, 3, 4, 5, 6, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 9, 9, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13 };
    attas.taskSlotMapping = new int[] { 9, 11, 16, 19, 16, 6, 12, 14, 1, 3, 9, 11, 17, 19, 4, 6, 12, 14, 1, 3, 9, 11, 17, 19, 4, 6, 12, 14, 1, 3, 9, 11, 17, 19, 4, 6, 12, 14, 1, 3, 9, 17, 11, 4, 19, 12, 6, 1, 14, 0, 3, 17, 11, 4, 19, 6, 12, 14, 1, 3, 9, 11, 19, 17, 4, 8, 6, 16, 5, 14, 13, 3, 2, 7, 15, 3, 14, 3, 17, 11, 4, 19, 12, 6, 1, 14, 9, 15, 16, 1, 14, 15, 10, 10, 5, 11, 4, 19, 12, 6, 0, 14, 9, 3, 17, 3, 9, 8, 2, 14, 15, 15, 17, 2, 5, 10, 13, 18, 0, 7, 8, 7, 8, 19, 16, 13, 13, 0, 7, 8, 10, 16, 2, 5, 10, 13, 18, 0, 18, 0, 7, 8, 5, 5, 13, 6, 9, 11, 17, 3, 4, 11, 12, 19, 1, 19, 1, 6, 9, 4, 4, 12 };
    attas.taskAreaMapping = new int[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 2, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 1, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 1, 1, 1, 1, 2, 1, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 1, 2, 2, 1, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 1, 1, 2, 1, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2 };
    
    oWB.Close();
    oXL.Quit();
    Console.WriteLine("Done!");
    /*
    ################################
    ||          SOLVING           ||
    ################################
    */
    Console.WriteLine("ATTAS - Start Solving");
    List<(int, int)>? results = attas.solve();
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

        foreach ((int, int) result in results)
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

