using ATTAS_CORE;
using Spectre.Console;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

const string inputExcelFilePath = @"D:\FPT\SEP490_G14\ATTAS_ORTOOLS\inputs\inputCF_SU23_NEW.xlsx";
const string outputExcelFolderPath = @"D:\FPT\SEP490_G14\ATTAS_ORTOOLS\results";

ATTAS_ORTOOLS attas = new ATTAS_ORTOOLS();

attas.objOption = new int[8] { 0, 0, 0, 0, 0, 0, 0, 0 };
attas.objWeight = new int[8] { 50, 25, 1, 1, 1, 1, 1, 1 };
attas.maxSearchingTimeOption = 300.0;
attas.strategyOption = 2;

string[] classNames = Array.Empty<string>();
string[] slotNames = Array.Empty<string>();
string[] instructorNames = Array.Empty<string>();
string[] subjectNames = Array.Empty<string>();

AnsiConsole.Write(new FigletText("ATTAS").LeftJustified().Color(Color.Gold1));
string choice;
bool read = false;
List<List<(int, int)>>? results = null;
do
{
    // Ask for the user's favorite fruit
    choice = AnsiConsole.Prompt(
        new SelectionPrompt<string>()
            .Title("")
            .PageSize(10)
            .MoreChoicesText("[grey](Move up and down to select option)[/]")
            .AddChoices(new[] {
            "Input", "Solve", "Output","Setting", "Quit"
            }));
    switch(choice)
    {
        case "Input":
            read = readInputExcel(inputExcelFilePath, attas, ref classNames, ref slotNames, ref instructorNames, ref subjectNames);
            cleanCOM();
            break;
        case "Solve":
            var rule = new Rule("Solve");
            rule.LeftJustified();
            AnsiConsole.Write(rule);
            if (read)
            {
                results = solve(attas);
            }
            else
            {
                AnsiConsole.Markup("\n[red]No input to solve[/]\n\n");
            }
            break;
        case "Output":
            writeOutputExcel(outputExcelFolderPath, attas, results,ref classNames,ref slotNames,ref instructorNames,ref subjectNames);
            cleanCOM();
            break;
    }
}
while (choice != "Quit");

static bool readInputExcel(string inputPath,ATTAS_ORTOOLS attas,ref string[] classNames,ref string[] slotNames,ref string[] instructorNames,ref string[] subjectNames)
{
    Application? oXL = null;
    Workbook? oWB = null;
    try
    {
        var rule = new Rule("Input");
        rule.LeftJustified();
        AnsiConsole.Write(rule);
        AnsiConsole.Markup($"\n Reading Data From [underline green]{inputPath}[/]\n\n");
        oXL = new Application();
        oWB = oXL.Workbooks.Open(inputExcelFilePath);
        Worksheet oWS_inputInfo = oWB.Sheets[1];
        Worksheet oWS_tasks = oWB.Sheets[2];
        Worksheet oWS_slotConflict = oWB.Sheets[3];
        Worksheet oWS_slotDay = oWB.Sheets[4];
        Worksheet oWS_slotTime = oWB.Sheets[5];
        Worksheet oWS_slotSegment = oWB.Sheets[6];
        Worksheet oWS_patternCost = oWB.Sheets[7];
        Worksheet oWS_instructorSubject = oWB.Sheets[8];
        Worksheet oWS_instructorSlot = oWB.Sheets[9];
        Worksheet oWS_instructorQuota = oWB.Sheets[10];
        Worksheet oWS_instructorPreassign = oWB.Sheets[11];
        Worksheet oWS_areaDistance = oWB.Sheets[12];
        Worksheet oWS_areaSlotCoefficient = oWB.Sheets[13];
        attas.numTasks = (int)oWS_inputInfo.Cells[1, 2].Value2;
        attas.numInstructors = (int)oWS_inputInfo.Cells[2, 2].Value2;
        attas.numSlots = (int)oWS_inputInfo.Cells[3, 2].Value2;
        attas.numDays = (int)oWS_inputInfo.Cells[4, 2].Value2;
        attas.numTimes = (int)oWS_inputInfo.Cells[5, 2].Value2;
        attas.numSegments = (int)oWS_inputInfo.Cells[6, 2].Value2;
        int numSlotSegmentRules = (int)oWS_inputInfo.Cells[7, 2].Value2;
        attas.numSubjects = (int)oWS_inputInfo.Cells[8, 2].Value2;
        attas.numAreas = (int)oWS_inputInfo.Cells[9, 2].Value2;
        attas.numBackupInstructors = (int)oWS_inputInfo.Cells[10, 2].Value2;
        // NAME
        classNames = excelToNameArray(oWS_tasks, attas.numTasks, true, 2, 1);
        slotNames = excelToNameArray(oWS_slotConflict, attas.numSlots, true, 2, 1);
        instructorNames = excelToNameArray(oWS_instructorSubject, attas.numInstructors, true, 2, 1);
        subjectNames = excelToNameArray(oWS_instructorSubject, attas.numSubjects, false, 1, 2);
        // SLOT
        attas.slotConflict = excelToArray(oWS_slotConflict, 2, 2, attas.numSlots, attas.numSlots);
        attas.slotDay = excelToArray(oWS_slotDay, 2, 2, attas.numSlots, attas.numDays);
        attas.slotTime = excelToArray(oWS_slotTime, 2, 2, attas.numSlots, attas.numTimes);
        attas.slotSegment = new int[attas.numSlots, attas.numDays, attas.numSegments];
        for (int i = 0; i < numSlotSegmentRules; i++)
        {
            int slot = Array.IndexOf(slotNames, (string)oWS_slotSegment.Cells[i + 2, 1].Value2);
            int day = (int)(double)oWS_slotSegment.Cells[i + 2, 2].Value2 - 1;
            int segment = (int)(double)oWS_slotSegment.Cells[i + 2, 3].Value2 - 1;
            attas.slotSegment[slot, day, segment] = 1;
        }
        attas.patternCost = flattenArray(excelToArray(oWS_patternCost, 2, 2, (1 << attas.numSegments), 1));
        // INSTRUCTOR
        attas.instructorSubjectPreference = excelToArray(oWS_instructorSubject, 2, 2, attas.numInstructors, attas.numSubjects);
        attas.instructorSubject = toBinaryArray(attas.instructorSubjectPreference);
        attas.instructorSlotPreference = excelToArray(oWS_instructorSlot, 2, 2, attas.numInstructors, attas.numSlots);
        attas.instructorSlot = toBinaryArray(attas.instructorSlotPreference);
        attas.instructorQuota = flattenArray(excelToArray(oWS_instructorQuota, 2, 3, attas.numInstructors, 1));
        attas.instructorMinQuota = flattenArray(excelToArray(oWS_instructorQuota, 2, 2, attas.numInstructors, 1));
        attas.instructorPreassign = new List<(int, int, int)>();
        for (int i = 0; i < attas.numInstructors; i++)
            for (int j = 0; j < attas.numSlots; j++)
            {
                var content = oWS_instructorPreassign.Cells[i + 2, j + 2].Value2;
                if (content != null)
                {
                    attas.instructorPreassign.Add((i, (int)content - 1, 1));
                }
            }
        // AREA
        attas.areaDistance = excelToArray(oWS_areaDistance, 2, 2, attas.numAreas, attas.numAreas);
        attas.areaSlotCoefficient = excelToArray(oWS_areaSlotCoefficient, 2, 2, attas.numSlots, attas.numSlots);
        // TASK
        attas.taskSubjectMapping = excelToMapping(oWS_tasks, attas.numTasks, 2, subjectNames);
        attas.taskSlotMapping = excelToMapping(oWS_tasks, attas.numTasks, 4, slotNames);
        attas.taskAreaMapping = new int[attas.numTasks];
        for (int i = 0; i < attas.numTasks; i++)
            attas.taskAreaMapping[i] = 1;
        oWB.Close();
        oXL.Quit();
        return true;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"An exception occurred while writing output: {ex.Message}");
        if ( oWB != null ) 
            oWB.Close();
        if ( oXL != null ) 
            oXL.Quit();
        return false;
    }
}
static List<List<(int, int)>>? solve(ATTAS_ORTOOLS attas)
{
    List<List<(int, int)>>? results = null;
    AnsiConsole.Status()
    .Start("Solving...\n", ctx =>
    {
        results = attas.solve();
    });
    object[] statistics = attas.getStatistic();
    // Create a table
    var table = new Table();

    // Add some columns
    table.AddColumn("Stat");
    table.AddColumn("Value");

    // Add some rows
    table.AddRow("Objective", $"{statistics[0]}");
    table.AddRow("Status", $"{statistics[1]}");
    table.AddRow("Conflicts", $"{statistics[2]}");
    table.AddRow("Branches", $"{statistics[3]}");
    table.AddRow("Wall Time", $"{statistics[4]}s");

    // Render the table to the console
    AnsiConsole.Write(table);
    return results;
}
static void writeOutputExcel(string outputPath,ATTAS_ORTOOLS attas, List<List<(int, int)>>? results,ref string[] classNames,ref string[] slotNames,ref string[] instructorNames,ref string[] subjectNames) 
{
    var line = new Rule("Output");
    line.LeftJustified();
    AnsiConsole.Write(line);
    if (results != null)
    {
        Application? oXL = null;
        Workbook? oWB = null;
        try
        {
            DateTime currentTime = DateTime.Now;
            string currentTimeString = currentTime.ToString("yyyy-MM-ddTHH-mm-ss");
            AnsiConsole.Markup($" - Start Export Result Into [underline green]{outputExcelFolderPath}\\result_{currentTimeString}.xlsx[/]\n");
            oXL = new Application();
            oWB = oXL.Workbooks.Add();
            Worksheet oWS = oWB.ActiveSheet;
            oWS.Name = "result";
            for (int i = 0; i < attas.numInstructors; i++)
            {
                oWS.Cells[i + 2, 1] = instructorNames[i];
                oWS.Cells[i + 2, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkOrange);
                alignMiddle(oWS.Cells[i + 2, 1]);
            }

            oWS.Cells[attas.numInstructors + 2, 1] = "UNASSIGNED";
            oWS.Cells[attas.numInstructors + 2, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkOrange);
            alignMiddle(oWS.Cells[attas.numInstructors + 2, 1]);

            for (int i = 0; i < attas.numSlots; i++)
            {
                oWS.Cells[1, i + 2] = slotNames[i];
                oWS.Cells[1, i + 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SteelBlue);
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
                }
                else
                {
                    oWS.Cells[attas.numInstructors + 2, attas.taskSlotMapping[result.Item1] + 2] = oWS.Cells[attas.numInstructors + 2, attas.taskSlotMapping[result.Item1] + 2].Value + $"{result.Item1 + 1}.{classNames[result.Item1]}.{subjectNames[attas.taskSubjectMapping[result.Item1]]}\n";
                }

            oWS.Columns.AutoFit();
            oWB.SaveAs($@"{outputExcelFolderPath}\result_{currentTimeString}.xlsx");
            oWB.Close();
            oXL.Quit();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An exception occurred while writing output: {ex.ToString()}");                
            if (oWB != null)
                oXL.DisplayAlerts = false;
                oWB.Close();
            if (oXL != null) 
                oXL.DisplayAlerts = true;
                oXL.Quit();
                
        }
    }
    else
    {
        AnsiConsole.Markup("\n[red]No solution to export[/]\n\n");
    }
}
/*
################################
||       Excel Utility        ||
################################
*/
static int[] excelToMapping(Worksheet oSheet,int numRows,int col, string[] namesArray)
{
    int[] mapping = new int[numRows];
    Microsoft.Office.Interop.Excel.Range oRng;
    for (int i = 2; i<=numRows+1; i++)
    {
        oRng = oSheet.Cells[i, col];
        mapping[i - 2]= Array.IndexOf(namesArray, oRng.Value2);
    }
    return mapping;
}
static int[,] excelToArray(Worksheet oSheet, int startRow, int startCol, int numRows, int numCols)
{
    Range oRng;
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
static string[] excelToNameArray(Worksheet oSheet, int count, bool isColumn, int posrow, int poscol)
{
    string[] data = new string[count];
    Range oRng;
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
static void alignMiddle(Range range)
{
    range.VerticalAlignment = XlVAlign.xlVAlignCenter;
    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
}
static void fullBorder(Range range)
{
    // Set the border style, weight, and color
    XlLineStyle lineStyle = XlLineStyle.xlContinuous;
    XlBorderWeight lineWeight = XlBorderWeight.xlThin;
    object lineColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

    // Add the border to the top edge of the range
    Border topBorder = range.Borders[XlBordersIndex.xlEdgeTop];
    topBorder.LineStyle = lineStyle;
    topBorder.Weight = lineWeight;
    topBorder.Color = lineColor;

    // Add the border to the bottom edge of the range
    Border bottomBorder = range.Borders[XlBordersIndex.xlEdgeBottom];
    bottomBorder.LineStyle = lineStyle;
    bottomBorder.Weight = lineWeight;
    bottomBorder.Color = lineColor;

    // Add the border to the left edge of the range
    Border leftBorder = range.Borders[XlBordersIndex.xlEdgeLeft];
    leftBorder.LineStyle = lineStyle;
    leftBorder.Weight = lineWeight;
    leftBorder.Color = lineColor;

    // Add the border to the right edge of the range
    Border rightBorder = range.Borders[XlBordersIndex.xlEdgeRight];
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
static void Log2DArray(int[,] array)
{
    for (int i = 0; i < array.GetLength(0); i++)
    {
        for (int j = 0; j < array.GetLength(1); j++)
        {
            Console.Write($"{array[i, j]} ");
        }
        Console.WriteLine();
    }
}
static void LogResult(List<List<(int,int)>> results,int size)
{
    if (results != null)
    {
        List<(int, int)> tmp = results[0];
        Console.Write("[");
        for (int i = 0; i < size; i++)
        {
            Console.Write(tmp[i].Item2);
            if (i != size - 1)
            {
                Console.Write(",");
            }
        }
        Console.Write("]");
    }
}
static void cleanCOM()
{
    do
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
    while (Marshal.AreComObjectsAvailableForCleanup());
}