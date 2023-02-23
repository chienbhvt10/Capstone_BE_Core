using ATTAS_CORE;
using Excel = Microsoft.Office.Interop.Excel;
ATTAS attas = new ATTAS();

static int[,] excelToArray(Excel._Worksheet oSheet, int startRow,int startCol,int numRows,int numCols)
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
static void printArray(int[,] data)
{
    for (int i = 0; i < data.GetLength(0); i++)
    {
        for (int j = 0; j < data.GetLength(1); j++)
        {
            Console.Write(data[i, j] + " ");
        }
        Console.WriteLine();
    }
}
attas.objOption = new int[6] { 0, 0, 0, 0, 0, 1 };
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

try
{
    Excel.Application oXL;
    Excel._Workbook oWB;

    //Start Excel and get Application object.
    oXL = new Excel.Application();
    //oXL.Visible = true;
    //Get a new workbook.
    oWB = oXL.Workbooks.Open(@"D:\FPT\SEP490_G14\input.xlsx");

    attas.slotConflict = excelToArray((Excel._Worksheet)oWB.Sheets[2], 2, 2, attas.numSlots, attas.numSlots);
    attas.slotCompatibilityCost = excelToArray((Excel._Worksheet)oWB.Sheets[3], 2, 2, attas.numSlots, attas.numSlots);
    attas.instructorSubjectPreference = excelToArray((Excel._Worksheet)oWB.Sheets[4], 2, 2, attas.numInstructors, attas.numSubjects);
    attas.instructorSubject = new int[attas.numInstructors, attas.numSubjects];
    for (int i = 0; i < attas.numInstructors; i++)
        for (int j = 0; j < attas.numSubjects; j++)
            if (attas.instructorSubjectPreference[i, j] > 0)
            { 
                attas.instructorSubject[i, j] = 1;
            }
            else
            {
                attas.instructorSubject[i, j] = 0;
            }
    
    attas.instructorSlotPreference = excelToArray((Excel._Worksheet)oWB.Sheets[5], 2, 2, attas.numInstructors, attas.numSlots);
    attas.instructorSlot = new int[attas.numInstructors, attas.numSlots];
    for (int i = 0; i < attas.numInstructors; i++)
        for (int j = 0; j < attas.numSlots; j++)
            if (attas.instructorSlotPreference[i, j] > 0)
            {
                attas.instructorSlot[i, j] = 1;
            }
            else
            {
                attas.instructorSlot[i, j] = 0;
            }

    attas.instructorQuota = flattenArray(excelToArray((Excel._Worksheet)oWB.Sheets[6], 2, 2, attas.numInstructors, 1));
    attas.areaDistance = excelToArray((Excel._Worksheet)oWB.Sheets[8], 2, 2, attas.numAreas, attas.numAreas);
    attas.areaSlotWeight = excelToArray((Excel._Worksheet)oWB.Sheets[9], 2, 2, attas.numSlots, attas.numSlots);

    attas.taskSubjectMapping = new int[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 2, 3, 4, 5, 6, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 9, 9, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 11, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13 };
    attas.taskSlotMapping = new int[] { 9, 11, 16, 19, 16, 6, 12, 14, 1, 3, 9, 11, 17, 19, 4, 6, 12, 14, 1, 3, 9, 11, 17, 19, 4, 6, 12, 14, 1, 3, 9, 11, 17, 19, 4, 6, 12, 14, 1, 3, 9, 17, 11, 4, 19, 12, 6, 1, 14, 0, 3, 17, 11, 4, 19, 6, 12, 14, 1, 3, 9, 11, 19, 17, 4, 8, 6, 16, 5, 14, 13, 3, 2, 7, 15, 3, 14, 3, 17, 11, 4, 19, 12, 6, 1, 14, 9, 15, 16, 1, 14, 15, 10, 10, 5, 11, 4, 19, 12, 6, 0, 14, 9, 3, 17, 3, 9, 8, 2, 14, 15, 15, 17, 2, 5, 10, 13, 18, 0, 7, 8, 7, 8, 19, 16, 13, 13, 0, 7, 8, 10, 16, 2, 5, 10, 13, 18, 0, 18, 0, 7, 8, 5, 5, 13, 6, 9, 11, 17, 3, 4, 11, 12, 19, 1, 19, 1, 6, 9, 4, 4, 12 };
    attas.taskAreaMapping = new int[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 2, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 1, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 1, 1, 1, 1, 2, 1, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 1, 2, 2, 1, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 1, 1, 2, 1, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2 };
    //attas.instructorPreassign = new List<(int, int, int)> { (32, 0, 1), (32, 1, 1), (32, 2, 1) };
    oWB.Close();
    oXL.Quit();

    List<(int, int)>? results = attas.solve();
    if (results != null)
        foreach ((int, int) result in results)
            if (result.Item2 >= 0)
                Console.WriteLine($"Tasks {result.Item1} assigned to instructor {result.Item2}");
            else
                Console.WriteLine($"Tasks {result.Item1} need backup instructor!");
    else
        Console.WriteLine("No Solution");
}
catch (Exception e)
{
    Console.WriteLine(e);
}




