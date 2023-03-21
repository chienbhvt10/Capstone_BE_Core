using ATTAS_API.Models;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;
using ATTAS_API.Utils;
using ATTAS_CORE;
using Task = ATTAS_API.Models.Task;

namespace ATTAS_API.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ATTASAPIController : ControllerBase
    {
        private readonly ILogger<ATTASAPIController> _logger;

        public ATTASAPIController(ILogger<ATTASAPIController> logger)
        {
            _logger = logger;
        }

        [HttpPost("excecute")]
        public IActionResult excecute([FromBody] Data data)
        {
            SqlServerConnector connector = new SqlServerConnector("SAKURA", "attas", "sa", "12345678");
            if( connector.validToken(data.token) ) { 
                Thread solver = new Thread(new ParameterizedThreadStart(Solve));
                data.sessionHash = SessionStringGenerator.Generate(32);
                solver.Start(data);
                var hash = new { sessionId = data.sessionHash };
                return Ok(hash);
            }
            else
            {
                var message = new { message = "Invalid Token" };
                return BadRequest(message);
            }
        }

        [HttpPost("get")]
        public IActionResult get()
        {
            Result result = new Result();
            var json = JsonSerializer.Serialize(result);
            return Ok(json);
        }

        static void Solve(object data)
        {
            Data _data = (Data)data;
            SqlServerConnector connector = new SqlServerConnector("SAKURA","attas","sa","12345678");
            int sessionId = connector.addSession(_data.sessionHash);
            foreach(Task task in _data.tasks)
            {
                connector.addTask(sessionId, task.Id, task.Order);
            }
            foreach(Instructor instructor in _data.instructors)
            {
                connector.addInstructor(sessionId, instructor.Id, instructor.Order);
            }
            foreach(Slot slot in _data.slots)
            {
                connector.addTime(sessionId, slot.Id, slot.Order);
            }
            ATTAS_ORTOOLS attas = new ATTAS_ORTOOLS();
            //SETTING
            attas.maxSearchingTimeOption = _data.Setting.maxSearchingTime;
            attas.strategyOption = _data.Setting.strategy;
            attas.objOption = _data.Setting.objectiveOption.ToArray();
            attas.objWeight = _data.Setting.objectiveWeight.ToArray();
            attas.debugLoggerOption = true;
            //INPUT
            if (attas.debugLoggerOption)
            {
                Console.WriteLine("DEBUG: Reading Data");
            }
            attas.numTasks = _data.numTasks;
            attas.numInstructors = _data.numInstructors;
            attas.numSlots = _data.numSlots;
            attas.numSubjects = _data.numSubjects;
            attas.numAreas = _data.numAreas;
            attas.numBackupInstructors = _data.backupInstructor;
            attas.slotConflict = listToArray(_data.slotConflict);
            attas.slotCompatibilityCost = listToArray(_data.slotCompability);
            attas.instructorSubjectPreference = listToArray(_data.instructorSubject);
            attas.instructorSubject = toBinaryArray(attas.instructorSubjectPreference);
            attas.instructorSlotPreference = listToArray(_data.instructorSlot);
            attas.instructorSlot = toBinaryArray(attas.instructorSlotPreference);
            attas.instructorQuota = _data.instructorQuota.ToArray();
            attas.instructorPreassign = new List<(int, int, int)>();
            if (_data.preassigns != null)
            {
                foreach (Preassign item in _data.preassigns)
                {
                    attas.instructorPreassign.Add((item.instructorOrder, item.taskOrder, 1));
                }
            }
            attas.areaDistance = listToArray(_data.areaDistance);
            attas.areaSlotWeight = listToArray(_data.areaSlotCoefficient);
            attas.taskSubjectMapping = new int[attas.numTasks];
            attas.taskSlotMapping = new int[attas.numTasks];
            attas.taskAreaMapping = new int[attas.numTasks];
            foreach(Task item in _data.tasks)
            {
                attas.taskSubjectMapping[item.Order] = item.subjectOrder;
                attas.taskSlotMapping[item.Order] = item.slotOrder;
                attas.taskAreaMapping[item.Order] = item.areaOrder;
            }
            if (attas.debugLoggerOption)
            {
                Console.WriteLine("DEBUG: Start Solving");
            }
            List<List<(int, int)>>? results = attas.solve();
            if (results!= null)
            {
                connector.updateSessionStatus(sessionId, 4, results.Count);
                int no = 1;
                foreach (var result in results)
                {
                    int taskAssigned = 0;
                    int slotCompability = 0;
                    int subjectDiversity = 0;
                    int quotaAvailable = 0;
                    int walkingDistance = 0;
                    int subjectPreference = 0;
                    int slotPreference = 0;
                    for (int i = 0; i < attas.numTasks; i++)
                    {
                        if (result[i].Item2 != -1)
                        {
                            taskAssigned++;
                            subjectPreference += attas.instructorSubjectPreference[result[i].Item2, attas.taskSubjectMapping[i]];
                            slotPreference += attas.instructorSlotPreference[result[i].Item2, attas.taskSlotMapping[i]];
                        }
                    }
                    List<List<int>> grouped = new List<List<int>>();
                    for (int i = 0; i < attas.numInstructors; i++)
                    {
                        grouped.Add(new List<int>());
                    }
                    
                    foreach (var item in result)
                    {
                        if (item.Item2 != -1)
                        {
                            grouped[item.Item2].Add(item.Item1);
                        }
                    }
                    for (int idx = 0; idx < attas.numInstructors; idx++)
                    {
                        int n = grouped[idx].Count;
                        subjectDiversity = Math.Max(subjectDiversity,(from x in grouped[idx] select attas.taskSubjectMapping[x]).Distinct().Count());
                        quotaAvailable = Math.Max(quotaAvailable, attas.instructorQuota[idx] - n);
                        for (int i = 0; i < n - 1; i++)
                        {
                            for (int j = i + 1; j < n; j++)
                            {
                                slotCompability += attas.slotCompatibilityCost[attas.taskSlotMapping[grouped[idx][i]], attas.taskSlotMapping[grouped[idx][j]]];
                                walkingDistance += attas.areaSlotWeight[attas.taskSlotMapping[grouped[idx][i]], attas.taskSlotMapping[grouped[idx][j]]] * attas.areaDistance[attas.taskAreaMapping[grouped[idx][i]], attas.taskAreaMapping[grouped[idx][j]]];
                            }
                        }
                    }
                    int solutionId = connector.addSolution(sessionId, no, taskAssigned, slotCompability, subjectDiversity, quotaAvailable, walkingDistance, subjectPreference, slotPreference);
                    foreach (var item in result)
                    {
                        connector.addResult(solutionId, item.Item1, item.Item2, attas.taskSlotMapping[item.Item1]);
                    }
                    if (attas.debugLoggerOption)
                    {
                        Console.WriteLine($"Solution {no} :");
                        List<int> tmp = new List<int>();
                        foreach (var item in result)
                        {
                            tmp.Add(item.Item2);
                        }
                        string formattedResults = "[" + string.Join(",", tmp) + "]";
                        Console.WriteLine(formattedResults);
                    }
                    no++;
                } 
            }
            else
            {
                connector.updateSessionStatus(sessionId, 3, 0);
            }
        }
        static int[,] listToArray(List<List<int>> list)
        {
            int rows = list.Count;
            int cols = list[0].Count;

            int[,] array = new int[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    array[i, j] = list[i][j];
                }
            }
            return array;
        }
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
        static void printArray(int[,] myArray)
        {
            int rows = myArray.GetLength(0);
            int columns = myArray.GetLength(1);

            // Loop through each row and column, printing out the value at each position
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    Console.Write(myArray[i, j] + " ");
                }
                Console.WriteLine();
            }
        }
    }
}