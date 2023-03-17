using ATTAS_API.Models;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;
using ATTAS_CORE;
using Google.Protobuf.WellKnownTypes;
using Google.OrTools.Sat;
using System;

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
            Thread t = new Thread(new ParameterizedThreadStart(Solve));
            t.Start(data);
            var hash = new { requestID = "requestID" };
            return Ok(hash);
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
            foreach(Models.Task item in _data.tasks)
            {
                attas.taskSubjectMapping[item.Order] = item.subjectOrder;
                attas.taskSlotMapping[item.Order] = item.slotOrder;
                attas.taskAreaMapping[item.Order] = item.areaOrder;
            }
            if (attas.debugLoggerOption)
            {
                Console.WriteLine("DEBUG: Start Solving");
            }
            List<(int, int)>? results = attas.solve();
            if (results!= null)
            {
                if (attas.debugLoggerOption)
                {
                    List<int> tmp = new List<int>();
                    foreach (var item in results)
                    {
                        tmp.Add(item.Item2);
                    }
                    string formattedResults = "[" + string.Join(",", tmp) + "]";
                    Console.WriteLine(formattedResults);
                }
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