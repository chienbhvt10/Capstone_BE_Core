using ATTAS_API.Models;
using ATTAS_API.Utils;
using ATTAS_CORE;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;
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
            //if (connector.validToken(data.token))
            //{
            Thread solver = new Thread(new ParameterizedThreadStart(Solve));
            data.sessionHash = SessionStringGenerator.Generate(32);
            solver.Start(data);
            var hash = new { sessionId = data.sessionHash };
            return Ok(hash);
            //}
            //else
            //{
            //    var message = new { message = "Invalid Token" };
            //    return BadRequest(message);
            //}
        }

        [HttpPost("get")]
        public IActionResult get([FromBody] GetData data)
        {
            SqlServerConnector connector = new SqlServerConnector("SAKURA", "attas", "sa", "12345678");
            //if (connector.validToken(data.token))
            //{
            Result result = new Result();
            Session session = connector.getSession(data.sessionHash);
            if (session != null)
            {
                result.status = session.statusId;
                result.numberofsolution = session.solutionCount;
            }
            else
            {
                var message = new { message = "Invalid Session Hash" };
                return BadRequest("message");
            }
            Solution solution = connector.getSolution(session.id, data.solutionNo);
            if (solution != null)
            {
                result.taskAssigned = solution.taskAssigned;
                result.workingDay = solution.workingDay;
                result.workingTime = solution.workingTime;
                result.waitingTime = solution.waitingTime;
                result.subjectDiversity = solution.subjectDiversity;
                result.quotaAvailable = solution.quotaAvailable;
                result.walkingDistance = solution.walkingDistance;
                result.subjectPreference = solution.subjectPreference;
                result.slotPreference = solution.slotPreference;
                result.results = connector.getResult(solution.Id, session.id);
            }
            var json = JsonSerializer.Serialize(result);
            return Ok(json);
            //}
            //else
            //{
            //    var message = new { message = "Invalid Token" };
            //    return BadRequest(message);
            //}
        }
        static void Solve(object data)
        {
            Data _data = (Data)data;
            SqlServerConnector connector = new SqlServerConnector("SAKURA", "attas", "sa", "12345678");
            int sessionId = connector.addSession(_data.sessionHash);
            foreach (Task task in _data.tasks)
            {
                connector.addTask(sessionId, task.Id, task.Order);
            }
            foreach (Instructor instructor in _data.instructors)
            {
                connector.addInstructor(sessionId, instructor.Id, instructor.Order);
            }
            foreach (Slot slot in _data.slots)
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
            attas.numDays = _data.numDays;
            attas.numTimes = _data.numTimes;
            attas.numSegments = _data.numSegments;
            int numSegmentRule = _data.numSegmentRules;
            attas.numSubjects = _data.numSubjects;
            attas.numAreas = _data.numAreas;
            attas.numBackupInstructors = _data.backupInstructor;
            attas.slotConflict = listToArray(_data.slotConflict);
            attas.slotDay = listToArray(_data.slotDay);
            attas.slotTime = listToArray(_data.slotTime);
            attas.slotSegment = new int[attas.numSlots, attas.numDays, attas.numSegments];
            for (int i = 0; i < numSegmentRule; i++)
            {
                attas.slotSegment[_data.slotSegment[i][0], _data.slotSegment[i][1], _data.slotSegment[i][2]] = 1;
            }
            attas.patternCost = _data.patternCost.ToArray();
            attas.instructorSubjectPreference = listToArray(_data.instructorSubject);
            attas.instructorSubject = toBinaryArray(attas.instructorSubjectPreference);
            attas.instructorSlotPreference = listToArray(_data.instructorSlot);
            attas.instructorSlot = toBinaryArray(attas.instructorSlotPreference);
            attas.instructorQuota = _data.instructorQuota.ToArray();
            attas.instructorMinQuota = _data.instructorMinQuota.ToArray();
            attas.instructorPreassign = new List<(int, int, int)>();
            if (_data.preassigns != null)
            {
                foreach (Preassign item in _data.preassigns)
                {
                    attas.instructorPreassign.Add((item.instructorOrder, item.taskOrder, 1));
                }
            }
            attas.areaDistance = listToArray(_data.areaDistance);
            attas.areaSlotCoefficient = listToArray(_data.areaSlotCoefficient);
            attas.taskSubjectMapping = new int[attas.numTasks];
            attas.taskSlotMapping = new int[attas.numTasks];
            attas.taskAreaMapping = new int[attas.numTasks];
            foreach (Task item in _data.tasks)
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
            if (results != null)
            {
                connector.updateSessionStatus(sessionId, 4, results.Count);
                int no = 1;
                foreach (var result in results)
                {
                    var sorted = result.OrderBy(t => t.Item2);
                    int currentId = -1;
                    int objQuota = 0;
                    int[] objDay = new int[attas.numDays];
                    int[,] objTime = new int[attas.numDays, attas.numTimes];
                    int objWaiting = 0;
                    int[] objSubjectDiversity = new int[attas.numSubjects];
                    int objQuotaAvailable = 0;
                    int objWalkingDistance = 0;
                    int objSubjectPreference = 0;
                    int objSlotPreference = 0;

                    int finalQuota = 0;
                    int finalDay = 0;
                    int finalTime = 0;
                    int finalWaiting = 0;
                    int finalSubjectDiversity = 0;
                    int finalQuotaAvailable = 0;
                    int finalWalkingDistance = 0;
                    int finalSubjectPreference = 0;
                    int finalSlotPreference = 0;

                    List<int> tasks = new List<int>();
                    foreach (var item in sorted)
                    {
                        if (currentId != item.Item2)
                        {
                            if (currentId != -1)
                            {
                                finalQuota += objQuota;
                                finalDay += objDay.Sum();
                                finalTime += flattenArray(objTime).Sum();
                                finalWaiting += calObjWaitingTime(tasks, attas);
                                finalSubjectDiversity = Math.Max(finalSubjectDiversity, objSubjectDiversity.Sum());
                                finalQuotaAvailable = Math.Max(finalQuotaAvailable, attas.instructorQuota[currentId] - objQuota);
                                finalWalkingDistance += calObjWalkingDistance(tasks, attas);
                                finalSubjectPreference += objSubjectPreference;
                                finalSlotPreference += objSlotPreference;
                            }
                            //reset
                            objQuota = 0;
                            Array.Clear(objDay, 0, objDay.Length);
                            Array.Clear(objTime, 0, objTime.Length);
                            objWaiting = 0;
                            Array.Clear(objSubjectDiversity, 0, objSubjectDiversity.Length);
                            objQuotaAvailable = 0;
                            objWalkingDistance = 0;
                            objSubjectPreference = 0;
                            objSlotPreference = 0;
                            tasks.Clear();
                            currentId = item.Item2;
                        }
                        if (currentId != -1)
                        {
                            tasks.Add(item.Item1);
                            int thisTaskSlot = attas.taskSlotMapping[item.Item1];
                            int thisTaskSubject = attas.taskSubjectMapping[item.Item1];
                            objQuota += 1;
                            for (int d = 0; d < attas.numDays; d++)
                            {
                                if (attas.slotDay[thisTaskSlot, d] == 1)
                                {
                                    objDay[d] = 1;
                                    for (int t = 0; t < attas.numTimes; t++)
                                    {
                                        if (attas.slotTime[thisTaskSlot, t] == 1)
                                            objTime[d, t] = 1;
                                    }
                                }
                            }
                            objSubjectDiversity[thisTaskSubject] = 1;
                            objSubjectPreference += attas.instructorSubjectPreference[item.Item2, thisTaskSubject];
                            objSlotPreference += attas.instructorSlotPreference[item.Item2, thisTaskSlot];
                        }
                    }
                    if (currentId != -1)
                    {
                        finalQuota += objQuota;
                        finalDay += objDay.Sum();
                        finalTime += flattenArray(objTime).Sum();
                        finalWaiting += calObjWaitingTime(tasks, attas);
                        finalSubjectDiversity = Math.Max(finalSubjectDiversity, objSubjectDiversity.Sum());
                        finalQuotaAvailable = Math.Max(finalQuotaAvailable, attas.instructorQuota[currentId] - objQuota);
                        finalWalkingDistance += calObjWalkingDistance(tasks, attas);
                        finalSubjectPreference += objSubjectPreference;
                        finalSlotPreference += objSlotPreference;
                    }

                    int solutionId = connector.addSolution(sessionId, no, finalQuota, finalDay, finalTime, finalWaiting, finalSubjectDiversity, finalQuotaAvailable, finalWalkingDistance, finalSubjectPreference, finalSlotPreference);
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
        /*
        ################################
        ||    CALCULATE OBJECTIVE     ||
        ################################
        */
        static int calObjWalkingDistance(List<int> tasks, ATTAS_ORTOOLS attas)
        {
            int distance = 0;
            int n = tasks.Count();
            for (int i = 0; i < n - 1; i++)
                for (int j = i + 1; j < n; j++)
                {
                    int t1 = tasks[i];
                    int t2 = tasks[j];
                    distance += attas.areaSlotCoefficient[attas.taskSlotMapping[t1], attas.taskSlotMapping[t2]] * attas.areaDistance[attas.taskAreaMapping[t1], attas.taskAreaMapping[t2]];
                }
            return distance;
        }
        static int calObjWaitingTime(List<int> tasks, ATTAS_ORTOOLS attas)
        {
            int result = 0;
            int[,] flag = new int[attas.numDays, attas.numSegments];
            foreach (int task in tasks)
            {
                int slot = attas.taskSlotMapping[task];
                foreach (int d in attas.allDays)
                    foreach (int s in attas.allSegments)
                        if (attas.slotSegment[slot, d, s] == 1)
                            flag[d, s] = 1;
            }
            foreach (int d in attas.allDays)
            {
                int pattern = 0;
                foreach (int s in attas.allSegments)
                    pattern += flag[d, s] * (1 << (attas.numSegments - s - 1));
                result += attas.patternCost[pattern];
            }
            return result;
        }
        /*
        ################################
        ||          UTILITY           ||
        ################################
        */
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