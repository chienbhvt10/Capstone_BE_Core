using Google.OrTools.Sat;

namespace ATTAS_CORE
{
    /*
    ################################
    ||       START OR-TOOLS       ||
    ################################
    */
    public class ATTAS_ORTOOLS
    {
        /*
        ################################
        ||           Option           ||
        ################################
        */
        public double maxSearchingTimeOption { get; set; } = 300.0;
        public int strategyOption { get; set; } = 2;
        public int[] objOption { get; set; } = new int[8] { 1, 1, 0, 0, 0, 0, 0, 0 };
        public int[] objWeight { get; set; } = new int[8] { 50 ,25, 1, 1, 1, 1, 1, 1 };
        public bool debugLoggerOption { get; set; } = false;

        /*
        ################################
        ||      SOLVER PARAMETER      ||
        ################################
        */

        //COUNT
        public int numSubjects { get; set; } = 0;
        public int numTasks { get; set; } = 0;
        public int numSlots { get; set; } = 0;
        public int numDays { get; set; } = 0;
        public int numTimes { get; set; } = 0;
        public int numSegments { get; set; } = 0;
        public int numInstructors { get; set; } = 0;
        public int numBackupInstructors { get; set; } = 0;
        public int numAreas { get; set; } = 0;
        //RANGE
        private int[] allSubjects = Array.Empty<int>();
        private int[] allTasks = Array.Empty<int>();
        private int[] allSlots = Array.Empty<int>();
        private int[] allDays = Array.Empty<int>();
        private int[] allTimes = Array.Empty<int>();
        private int[] allSegments = Array.Empty<int>();
        private int[] allInstructors = Array.Empty<int>();
        private int[] allInstructorsWithBackup = Array.Empty<int>();
        //INPUT DATA
        public int[,] slotConflict { get; set; } = new int[0, 0];
        public int[,] slotDay { get; set; } = new int[0, 0];
        public int[,] slotTime { get; set; } = new int[0, 0];
        public int[,,] slotSegment { get; set; } = new int[0, 0 ,0];
        public int[] patternCost { get; set; } = Array.Empty<int>();
        public int[,] instructorSubject { get; set; } = new int[0, 0];
        public int[,] instructorSubjectPreference { get; set; } = new int[0, 0];
        public int[,] instructorSlot { get; set; } = new int[0, 0];
        public int[,] instructorSlotPreference { get; set; } = new int[0, 0];
        public List<(int, int, int)> instructorPreassign { get; set; } = new List<(int, int, int)>();
        public int[] instructorQuota { get; set; } = Array.Empty<int>();
        public int[] instructorMinQuota { get; set; } = Array.Empty<int>();
        public int[] taskSubjectMapping { get; set; } = Array.Empty<int>();
        public int[] taskSlotMapping { get; set; } = Array.Empty<int>();
        public int[] taskAreaMapping { get; set; } = Array.Empty<int>();
        public int[,] areaDistance { get; set; } = new int[0, 0];
        public int[,] areaSlotCoefficient { get; set; } = new int[0, 0];  

        /*
        ################################
        ||           MODEL            ||
        ################################
         */

        private CpModel model;
        CpSolver solver;
        CpSolverStatus status;
        // Desicion variable
        private Dictionary<(int, int), BoolVar> assigns;
        private Dictionary<(int, int), BoolVar> instructorDayStatus;
        private Dictionary<(int, int, int), BoolVar> instructorTimeStatus;
        private Dictionary<(int, int), BoolVar> instructorSubjectStatus;
        private Dictionary<(int, int, int), BoolVar> instructorSegmentStatus;
        private Dictionary<(int, int, int), BoolVar> instructorPatternStatus;
        private Dictionary<(int, int), LinearExpr> assignsProduct;
        public void setSolverCount()
        {
            allSubjects = Enumerable.Range(0, numSubjects).ToArray();
            allTasks = Enumerable.Range(0, numTasks).ToArray();
            allSlots = Enumerable.Range(0, numSlots).ToArray();
            allDays = Enumerable.Range(0, numDays).ToArray();
            allTimes = Enumerable.Range(0, numTimes).ToArray();
            allSegments = Enumerable.Range(0, numSegments).ToArray();
            allInstructors = Enumerable.Range(0, numInstructors).ToArray();
            if (numBackupInstructors > 0)
            {
                allInstructorsWithBackup = Enumerable.Range(0, numInstructors + 1).ToArray();
                instructorQuota = instructorQuota.Concat(new int[] { numBackupInstructors }).ToArray();
                instructorMinQuota = instructorMinQuota.Concat(new int[] { 0 }).ToArray();
            }
            else
            {
                allInstructorsWithBackup = Enumerable.Range(0, numInstructors).ToArray();
            }
        }
        public void createModel()
        {
            model = new CpModel(); 

            assigns = new Dictionary<(int, int), BoolVar>();
            foreach (int n in allTasks)
                foreach (int i in allInstructorsWithBackup)
                {
                    assigns.Add((n, i), model.NewBoolVar($"n{n}i{i}"));
                }
            
            List<ILiteral> literals = new List<ILiteral>();
            //C-00 EACH TASK ASSIGN TO ATLEAST ONE AND ONLY ONE
            foreach (int n in allTasks)
            {
                foreach (int i in allInstructorsWithBackup)
                    literals.Add(assigns[(n, i)]);
                model.AddExactlyOne(literals);
                literals.Clear();
            }
            
            //C-00 CONSTRAINT INSTRUCTOR QUOTA MUST IN RANGE
            List<IntVar> taskAssigned = new List<IntVar>();
            foreach (int i in allInstructorsWithBackup)
            {
                foreach (int n in allTasks)
                    taskAssigned.Add(assigns[(n, i)]);
                model.AddLinearConstraint(LinearExpr.Sum(taskAssigned), instructorMinQuota[i], instructorQuota[i]);
                taskAssigned.Clear();
            }
            
            List<List<int>> task_in_this_slot = new List<List<int>>();
            List<List<int>> task_conflict_with_this_slot = new List<List<int>>();

            foreach (int s in allSlots)
            {
                List<int> sublist_task_in_this_slot = new List<int>();
                List<int> sublist_task_conflict_with_this_slot = new List<int>();

                foreach (int n in allTasks)
                {
                    if (taskSlotMapping[n] == s && slotConflict[s,s] == 1)
                        sublist_task_in_this_slot.Add(n);

                    if (slotConflict[taskSlotMapping[n], s] == 1)
                        sublist_task_conflict_with_this_slot.Add(n);
                }
                task_in_this_slot.Add(sublist_task_in_this_slot);
                task_conflict_with_this_slot.Add(sublist_task_conflict_with_this_slot);
            }
            //C-01 NO SLOT CONFLICT
            List<LinearExpr> taskAssignedThatSlot = new List<LinearExpr>();
            List<LinearExpr> taskAssignedConflictWithThatSlot = new List<LinearExpr>();
            foreach (int i in allInstructors)
                foreach (int s in allSlots)
                {
                    foreach (int n in task_in_this_slot[s])
                        taskAssignedThatSlot.Add(assigns[(n, i)]);
                    foreach (int n in task_conflict_with_this_slot[s])
                        taskAssignedConflictWithThatSlot.Add(assigns[(n, i)]);
                    ILiteral tmp = model.NewBoolVar("");
                    model.Add(LinearExpr.Sum(taskAssignedThatSlot) > 0).OnlyEnforceIf(tmp);
                    model.Add(LinearExpr.Sum(taskAssignedThatSlot) == 0).OnlyEnforceIf(tmp.Not());
                    model.Add(LinearExpr.Sum(taskAssignedConflictWithThatSlot) == 1).OnlyEnforceIf(tmp);
                    taskAssignedThatSlot.Clear();
                    taskAssignedConflictWithThatSlot.Clear();
                }

            //C-02 PREASSIGN MUST BE SATISFY
            foreach (var data in instructorPreassign)
            {
                if (data.Item3 == 1)
                {
                    model.AddHint(assigns[(data.Item2, data.Item1)], 1);
                    model.Add(assigns[(data.Item2, data.Item1)] == 1);
                }
                if (data.Item3 == -1)
                {
                    model.AddHint(assigns[(data.Item2, data.Item1)], 0);
                    model.Add(assigns[(data.Item2, data.Item1)] == 0);
                }
            }
            //C-03 INSTRUCTOR MUST HAVE ABILITY FOR THAT SUBJECT
            foreach (int n in allTasks)
                foreach (int i in allInstructors)
                    model.Add(instructorSubject[i, taskSubjectMapping[n]] - assigns[(n, i)] > -1);
            //C-04 INSTRUCTOR MUST BE ABLE TO TEACH IN THAT SLOT
            foreach (int n in allTasks)
                foreach (int i in allInstructors)
                    model.Add(instructorSlot[i, taskSlotMapping[n]] - assigns[(n, i)] > -1);
            
        }
        public List<List<(int, int)>>? constraintOnly()
        {
            setSolverCount();
            createModel();
            List<ILiteral> obj = new List<ILiteral> ();
            foreach (int n in allTasks)
                foreach (int i in allInstructors)
                    obj.Add(assigns[(n, i)]);
            model.Minimize( createDelta( numTasks ,LinearExpr.Sum(obj),numTasks ));
            solver = new CpSolver();
            solver.StringParameters += "linearization_level:0 " + $"max_time_in_seconds:{maxSearchingTimeOption} ";
            status = solver.Solve(model);

            if (debugLoggerOption)
            {
                Console.WriteLine("Statistics");
                Console.WriteLine($"  {strategyOption}: {solver.ObjectiveValue}");
                Console.WriteLine($"  status: {status}");
                Console.WriteLine($"  conflicts: {solver.NumConflicts()}");
                Console.WriteLine($"  branches : {solver.NumBranches()}");
                Console.WriteLine($"  wall time: {solver.WallTime()}s");
            }
            if (status == CpSolverStatus.Optimal || status == CpSolverStatus.Feasible)
                return getResults(solver);
            else return null;
        }
        public int findBackupInstructor()
        {
            numBackupInstructors = numTasks;
            setSolverCount();
            createModel();
            List<ILiteral> obj = new List<ILiteral>();
            foreach (int n in allTasks)
                foreach (int i in allInstructors)
                    obj.Add(assigns[(n, i)]);
            model.Minimize(createDelta(numTasks, LinearExpr.Sum(obj), numTasks));
            solver = new CpSolver();
            solver.StringParameters += "linearization_level:0 " + $"max_time_in_seconds:{maxSearchingTimeOption} ";
            status = solver.Solve(model);
            if (status == CpSolverStatus.Optimal || status == CpSolverStatus.Feasible)
                return (int)solver.ObjectiveValue;
            else 
                return numTasks;
        }
        /*
        ################################
        ||         OBJECTIVE          ||
        ################################
        */

        // O-01 MINIMIZE DAY
        public LinearExpr objTeachingDay()
        {
            List<ILiteral> teachingDay = new List<ILiteral>();
            foreach (int i in allInstructors)
                foreach (int d in allDays)
                    if(instructorDayStatus.TryGetValue((i, d),out BoolVar value) )
                        teachingDay.Add(value);
            return LinearExpr.Sum(teachingDay);
        }
        // O-02 MINIMIZE TIME
        public LinearExpr objTeachingTime()
        {
            List<ILiteral> teachingTime = new List<ILiteral>();
            foreach (int i in allInstructors)
                foreach (int d in allDays)
                    foreach (int t in allTimes)
                        if (instructorTimeStatus.TryGetValue((i, d, t), out BoolVar value))
                            teachingTime.Add(value);
            return LinearExpr.Sum(teachingTime);
        }
        // O-03 MINIMIZE SEGMENT COST
        public LinearExpr objPatternCost()
        {
            List<LinearExpr> allPatternCost = new List<LinearExpr>();
            foreach(int i  in allInstructors)
                foreach(int d in allDays)
                    for (int p = 0; p < (1 << numSegments); p++)
                        allPatternCost.Add(patternCost[p] * instructorPatternStatus[(i,d,p)]);
            return LinearExpr.Sum(allPatternCost);
        }
        // O-04 MINIMIZE SUBJECT DIVERSITY
        public LinearExpr objSubjectDiversity()
        {
            List<ILiteral> literals = new List<ILiteral>();
            List<LinearExpr> subjectDiversity = new List<LinearExpr>();
            foreach (int i in allInstructors)
            {
                foreach (int s in allSubjects)
                    literals.Add(instructorSubjectStatus[(i, s)]);
                subjectDiversity.Add(LinearExpr.Sum(literals));
                literals.Clear();
            }
            IntVar obj = model.NewIntVar(0, numSubjects, "subjectDiversity");
            model.AddMaxEquality(obj, subjectDiversity);
            return obj;
        }
        // O-05 MINIMIZE QUOTA DIFF
        public LinearExpr objQuotaReached()
        {
            List<LinearExpr> quotaDifference = new List<LinearExpr>();
            foreach (int i in allInstructors)
            {
                IntVar[] x = new IntVar[numTasks];
                foreach (int n in allTasks)
                    x[n] = assigns[(n, i)];
                quotaDifference.Add(instructorQuota[i] - LinearExpr.Sum(x));
            }
            IntVar obj = model.NewIntVar(0, numTasks, "maxQuotaDifference");
            model.AddMaxEquality(obj, quotaDifference);
            return obj;
        }
        // O-06 MINIMIZE WALKING DISTANCE
        public LinearExpr objWalkingDistance()
        {
            List<LinearExpr> walkingDistance = new List<LinearExpr>();
            for (int n1 = 0; n1 < numTasks - 1; n1++)
                for (int n2 = n1 + 1; n2 < numTasks; n2++)
                {
                    if (areaSlotCoefficient[taskSlotMapping[n1], taskSlotMapping[n2]] == 0 || areaDistance[taskAreaMapping[n1], taskAreaMapping[n2]] == 0)
                        continue;
                    walkingDistance.Add(assignsProduct[(n1, n2)] * areaSlotCoefficient[taskSlotMapping[n1], taskSlotMapping[n2]] * areaDistance[taskAreaMapping[n1], taskAreaMapping[n2]]);
                }
            return LinearExpr.Sum(walkingDistance);
        }
        // O-07 MAXIMIZE SUBJECT PREFERENCE
        public LinearExpr objSubjectPreference()
        {
            IntVar[] assignedTasks = new IntVar[numTasks * numInstructors];
            int[] assignedTaskSubjectPreferences = new int[numTasks * numInstructors];
            foreach (int n in allTasks)
            {
                foreach (int i in allInstructors)
                {
                    assignedTasks[n * numInstructors + i] = assigns[(n, i)];
                    assignedTaskSubjectPreferences[n * numInstructors + i] = instructorSubjectPreference[i, taskSubjectMapping[n]];
                }
            }
            return LinearExpr.WeightedSum(assignedTasks, assignedTaskSubjectPreferences);
        }
        // O-08 MAXIMIZE SLOT PREFERENCE
        public LinearExpr objSlotPreference()
        {
            IntVar[] assignedTasks = new IntVar[numTasks * numInstructors];
            int[] assignedTaskSlotPreferences = new int[numTasks * numInstructors];
            foreach (int n in allTasks)
            {
                foreach (int i in allInstructors)
                {
                    assignedTasks[n * numInstructors + i] = assigns[(n, i)];
                    assignedTaskSlotPreferences[n * numInstructors + i] = instructorSlotPreference[i, taskSlotMapping[n]];
                }
            }
            return LinearExpr.WeightedSum(assignedTasks, assignedTaskSlotPreferences);
        }
        public List<List<(int, int)>>? objectiveOptimize()
        {
            setSolverCount();
            // ( numTasks * numInstructors )
            createModel();
            solver = new CpSolver();
            status = new CpSolverStatus();
            solver.StringParameters += "linearization_level:0 " + $"max_time_in_seconds:{maxSearchingTimeOption} ";
            List<LinearExpr> totalDeltas = new List<LinearExpr>();
            // O-01 MINIMIZE DAY ( numInstructors * numDays )
            if (objOption[0] > 0)
            {
                List<ILiteral> literals = new List<ILiteral>();
                instructorDayStatus = new Dictionary<(int, int), BoolVar>();
                foreach (int i in allInstructors)
                    foreach (int d in allDays)
                    {
                        foreach (int n in allTasks)
                            if (instructorSlot[i, taskSlotMapping[n]] == 1 && instructorSlot[i, taskSlotMapping[n]] == 1 && slotDay[taskSlotMapping[n], d] == 1)
                                literals.Add(assigns[(n, i)]);
                        if (literals.Count() != 0)
                        {
                            instructorDayStatus.Add((i, d), model.NewBoolVar($"i{i}d{d}"));
                            model.Add(LinearExpr.Sum(literals) > 0).OnlyEnforceIf(instructorDayStatus[(i, d)]);
                            model.Add(LinearExpr.Sum(literals) == 0).OnlyEnforceIf(instructorDayStatus[(i, d)].Not());
                        }
                        literals.Clear();
                    }

                switch (strategyOption)
                {
                    case 1:
                        totalDeltas.Add(objWeight[0] * objTeachingDay());
                        break;
                    case 2:
                        totalDeltas.Add(objWeight[0] * createDelta(numDays * numInstructors, objTeachingDay(), 0));
                        break;
                    case 3:
                        totalDeltas.Add(objWeight[0] * createSquare(objTeachingDay(), 0));
                        break;
                }
            }

            // O-02 MINIMIZE TIME ( numInstructors * numDays * numTimes )
            if (objOption[1] > 0)
            {
                List<ILiteral> literals = new List<ILiteral>();
                instructorTimeStatus = new Dictionary<(int, int, int), BoolVar>();
                foreach (int i in allInstructors)
                    foreach (int d in allDays)
                        foreach (int t in allTimes)
                        {
                            foreach (int n in allTasks)
                                if (instructorSlot[i, taskSlotMapping[n]] == 1 && instructorSlot[i, taskSlotMapping[n]] == 1 && slotDay[taskSlotMapping[n], d] == 1 && slotTime[taskSlotMapping[n], t] == 1)
                                    literals.Add(assigns[(n, i)]);
                            if (literals.Count() != 0)
                            {
                                instructorTimeStatus.Add((i, d, t), model.NewBoolVar($"i{i}d{d}s{t}"));
                                model.Add(LinearExpr.Sum(literals) > 0).OnlyEnforceIf(instructorTimeStatus[(i, d, t)]);
                                model.Add(LinearExpr.Sum(literals) == 0).OnlyEnforceIf(instructorTimeStatus[(i, d, t)].Not());
                            }
                            literals.Clear();
                        }
                switch (strategyOption)
                {
                    case 1:
                        totalDeltas.Add(objWeight[1] * objTeachingTime());
                        break;
                    case 2:
                        totalDeltas.Add(objWeight[1] * createDelta(numTimes * numDays * numInstructors, objTeachingTime(), 0));
                        break;
                    case 3:
                        totalDeltas.Add(objWeight[1] * createSquare(objTeachingTime(), 0));
                        break;
                }
            }

            // O-03 MINIMIZE PATTERN COST ( numInstructors * numDays * ( numSegments + 2^num Segments )
            if (objOption[2] > 0) {
                List<ILiteral> literals = new List<ILiteral>();
                instructorSegmentStatus = new Dictionary<(int, int, int), BoolVar> ();
                foreach(int i in allInstructors)
                    foreach(int d in allDays)
                        foreach(int sm in allSegments)
                        {
                            foreach(int n in allTasks)
                                if(instructorSlot[i, taskSlotMapping[n]] == 1 && instructorSlot[i, taskSlotMapping[n]] == 1 && slotSegment[taskSlotMapping[n], d , sm] == 1)
                                    literals.Add(assigns[(n, i)]);
                            instructorSegmentStatus.Add((i,d,sm), model.NewBoolVar($"i{i}d{d}sm{sm}"));
                            if (literals.Count() == 0)
                                model.AddHint(instructorSegmentStatus[(i, d, sm)], 0);
                            model.Add(LinearExpr.Sum(literals) > 0).OnlyEnforceIf(instructorSegmentStatus[(i, d, sm)]);
                            model.Add(LinearExpr.Sum(literals) == 0).OnlyEnforceIf(instructorSegmentStatus[(i, d, sm)].Not());
                            literals.Clear();
                        }
                instructorPatternStatus = new Dictionary<(int, int, int), BoolVar>();
                foreach (int i in allInstructors)
                    foreach (int d in allDays)
                        for (int p = 0; p < (1 << numSegments); p++) 
                        {
                            foreach (int sm in allSegments)
                                if ((p & (1 << (numSegments - sm - 1))) > 0)
                                    literals.Add(boolState(instructorSegmentStatus[(i, d, sm)], true));
                                else
                                    literals.Add(boolState(instructorSegmentStatus[(i, d, sm)], false));
                            instructorPatternStatus.Add((i, d, p), model.NewBoolVar($"i{i}d{d}p{p}"));
                            model.Add(LinearExpr.Sum(literals) == numSegments).OnlyEnforceIf(instructorPatternStatus[(i, d, p)]);
                            model.Add(LinearExpr.Sum(literals) != numSegments).OnlyEnforceIf(instructorPatternStatus[(i, d, p)].Not());
                            literals.Clear(); 
                        }
                switch (strategyOption)
                {
                    case 1:
                        totalDeltas.Add(objWeight[2] * objPatternCost());
                        break;
                    case 2:
                        totalDeltas.Add(objWeight[2] * createDelta((1<<numSegments) * numDays * numInstructors * numSegments, objPatternCost(), 0));
                        break;
                    case 3:
                        totalDeltas.Add(objWeight[2] * createSquare(objPatternCost(), 0));
                        break;
                }
            }

            // O-04 MINIMIZE SUBJECT DIVERSITY ( numInstructor * numSubject )
            if (objOption[3] > 0)
            {
                instructorSubjectStatus = new Dictionary<(int, int), BoolVar>();
                List<ILiteral> literals = new List<ILiteral>();
                foreach (int i in allInstructors)
                    foreach (int s in allSubjects)
                    {
                        foreach (int n in allTasks)
                            if (taskSubjectMapping[n] == s)
                                literals.Add(assigns[(n, i)]);
                        if (literals.Count() == 0)
                            model.AddHint(instructorSubjectStatus[(i, s)], 0);
                        instructorSubjectStatus.Add((i, s), model.NewBoolVar($"i{i}s{s}"));
                        model.Add(LinearExpr.Sum(literals) > 0).OnlyEnforceIf(instructorSubjectStatus[(i, s)]);
                        model.Add(LinearExpr.Sum(literals) == 0).OnlyEnforceIf(instructorSubjectStatus[(i, s)].Not());
                        literals.Clear();
                    }
                switch (strategyOption)
                {
                    case 1:
                        totalDeltas.Add(objWeight[3] * objSubjectDiversity());
                        break;
                    case 2:
                        totalDeltas.Add(objWeight[3] * createDelta(numSubjects, objSubjectDiversity(), 0));
                        break;
                    case 3:
                        totalDeltas.Add(objWeight[3] * createSquare(objSubjectDiversity(), 0));
                        break;
                }   
            }

            // O-05 MINIMIZE QUOTA DIFF ( 0 )
            if (objOption[4] > 0)
            {
                switch (strategyOption)
                {
                    case 1:
                        totalDeltas.Add(objWeight[4] * objQuotaReached());
                        break;
                    case 2:
                        totalDeltas.Add(objWeight[4] * createDelta(numTasks, objQuotaReached(), 0));
                        break;
                    case 3:
                        totalDeltas.Add(objWeight[4] * createSquare(objQuotaReached(), 0));
                        break;
                }
            }

            // O-06 MINIMIZE WALKING DISTANCE ( numTask^2 )
            if (objOption[5] > 0)
            {
                /*
                NEED FURTHER OPTIMIZE
                THIS OBJECTIVE REQUIRE NON LINEAR OPTIMIZE
                MODEL SPEED DEPEND ON NUMBER OF VARIABLE
                REDUCE VARIABLE WASTE BY ADDING MORE FILTER
                */
                assignsProduct = new Dictionary<(int, int), LinearExpr>();
                List<LinearExpr> mul = new List<LinearExpr>();
                List<LinearExpr> tmp = new List<LinearExpr>();
                // symmetry breaking 
                try
                {
                    for (int n1 = 0; n1 < numTasks - 1; n1++)
                        for (int n2 = n1 + 1; n2 < numTasks; n2++)
                        {
                            // REDUCE MODEL VARIABLE WASTE
                            if (areaSlotCoefficient[taskSlotMapping[n1], taskSlotMapping[n2]] == 0 || areaDistance[taskAreaMapping[n1], taskAreaMapping[n2]] == 0)
                                continue;
                            foreach (int i in allInstructors)
                            {
                                // REDUCE MODEL VARIABLE WASTE
                                if (instructorSlot[i, taskSlotMapping[n1]] == 0 || instructorSlot[i, taskSlotMapping[n2]] == 0 || instructorSubject[i, taskSubjectMapping[n1]] == 0 || instructorSubject[i, taskSubjectMapping[n2]] == 0)
                                    continue;
                                mul.Add(assigns[(n1, i)]);
                                mul.Add(assigns[(n2, i)]);
                                LinearExpr product = model.NewBoolVar("");
                                model.AddMultiplicationEquality(product, mul);
                                tmp.Add(product);
                                mul.Clear();
                            }
                            assignsProduct.Add((n1, n2), LinearExpr.Sum(tmp));
                            tmp.Clear();
                        }
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine("An exception occurred: " + ex.Message + " on line " + ex.StackTrace);
                }
                switch (strategyOption)
                {
                    case 1:
                        totalDeltas.Add(objWeight[5] * objWalkingDistance());
                        break;
                    case 2:
                        totalDeltas.Add(objWeight[5] * createDelta(Int32.MaxValue, objWalkingDistance(), 0));
                        break;
                    case 3:
                        totalDeltas.Add(objWeight[5] * createSquare(objWalkingDistance(), 0));
                        break;
                }
            }

            // O-07 MAXIMIZE SUBJECT PREFERENCE (0)
            if (objOption[6] > 0)
            {
                switch (strategyOption)
                {
                    case 1:
                        totalDeltas.Add(-1 * objWeight[6] * objSubjectPreference());
                        break;
                    case 2:
                        totalDeltas.Add(objWeight[6] * createDelta(numTasks * 5, objSubjectPreference(), numTasks * 5));
                        break;
                    case 3:
                        totalDeltas.Add(objWeight[6] * createSquare(objSubjectPreference(), numTasks * 5));
                        break;
                }
            }

            // O-08 MAXIMIZE SLOT PREFERENCE (0)
            if (objOption[7] > 0)
            {
                switch (strategyOption)
                {
                    case 1:
                        totalDeltas.Add(-1 * objWeight[7]*objSlotPreference());
                        break;
                    case 2:
                        totalDeltas.Add(objWeight[7] * createDelta(numTasks * 5, objSlotPreference(), numTasks * 5));
                        break;
                    case 3:
                        totalDeltas.Add(objWeight[7]*createSquare(objSlotPreference(), numTasks * 5));
                        break;
                }
            }

            // SOLVING
            switch (strategyOption)
            {
                case 1:
                    model.Minimize(LinearExpr.Sum(totalDeltas));
                    break;
                case 2:
                    model.Minimize(LinearExpr.Sum(totalDeltas));
                    break;
                case 3:
                    model.Minimize(LinearExpr.Sum(totalDeltas));
                    break;
            }
            status = solver.Solve(model);
            if (status == CpSolverStatus.Optimal || status == CpSolverStatus.Feasible)
                return getResults(solver);
            else return null;
        }
        /*
        ################################
        ||          Utility           ||
        ################################
        */
        public object[] getStatistic()
        {
            return new object[] { solver.ObjectiveValue,status.ToString(), solver.NumConflicts(), solver.NumBranches(), solver.WallTime() };
/*            Console.WriteLine("Statistics");
            Console.WriteLine($"  {strategyOption}: {}");
            Console.WriteLine($"  status: {status}");
            Console.WriteLine($"  conflicts: {solver.NumConflicts()}");
            Console.WriteLine($"  branches : {solver.NumBranches()}");
            Console.WriteLine($"  wall time: {solver.WallTime()}s");*/
        }
        public LinearExpr createDelta(int maxDelta,LinearExpr actualValue,int targetValue)
        {
            IntVar delta = model.NewIntVar(0, maxDelta, "");
            model.Add(actualValue <= targetValue + delta);
            model.Add(actualValue >= targetValue - delta);
            return delta;
        }
        public LinearExpr createSquare(LinearExpr actualValue,int targetValue)
        {
            IntVar obj = model.NewIntVar(0, Int32.MaxValue, "");
            List<LinearExpr> linearExprs = new List<LinearExpr>
            {
                actualValue - targetValue,
                actualValue - targetValue
            };
            model.AddMultiplicationEquality(obj, linearExprs);
            return obj;
        }
        public ILiteral boolState(ILiteral variable,bool state)
        {
            if (state) return variable;
            else return variable.Not();
        }
        public List<List<(int, int)>> getResults(CpSolver solver)
        {
            List<(int,int)> result = new List<(int,int)> ();
            foreach (int n in allTasks)
            {
                bool isAssigned = false;
                foreach (int i in allInstructors)
                {
                    if (solver.Value(assigns[(n, i)]) == 1L)
                    {
                        isAssigned = true;
                        result.Add((n, i));
                    }
                }
                if (!isAssigned)
                {
                    result.Add((n, -1));
                }
            }
            List<List<(int, int)>> results = new List<List<(int, int)>>{result};
            return results;
        }
        public List<List<(int, int)>>? solve()
        {
            if (numBackupInstructors == -1)
            {
                if (debugLoggerOption)
                {
                    Console.WriteLine("ATTAS - Finding Optimal Backup Quota");
                }
                numBackupInstructors = findBackupInstructor();
                if (debugLoggerOption)
                {
                    Console.WriteLine($"Backup Quota = {numBackupInstructors}");
                }
            }
            if (objOption.Sum() == 0)
                return constraintOnly();
            else
                return objectiveOptimize();
        }
    }
    /*
    ################################
    ||        END OR-TOOLS        ||
    ################################
    */
}
/*  Solver
    *  1: OR-TOOLS  ( 1,2,3 )
    *  2: CPLEX ( 1 , 2 , 3 )
    *  3: NGSA-II ( 4 )
    *  Strategy
    *  1: Scalazation
    *  2: Constraint Programming
    *  3: Compromise Programming
    *  4: Pareto-based
*/