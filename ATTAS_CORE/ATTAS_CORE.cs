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
        public double maxSearchingTimeOption { get; set; } = 30.0;
        public int strategyOption { get; set; } = 2;
        public int[] objOption { get; set; } = new int[6] { 1, 1, 1, 1, 1, 1 };
        public int[] objWeight { get; set; } = new int[6] { 1, 1, 1, 1, 1, 1 };
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
        public int numInstructors { get; set; } = 0;
        public int numBackupInstructors { get; set; } = 0;
        public int numAreas { get; set; } = 0;
        //RANGE
        private int[] allSubjects = Array.Empty<int>();
        private int[] allTasks = Array.Empty<int>();
        private int[] allSlots = Array.Empty<int>();
        private int[] allInstructors = Array.Empty<int>();
        private int[] allInstructorsWithBackup = Array.Empty<int>();
        //private int[] allAreas = Array.Empty<int>();
        //INPUT DATA
        public int[,] slotConflict { get; set; } = new int[0, 0];
        public int[,] slotCompatibilityCost { get; set; } = new int[0, 0];
        public int[,] instructorSubject { get; set; } = new int[0, 0];
        public int[,] instructorSubjectPreference { get; set; } = new int[0, 0];
        public int[,] instructorSlot { get; set; } = new int[0, 0];
        public int[,] instructorSlotPreference { get; set; } = new int[0, 0];
        public List<(int, int, int)> instructorPreassign { get; set; } = new List<(int, int, int)>();
        public int[] instructorQuota { get; set; } = Array.Empty<int>();
        public int[] taskSubjectMapping { get; set; } = Array.Empty<int>();
        public int[] taskSlotMapping { get; set; } = Array.Empty<int>();
        public int[] taskAreaMapping { get; set; } = Array.Empty<int>();
        public int[,] areaDistance { get; set; } = new int[0, 0];
        public int[,] areaSlotWeight { get; set; } = new int[0, 0];  

        /*
        ################################
        ||           MODEL            ||
        ################################
         */

        private CpModel model;
        // Desicion variable
        private Dictionary<(int, int), BoolVar> assigns;
        private Dictionary<(int, int), BoolVar> instructorSubjectStatus;
        private Dictionary<(int, int), LinearExpr> assignsProduct;
        public void setSolverCount()
        {
            allSubjects = Enumerable.Range(0, numSubjects).ToArray();
            allTasks = Enumerable.Range(0, numTasks).ToArray();
            allSlots = Enumerable.Range(0, numSlots).ToArray();
            allInstructors = Enumerable.Range(0, numInstructors).ToArray();
            //allAreas = Enumerable.Range(0, numAreas).ToArray();

            if (numBackupInstructors > 0)
            {
                allInstructorsWithBackup = Enumerable.Range(0, numInstructors + 1).ToArray();
                instructorQuota = instructorQuota.Concat(new int[] { numBackupInstructors }).ToArray();
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
                model.AddLinearConstraint(LinearExpr.Sum(taskAssigned), 0, instructorQuota[i]);
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
                    if (taskSlotMapping[n] == s)
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
                    model.Add(assigns[(data.Item2, data.Item1)] == 1);
                if (data.Item3 == -1)
                    model.Add(assigns[(data.Item2, data.Item1)] == 0);
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
        public List<(int, int)> constraintOnly()
        {
            setSolverCount();
            createModel();
            List<ILiteral> obj = new List<ILiteral> ();
            foreach (int n in allTasks)
                foreach (int i in allInstructors)
                    obj.Add(assigns[(n, i)]);
            model.Minimize( createDelta( numTasks ,LinearExpr.Sum(obj),numTasks ));
            CpSolver solver = new CpSolver();
            solver.StringParameters += "linearization_level:0 " + $"max_time_in_seconds:{maxSearchingTimeOption} ";
            // Tell the solver to enumerate all solutions.
            //solver.StringParameters += "linearization_level:0 " + "enumerate_all_solutions:true " + $"max_time_in_seconds:{maxSearchingTimeOption} " + $"random_seed:{rnd.Next(1, 31)} ";
            //int solutionLimit = 1;
            //SolutionPrinter cb = new SolutionPrinter(allInstructors, allTasks, assigns, solutionLimit, debugLoggerOption);
            CpSolverStatus status = solver.Solve(model);

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
        /*
        ################################
        ||         OBJECTIVE          ||
        ################################
        */
        //O-01
        public LinearExpr objSlotCompatibilityCost()
        {
            List<LinearExpr> slotCompatibility_ = new List<LinearExpr>();
            for (int n1 = 0; n1 < numTasks - 1; n1++)
                for (int n2 = n1 + 1; n2 < numTasks; n2++)
                {
                    if (slotCompatibilityCost[taskSlotMapping[n1], taskSlotMapping[n2]] == 0)
                        continue;
                    slotCompatibility_.Add(assignsProduct[(n1, n2)] * slotCompatibilityCost[taskSlotMapping[n1], taskSlotMapping[n2]]);
                }
            return LinearExpr.Sum(slotCompatibility_);
        }
        //O-02
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
        //O-03
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
        //O-04
        public LinearExpr objWalkingDistance()
        {
            List<LinearExpr> walkingDistance = new List<LinearExpr>();
            for (int n1 = 0; n1 < numTasks - 1; n1++)
                for (int n2 = n1 + 1; n2 < numTasks; n2++)
                {
                    if (areaSlotWeight[taskSlotMapping[n1], taskSlotMapping[n2]] == 0 || areaDistance[taskAreaMapping[n1], taskAreaMapping[n2]] == 0)
                        continue;
                    walkingDistance.Add(assignsProduct[(n1, n2)] * areaSlotWeight[taskSlotMapping[n1], taskSlotMapping[n2]] * areaDistance[taskAreaMapping[n1], taskAreaMapping[n2]]);
                }
            return LinearExpr.Sum(walkingDistance);
        }
        //O-05
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
        //O-06
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
        public List<(int, int)>? objectiveOptimize()
        {
            setSolverCount();
            createModel();
            CpSolver solver = new CpSolver();
            CpSolverStatus status = new CpSolverStatus();
            solver.StringParameters += "linearization_level:0 " + $"max_time_in_seconds:{maxSearchingTimeOption} ";

            List<int> weights = new List<int>();
            List<LinearExpr> totalDeltas = new List<LinearExpr>();
            //O-03 MINIMIZE QUOTA DIFF
            if (objOption[2] > 0)
            {
                switch (strategyOption)
                {
                    case 1:
                        weights.Add(objWeight[2]);
                        totalDeltas.Add(objQuotaReached());
                        break;
                    case 2:
                        weights.Add(objWeight[2]);
                        totalDeltas.Add(createDelta(numTasks, objQuotaReached(), 0));
                        break;
                    case 3:
                        weights.Add(objWeight[2]);
                        totalDeltas.Add(createPow2(objQuotaReached(), 0));
                        break;
                }
            }
            //O-05 MAXIMIZE SUBJECT PREFERENCE
            if (objOption[4] > 0)
            {
                switch (strategyOption)
                {
                    case 1:
                        weights.Add(-1 * objWeight[4]);
                        totalDeltas.Add(objSubjectPreference());
                        break;
                    case 2:
                        weights.Add(objWeight[4]);
                        totalDeltas.Add(createDelta(numTasks * 5, objSubjectPreference(), numTasks * 5));
                        break;
                    case 3:
                        weights.Add(objWeight[4]);
                        totalDeltas.Add(createPow2(objSubjectPreference(),numTasks*5));
                        break;
                }
                
            }
            //O-06 MAXIMIZE SLOT PREFERENCE
            if (objOption[5] > 0)
            {
                switch (strategyOption)
                {
                    case 1:
                        weights.Add(-1 * objWeight[5]);
                        totalDeltas.Add(objSlotPreference());
                        break;
                    case 2:
                        weights.Add(objWeight[5]);
                        totalDeltas.Add(createDelta(numTasks * 5, objSlotPreference(), numTasks * 5));
                        break;
                    case 3:
                        weights.Add(objWeight[5]);
                        totalDeltas.Add(createPow2(objSlotPreference(),numTasks*5));
                        break;
                }
                
            }
            // O-02 MINIMIZE SUBJECT DIVERSITY
            if (objOption[1] > 0)
            {
                instructorSubjectStatus = new Dictionary<(int, int), BoolVar>();
                List<ILiteral> literals = new List<ILiteral>();
                foreach (int i in allInstructors)
                    foreach (int s in allSubjects)
                    {
                        foreach (int n in allTasks)
                            if (taskSubjectMapping[n] == s)
                                literals.Add(assigns[(n, i)]);
                        instructorSubjectStatus.Add((i, s), model.NewBoolVar($"i{i}s{s}"));
                        model.Add(LinearExpr.Sum(literals) > 0).OnlyEnforceIf(instructorSubjectStatus[(i, s)]);
                        model.Add(LinearExpr.Sum(literals) == 0).OnlyEnforceIf(instructorSubjectStatus[(i, s)].Not());
                        literals.Clear();
                    }
                switch (strategyOption)
                {
                    case 1:
                        weights.Add(objWeight[1]);
                        totalDeltas.Add(objSubjectDiversity());
                        break;
                    case 2:
                        weights.Add(objWeight[1]);
                        totalDeltas.Add(createDelta(numSubjects, objSubjectDiversity(), 0));
                        break;
                    case 3:
                        weights.Add(objWeight[1]);
                        totalDeltas.Add(createPow2(objSubjectDiversity(), 0));
                        break;
                }
                
            }
            
            if (objOption[0] > 0 || objOption[3] > 0)
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
                            if ((areaSlotWeight[taskSlotMapping[n1], taskSlotMapping[n2]] == 0 || areaDistance[taskAreaMapping[n1], taskAreaMapping[n2]] == 0) && slotCompatibilityCost[taskSlotMapping[n1], taskSlotMapping[n2]] == 0)
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
        }
            
            //O-01 MINIMIZE SLOT COMPATIBILITY COST
            if (objOption[0] > 0)
            {
                switch (strategyOption)
                {
                    case 1:
                        weights.Add(objWeight[0]);
                        totalDeltas.Add(objSlotCompatibilityCost());
                        break;
                    case 2:
                        weights.Add(objWeight[0]);
                        totalDeltas.Add(createDelta(numTasks * numTasks * 5, objSlotCompatibilityCost(), numTasks * numTasks * -5));
                        break;
                    case 3:
                        weights.Add(objWeight[0]);
                        totalDeltas.Add(createPow2(objSlotCompatibilityCost(), 0));
                        break;
                }
                
            }
            //O-04 MINIMIZE WALKING DISTANCE
            if (objOption[3] > 0)
            {
                switch (strategyOption)
                {
                    case 1:
                        weights.Add(objWeight[3]);
                        totalDeltas.Add(objWalkingDistance());
                        break;
                    case 2:
                        weights.Add(objWeight[3]);
                        totalDeltas.Add(createDelta(numTasks * numTasks * 5 * 5, objWalkingDistance(), 0));
                        break;
                    case 3:
                        weights.Add(objWeight[3]);
                        totalDeltas.Add(createPow2(objWalkingDistance(), 0));
                        break;
                }
            }
            switch (strategyOption)
            {
                case 1:
                    model.Minimize(LinearExpr.WeightedSum(totalDeltas, weights));
                    break;
                case 2:
                    model.Minimize(LinearExpr.WeightedSum(totalDeltas, weights));
                    break;
                case 3:
                    model.Minimize(LinearExpr.WeightedSum(totalDeltas, weights));
                    break;
            }
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
        /*
        ################################
        ||          Utility           ||
        ################################
        */
        public LinearExpr createDelta(int maxDelta,LinearExpr actualValue,int targetValue)
        {
            IntVar delta = model.NewIntVar(0, maxDelta, "");
            model.Add(actualValue <= targetValue + delta);
            model.Add(actualValue >= targetValue - delta);
            return delta;
        }
        public LinearExpr createPow2(LinearExpr actualValue,int targetValue)
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
        public List<(int,int)> getResults(CpSolver solver)
        {
            List<(int,int)> results = new List<(int,int)> ();
            foreach (int n in allTasks)
            {
                bool isAssigned = false;
                foreach (int i in allInstructors)
                {
                    if (solver.Value(assigns[(n, i)]) == 1L)
                    {
                        isAssigned = true;
                        results.Add((n, i));
                    }
                }
                if (!isAssigned)
                {
                    results.Add((n, -1));
                }
            }
            return results;
        }
        public List<(int, int)>? solve()
        {
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
    
    public class ATTAS_NSGA2 
    {
        
    }
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