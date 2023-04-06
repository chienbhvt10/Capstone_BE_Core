namespace ATTAS_API.Models
{
    public class Data
    {
        public string token { get; set; }
        public string? sessionHash { get; set; }
        public Setting Setting { get; set; }
        public List<Task> tasks { get; set; }
        public List<Slot> slots { get; set; }
        public List<Instructor> instructors { get; set; }
        public int numTasks { get; set; }
        public int numInstructors { get; set; }
        public int numSlots { get; set; }
        public int numDays { get; set; }
        public int numTimes { get; set; }
        public int numSegments { get; set; }
        public int numSegmentRules { get; set; }
        public int numSubjects { get; set; }
        public int numAreas { get; set; }
        public int backupInstructor { get; set; }
        public List<List<int>> slotConflict { get; set; }
        public List<List<int>> slotDay { get; set; }
        public List<List<int>> slotTime { get; set; }
        public List<List<int>> slotSegment { get; set; }
        public List<int> patternCost { get; set; }
        public List<List<int>> instructorSubject { get; set; }
        public List<List<int>> instructorSlot { get; set; }
        public List<int> instructorQuota { get; set; }
        public List<int> instructorMinQuota { get; set; }
        public List<List<int>> areaDistance { get; set; }
        public List<List<int>> areaSlotCoefficient { get; set; }
        public List<Preassign>? preassigns { get; set; } 

    }
}
