namespace ATTAS_API.Models
{
    public class Solution
    {
        public int Id { get; set; }
        public int sessionId { get; set; }
        public int no { get; set; }
        public int taskAssigned { get; set; }
        public int workingDay { get; set; }
        public int workingTime { get; set; }
        public int waitingTime { get; set; }
        public int subjectDiversity { get; set; }
        public int quotaAvailable { get; set; }
        public int walkingDistance { get; set; }
        public int subjectPreference { get; set; }
        public int slotPreference { get; set; }
    }
}
