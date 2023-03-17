namespace ATTAS_API.Models
{
    public class Setting
    {
        public int maxSearchingTime { get; set; }
        public int solver { get; set; }
        public int strategy { get; set; }
        public List<int> objectiveOption { get; set; }
        public List<int> objectiveWeight { get; set; }
    }
}
