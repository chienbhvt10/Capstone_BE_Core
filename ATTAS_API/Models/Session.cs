namespace ATTAS_API.Models
{
    public class Session
    {
        public int id { get; set; }
        public string hash { get; set; }
        public int statusId { get; set; }
        public int solutionCount { get; set; }
    }
}
