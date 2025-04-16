namespace Employee_Management_System.Models
{
    public class Dependent
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Relation { get; set; }
        public int Age { get; set; }

        public int EmployeeId { get; set; }
        public virtual Employee Employee { get; set; }
    }
}
