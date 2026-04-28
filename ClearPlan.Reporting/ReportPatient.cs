using System;

namespace ClearPlan.Reporting
{
    public class ReportPatient
    {
        public string Id { get; set; }
        public string planId { get; set; }
        public string planType { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string userId { get; set; }
        public Sex Sex { get; set; }
        public DateTime Birthdate { get; set; }
        public Doctor Doctor { get; set; }
    }
}
