using System;

namespace ClearPlan.Reporting
{
    public class ReportPlanningItem
    {
        public string Id { get; set; }
        public string Type { get; set; }
        public DateTime Created { get; set; }
        public string TreatmentApproverDisplayName { get; set; }
        public string ConstraintPath { get; set; }        
        public int PlansInPlansum { get; set; }
        public bool PrintPC_CheckboxState { get; set; }
        public double prinTtotaldose { get; set; }
        public double prinTfractiondose { get; set; }
        public string prinTfractions { get; set; }
        public string algorithmus { get; set; }
        public string technique { get; set; }
        public string planningApprover { get; set; }
        public string orientation { get; set; }
        public string treatmentapproveDate { get; set; }
        public string gatingExtenedString { get; set; }
    }
}
