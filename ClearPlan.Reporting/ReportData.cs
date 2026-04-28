namespace ClearPlan.Reporting
{
    public class ReportData
    {
        public ReportPatient ReportPatient { get; set; }
        public ReportPlanningItem ReportPlanningItem { get; set; }
        public ReportStructureSet ReportStructureSet { get; set; }
        public ReportPQMs ReportPQMs { get; set; }
        public ReportPCs ReportPCs { get; set; }
        public ReportRefs ReportRefs { get; set; }
        public ReportDVH ReportDVH { get; set; }
        public ReportDVHstats ReportDVHstats { get; set; }
    }
}
