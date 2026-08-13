namespace ClearPlan.Reporting
{
    public interface IReport
    {
        void Export(string path, ReportData data);
    }
}
