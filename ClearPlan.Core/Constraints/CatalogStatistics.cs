namespace ClearPlan.Core.Constraints
{
    public sealed class CatalogStatistics
    {
        public int TotalTables { get; set; }
        public int ActiveTables { get; set; }
        public int TotalConstraints { get; set; }
        public int TotalStructures { get; set; }
        public int ActiveStructures { get; set; }
    }
}
