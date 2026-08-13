namespace ClearPlan.Core.Constraints
{
    public enum CatalogIssueSeverity
    {
        Information,
        Warning,
        Error
    }

    public sealed class CatalogValidationIssue
    {
        public CatalogIssueSeverity Severity { get; set; }
        public string SourceLocation { get; set; }
        public string Field { get; set; }
        public string RawValue { get; set; }
        public string Message { get; set; }

        public bool IsFatal
        {
            get { return Severity == CatalogIssueSeverity.Error; }
        }
    }
}
