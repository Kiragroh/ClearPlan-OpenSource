namespace ClearPlan.Core.Fields
{
    public sealed class FieldNameSuggestion
    {
        public string CurrentId { get; set; }
        public string ExpectedId { get; set; }
        public string CurrentName { get; set; }
        public int BeamNumber { get; set; }
        public int DisplayOrder { get; set; }
        public string SuggestedName { get; set; }
        public bool IdWouldChange { get; set; }
        public bool NameWouldChange { get; set; }
        public bool WouldChange { get; set; }
        public bool IsEvaluated { get; set; } = true;
        public string EvaluationMessage { get; set; }
    }
}
