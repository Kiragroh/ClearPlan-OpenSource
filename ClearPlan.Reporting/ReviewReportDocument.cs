using System;
using System.Collections.Generic;

namespace ClearPlan.Reporting
{
    public sealed class ReviewReportDocument
    {
        public ReviewReportDocument()
        {
            Notes = new List<string>();
            Sources = new List<ReviewReportSourceRow>();
            Plans = new List<ReviewReportPlanRow>();
            PqmRows = new List<ReviewReportPqmRow>();
            PlanCheckRows = new List<ReviewReportCheckRow>();
            FieldRows = new List<ReviewReportFieldRow>();
            StructureMappings =
                new List<ReviewReportStructureMappingRow>();
            DvhSeries = new List<ReviewReportDvhSeries>();
        }

        public int SchemaVersion { get; set; }
        public string ScenarioId { get; set; }
        public string ScenarioTitle { get; set; }
        public string ScenarioDescription { get; set; }
        public int Seed { get; set; }
        public bool Synthetic { get; set; }
        public DateTimeOffset GeneratedUtc { get; set; }
        public string PatientDisplayLabel { get; set; }
        public string PlanDisplayLabel { get; set; }
        public string ProvenanceText { get; set; }
        public string ActivePlanKey { get; set; }
        public string Title { get; set; }
        public string Subtitle { get; set; }
        public string ModeLabel { get; set; }
        public string Watermark { get; set; }
        public string OutputFileLabel { get; set; }
        public List<string> Notes { get; set; }
        public List<ReviewReportSourceRow> Sources { get; set; }
        public List<ReviewReportPlanRow> Plans { get; set; }
        public List<ReviewReportPqmRow> PqmRows { get; set; }
        public List<ReviewReportCheckRow> PlanCheckRows { get; set; }
        public List<ReviewReportFieldRow> FieldRows { get; set; }
        public List<ReviewReportStructureMappingRow> StructureMappings
        {
            get;
            set;
        }

        public List<ReviewReportDvhSeries> DvhSeries { get; set; }
    }

    public sealed class ReviewReportSourceRow
    {
        public string StableId { get; set; }
        public string SourceCode { get; set; }
        public string SourceType { get; set; }
        public string Status { get; set; }
        public bool Optional { get; set; }
        public bool UsedFallback { get; set; }
        public string PathDisplayLabel { get; set; }
        public string Message { get; set; }
    }

    public sealed class ReviewReportPlanRow
    {
        public string PlanKey { get; set; }
        public string DisplayLabel { get; set; }
        public DateTimeOffset? CreatedUtc { get; set; }
        public double? DosePerFractionGy { get; set; }
        public double? TotalDoseGy { get; set; }
        public int? FractionCount { get; set; }
        public string TargetDisplayLabel { get; set; }
        public string Status { get; set; }
    }

    public sealed class ReviewReportPqmRow
    {
        public ReviewReportPqmRow()
        {
            StructureOptions = new List<string>();
        }

        public string StableId { get; set; }
        public string TemplateCode { get; set; }
        public string TemplateStructure { get; set; }
        public string ResolvedStructureId { get; set; }
        public List<string> StructureOptions { get; set; }
        public string Objective { get; set; }
        public string Comparator { get; set; }
        public double? Goal { get; set; }
        public double? Variation { get; set; }
        public double? AchievedValue { get; set; }
        public string Unit { get; set; }
        public string Status { get; set; }
        public string Severity { get; set; }
        public string Explanation { get; set; }
    }

    public sealed class ReviewReportCheckRow
    {
        public string CheckCode { get; set; }
        public string Category { get; set; }
        public string Status { get; set; }
        public string Severity { get; set; }
        public string ObservedValue { get; set; }
        public string ExpectedValue { get; set; }
        public string Unit { get; set; }
        public string Message { get; set; }
    }

    public sealed class ReviewReportFieldRow
    {
        public string StableId { get; set; }
        public int TreatmentOrder { get; set; }
        public int BeamNumber { get; set; }
        public string CurrentId { get; set; }
        public string ExpectedId { get; set; }
        public string CurrentName { get; set; }
        public string SuggestedName { get; set; }
        public string IdStatus { get; set; }
        public string NameStatus { get; set; }
    }

    public sealed class ReviewReportStructureMappingRow
    {
        public ReviewReportStructureMappingRow()
        {
            AvailableStructureIds = new List<string>();
        }

        public string StableId { get; set; }
        public string TemplateStructure { get; set; }
        public string SelectedStructureId { get; set; }
        public List<string> AvailableStructureIds { get; set; }
        public string Status { get; set; }
        public string Message { get; set; }
    }

    public sealed class ReviewReportDvhSeries
    {
        public ReviewReportDvhSeries()
        {
            Points = new List<ReviewReportDvhPoint>();
        }

        public string StableId { get; set; }
        public string StructureId { get; set; }
        public string DisplayName { get; set; }
        public string Role { get; set; }
        public string ColorHex { get; set; }
        public string LineStyle { get; set; }
        public bool Selected { get; set; }
        public double? VolumeCc { get; set; }
        public string DoseUnit { get; set; }
        public string VolumeUnit { get; set; }
        public List<ReviewReportDvhPoint> Points { get; set; }
    }

    public sealed class ReviewReportDvhPoint
    {
        public double DoseGy { get; set; }
        public double VolumePercent { get; set; }
    }
}
