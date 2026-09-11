using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewSnapshot
    {
        public const int CurrentSchemaVersion = 1;

        public ReviewSnapshot()
        {
            Sources = new List<ReviewSourceStatus>();
            Plans = new List<ReviewPlanRow>();
            PqmRows = new List<ReviewPqmRow>();
            PlanCheckRows = new List<ReviewCheckRow>();
            FieldRows = new List<ReviewFieldRow>();
            StructureMappings = new List<ReviewStructureMapping>();
            DvhSeries = new List<ReviewDvhSeries>();
            Report = new ReviewReportMetadata();
            PlanAnalysis = new PlanAnalysis.ReviewPlanAnalysis();
            PlanImages = new List<ReviewPlanImage>();
        }

        [JsonProperty("schemaVersion", Order = 0)]
        public int SchemaVersion { get; set; }

        [JsonProperty("scenarioId", Order = 1)]
        public string ScenarioId { get; set; }

        [JsonProperty("scenarioTitle", Order = 2)]
        public string ScenarioTitle { get; set; }

        [JsonProperty("scenarioDescription", Order = 3)]
        public string ScenarioDescription { get; set; }

        [JsonProperty("seed", Order = 4)]
        public int Seed { get; set; }

        [JsonProperty("synthetic", Order = 5)]
        public bool Synthetic { get; set; }

        [JsonProperty("generatedUtc", Order = 6)]
        public DateTimeOffset GeneratedUtc { get; set; }

        [JsonProperty("patientDisplayLabel", Order = 7)]
        public string PatientDisplayLabel { get; set; }

        [JsonProperty("planDisplayLabel", Order = 8)]
        public string PlanDisplayLabel { get; set; }

        [JsonProperty("provenanceText", Order = 9)]
        public string ProvenanceText { get; set; }

        [JsonProperty("activePlanKey", Order = 10)]
        public string ActivePlanKey { get; set; }

        [JsonProperty("sources", Order = 11)]
        public List<ReviewSourceStatus> Sources { get; set; }

        [JsonProperty("plans", Order = 12)]
        public List<ReviewPlanRow> Plans { get; set; }

        [JsonProperty("pqmRows", Order = 13)]
        public List<ReviewPqmRow> PqmRows { get; set; }

        [JsonProperty("planCheckRows", Order = 14)]
        public List<ReviewCheckRow> PlanCheckRows { get; set; }

        [JsonProperty("disabledCheckCount", Order = 21)]
        public int DisabledCheckCount { get; set; }

        [JsonProperty("fieldRows", Order = 15)]
        public List<ReviewFieldRow> FieldRows { get; set; }

        [JsonProperty("structureMappings", Order = 16)]
        public List<ReviewStructureMapping> StructureMappings { get; set; }

        [JsonProperty("dvhSeries", Order = 17)]
        public List<ReviewDvhSeries> DvhSeries { get; set; }

        [JsonProperty("report", Order = 18)]
        public ReviewReportMetadata Report { get; set; }

        [JsonProperty("planAnalysis", Order = 19)]
        public PlanAnalysis.ReviewPlanAnalysis PlanAnalysis { get; set; }

        [JsonProperty("planImages", Order = 20)]
        public List<ReviewPlanImage> PlanImages { get; set; }

        [OnDeserialized]
        internal void RestoreOptionalCollections(StreamingContext context)
        {
            Sources = Sources ?? new List<ReviewSourceStatus>();
            Plans = Plans ?? new List<ReviewPlanRow>();
            PqmRows = PqmRows ?? new List<ReviewPqmRow>();
            PlanCheckRows = PlanCheckRows ?? new List<ReviewCheckRow>();
            FieldRows = FieldRows ?? new List<ReviewFieldRow>();
            StructureMappings = StructureMappings ?? new List<ReviewStructureMapping>();
            DvhSeries = DvhSeries ?? new List<ReviewDvhSeries>();
            Report = Report ?? new ReviewReportMetadata();
            PlanAnalysis = PlanAnalysis ?? new PlanAnalysis.ReviewPlanAnalysis();
            PlanImages = PlanImages ?? new List<ReviewPlanImage>();
        }
    }

    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewStructureMapping
    {
        public ReviewStructureMapping()
        {
            AvailableStructureIds = new List<string>();
        }

        [JsonProperty("stableId", Order = 0)]
        public string StableId { get; set; }

        [JsonProperty("templateStructure", Order = 1)]
        public string TemplateStructure { get; set; }

        [JsonProperty("selectedStructureId", Order = 2)]
        public string SelectedStructureId { get; set; }

        [JsonProperty("availableStructureIds", Order = 3)]
        public List<string> AvailableStructureIds { get; set; }

        [JsonProperty("status", Order = 4)]
        public string Status { get; set; }

        [JsonProperty("message", Order = 5)]
        public string Message { get; set; }

        [OnDeserialized]
        internal void RestoreOptionalCollections(StreamingContext context)
        {
            AvailableStructureIds = AvailableStructureIds ?? new List<string>();
        }
    }

    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewReportMetadata
    {
        public ReviewReportMetadata()
        {
            Notes = new List<string>();
        }

        [JsonProperty("title", Order = 0)]
        public string Title { get; set; }

        [JsonProperty("subtitle", Order = 1)]
        public string Subtitle { get; set; }

        [JsonProperty("modeLabel", Order = 2)]
        public string ModeLabel { get; set; }

        [JsonProperty("watermark", Order = 3)]
        public string Watermark { get; set; }

        [JsonProperty("outputFileLabel", Order = 4)]
        public string OutputFileLabel { get; set; }

        [JsonProperty("notes", Order = 5)]
        public List<string> Notes { get; set; }

        [OnDeserialized]
        internal void RestoreOptionalCollections(StreamingContext context)
        {
            Notes = Notes ?? new List<string>();
        }
    }

    public static class ReviewStatusCodes
    {
        public const string Pass = "pass";
        public const string Variation = "variation";
        public const string Fail = "fail";
        public const string Info = "info";
        public const string NotEvaluated = "not-evaluated";
        public const string Available = "available";
        public const string Fallback = "fallback";
        public const string Unavailable = "unavailable";
        public const string NotConfigured = "not-configured";
    }

    public static class ReviewSeverityCodes
    {
        public const string None = "none";
        public const string Info = "info";
        public const string Warning = "warning";
        public const string Error = "error";
    }

    public static class ReviewUnitCodes
    {
        public const string Gray = "Gy";
        public const string Percent = "%";
        public const string CubicCentimeter = "cm3";
        public const string Count = "count";
        public const string Text = "text";
        public const string Boolean = "boolean";
        public const string Degree = "degree";
    }

    public static class ReviewDvhRoleCodes
    {
        public const string Target = "target";
        public const string OrganAtRisk = "oar";
        public const string External = "external";
        public const string Other = "other";
    }

    public static class ReviewLineStyleCodes
    {
        public const string Solid = "solid";
        public const string Dash = "dash";
        public const string Dot = "dot";
    }
}
