using System.Collections.Generic;
using System.Runtime.Serialization;
using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewPqmRow
    {
        public ReviewPqmRow()
        {
            StructureOptions = new List<string>();
        }

        [JsonProperty("stableId", Order = 0)]
        public string StableId { get; set; }

        [JsonProperty("templateCode", Order = 1)]
        public string TemplateCode { get; set; }

        [JsonProperty("templateStructure", Order = 2)]
        public string TemplateStructure { get; set; }

        [JsonProperty("resolvedStructureId", Order = 3)]
        public string ResolvedStructureId { get; set; }

        [JsonProperty("structureOptions", Order = 4)]
        public List<string> StructureOptions { get; set; }

        [JsonProperty("objective", Order = 5)]
        public string Objective { get; set; }

        [JsonProperty("comparator", Order = 6)]
        public string Comparator { get; set; }

        [JsonProperty("goal", Order = 7)]
        public double? Goal { get; set; }

        [JsonProperty("variation", Order = 8)]
        public double? Variation { get; set; }

        [JsonProperty("achievedValue", Order = 9)]
        public double? AchievedValue { get; set; }

        [JsonProperty("unit", Order = 10)]
        public string Unit { get; set; }

        [JsonProperty("status", Order = 11)]
        public string Status { get; set; }

        [JsonProperty("severity", Order = 12)]
        public string Severity { get; set; }

        [JsonProperty("explanation", Order = 13)]
        public string Explanation { get; set; }

        [OnDeserialized]
        internal void RestoreOptionalCollections(StreamingContext context)
        {
            StructureOptions = StructureOptions ?? new List<string>();
        }
    }
}
