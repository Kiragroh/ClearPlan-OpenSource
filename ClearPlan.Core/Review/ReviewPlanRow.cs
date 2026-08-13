using System;
using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewPlanRow
    {
        [JsonProperty("planKey", Order = 0)]
        public string PlanKey { get; set; }

        [JsonProperty("displayLabel", Order = 1)]
        public string DisplayLabel { get; set; }

        [JsonProperty("createdUtc", Order = 2)]
        public DateTimeOffset? CreatedUtc { get; set; }

        [JsonProperty("dosePerFractionGy", Order = 3)]
        public double? DosePerFractionGy { get; set; }

        [JsonProperty("totalDoseGy", Order = 4)]
        public double? TotalDoseGy { get; set; }

        [JsonProperty("fractionCount", Order = 5)]
        public int? FractionCount { get; set; }

        [JsonProperty("targetDisplayLabel", Order = 6)]
        public string TargetDisplayLabel { get; set; }

        [JsonProperty("status", Order = 7)]
        public string Status { get; set; }
    }
}
