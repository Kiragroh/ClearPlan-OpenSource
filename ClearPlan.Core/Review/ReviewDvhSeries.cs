using System.Collections.Generic;
using System.Runtime.Serialization;
using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewDvhSeries
    {
        public ReviewDvhSeries()
        {
            Points = new List<ReviewDvhPoint>();
        }

        [JsonProperty("stableId", Order = 0)]
        public string StableId { get; set; }

        [JsonProperty("structureId", Order = 1)]
        public string StructureId { get; set; }

        [JsonProperty("displayName", Order = 2)]
        public string DisplayName { get; set; }

        [JsonProperty("role", Order = 3)]
        public string Role { get; set; }

        [JsonProperty("colorHex", Order = 4)]
        public string ColorHex { get; set; }

        [JsonProperty("lineStyle", Order = 5)]
        public string LineStyle { get; set; }

        [JsonProperty("selected", Order = 6)]
        public bool Selected { get; set; }

        [JsonProperty("volumeCc", Order = 7)]
        public double? VolumeCc { get; set; }

        [JsonProperty("points", Order = 8)]
        public List<ReviewDvhPoint> Points { get; set; }

        // Native TPS statistics, captured in Gy independently of the plotted bins.
        // Older/curve-only snapshots leave these null; zero is a valid dose value.
        [JsonProperty("minimumDoseGy", Order = 9)]
        public double? MinimumDoseGy { get; set; }

        [JsonProperty("meanDoseGy", Order = 10)]
        public double? MeanDoseGy { get; set; }

        [JsonProperty("maximumDoseGy", Order = 11)]
        public double? MaximumDoseGy { get; set; }

        [JsonProperty("targetKind", Order = 12, DefaultValueHandling = DefaultValueHandling.Ignore)]
        public string TargetKind { get; set; }

        [JsonProperty("requiredForTargetReview", Order = 13, DefaultValueHandling = DefaultValueHandling.Ignore)]
        public bool RequiredForTargetReview { get; set; }

        [JsonProperty("targetSelectionReason", Order = 14, DefaultValueHandling = DefaultValueHandling.Ignore)]
        public string TargetSelectionReason { get; set; }

        [JsonProperty("d98DoseGy", Order = 15, DefaultValueHandling = DefaultValueHandling.Ignore)]
        public double? D98DoseGy { get; set; }

        [OnDeserialized]
        internal void RestoreOptionalCollections(StreamingContext context)
        {
            Points = Points ?? new List<ReviewDvhPoint>();
        }
    }
}
