using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewFieldRow
    {
        [JsonProperty("stableId", Order = 0)]
        public string StableId { get; set; }

        [JsonProperty("treatmentOrder", Order = 1)]
        public int TreatmentOrder { get; set; }

        [JsonProperty("beamNumber", Order = 2)]
        public int BeamNumber { get; set; }

        [JsonProperty("currentId", Order = 3)]
        public string CurrentId { get; set; }

        [JsonProperty("expectedId", Order = 4)]
        public string ExpectedId { get; set; }

        [JsonProperty("currentName", Order = 5)]
        public string CurrentName { get; set; }

        [JsonProperty("suggestedName", Order = 6)]
        public string SuggestedName { get; set; }

        [JsonProperty("idStatus", Order = 7)]
        public string IdStatus { get; set; }

        [JsonProperty("nameStatus", Order = 8)]
        public string NameStatus { get; set; }
    }
}
