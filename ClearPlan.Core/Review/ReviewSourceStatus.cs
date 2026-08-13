using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewSourceStatus
    {
        [JsonProperty("stableId", Order = 0)]
        public string StableId { get; set; }

        [JsonProperty("sourceCode", Order = 1)]
        public string SourceCode { get; set; }

        [JsonProperty("sourceType", Order = 2)]
        public string SourceType { get; set; }

        [JsonProperty("status", Order = 3)]
        public string Status { get; set; }

        [JsonProperty("optional", Order = 4)]
        public bool Optional { get; set; }

        [JsonProperty("usedFallback", Order = 5)]
        public bool UsedFallback { get; set; }

        [JsonProperty("pathDisplayLabel", Order = 6)]
        public string PathDisplayLabel { get; set; }

        [JsonProperty("message", Order = 7)]
        public string Message { get; set; }
    }
}
