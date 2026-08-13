using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewCheckRow
    {
        [JsonProperty("checkCode", Order = 0)]
        public string CheckCode { get; set; }

        [JsonProperty("category", Order = 1)]
        public string Category { get; set; }

        [JsonProperty("status", Order = 2)]
        public string Status { get; set; }

        [JsonProperty("severity", Order = 3)]
        public string Severity { get; set; }

        [JsonProperty("observedValue", Order = 4)]
        public string ObservedValue { get; set; }

        [JsonProperty("expectedValue", Order = 5)]
        public string ExpectedValue { get; set; }

        [JsonProperty("unit", Order = 6)]
        public string Unit { get; set; }

        [JsonProperty("message", Order = 7)]
        public string Message { get; set; }
    }
}
