using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewDvhPoint
    {
        public ReviewDvhPoint()
        {
        }

        public ReviewDvhPoint(double doseGy, double volumePercent)
        {
            DoseGy = doseGy;
            VolumePercent = volumePercent;
        }

        [JsonProperty("doseGy", Order = 0)]
        public double DoseGy { get; set; }

        [JsonProperty("volumePercent", Order = 1)]
        public double VolumePercent { get; set; }
    }
}
