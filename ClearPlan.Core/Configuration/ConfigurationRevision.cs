using System;
using Newtonsoft.Json;

namespace ClearPlan.Core.Configuration
{
    /// <summary>Local provenance, not an authenticated identity or a tamper-proof audit record.</summary>
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ConfigurationRevision
    {
        [JsonProperty("key")] public string Key { get; internal set; }
        [JsonProperty("category")] public string Category { get; internal set; }
        [JsonProperty("extension")] public string Extension { get; internal set; }
        [JsonProperty("revision_number")] public int RevisionNumber { get; internal set; }
        [JsonProperty("sha256")] public string Sha256 { get; internal set; }
        [JsonProperty("created_utc")] public DateTime CreatedUtc { get; internal set; }
        [JsonProperty("actor")] public string Actor { get; internal set; }
        [JsonProperty("change_reason")] public string ChangeReason { get; internal set; }
        [JsonProperty("action")] public string Action { get; internal set; }
        [JsonProperty("source_revision")] public int? SourceRevision { get; internal set; }
    }
}
