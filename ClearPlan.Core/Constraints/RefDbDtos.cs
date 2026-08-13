using System.Collections.Generic;
using Newtonsoft.Json;

namespace ClearPlan.Core.Constraints
{
    internal sealed class RefDbRootDto
    {
        [JsonProperty("schema")]
        public string Schema { get; set; }

        [JsonProperty("structures")]
        public List<RefDbStructureDto> Structures { get; set; }

        [JsonProperty("tables")]
        public List<RefDbTableDto> Tables { get; set; }

        [JsonProperty("details")]
        public Dictionary<string, RefDbTableDto> Details { get; set; }
    }

    internal sealed class RefDbStructureDto
    {
        [JsonProperty("id")]
        public int? Id { get; set; }

        [JsonProperty("canonical_name")]
        public string CanonicalName { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; }

        [JsonProperty("laterality")]
        public string Laterality { get; set; }

        [JsonProperty("aliases")]
        public List<RefDbAliasDto> Aliases { get; set; }

        [JsonProperty("side_aliases")]
        public Dictionary<string, List<string>> SideAliases { get; set; }
    }

    internal sealed class RefDbAliasDto
    {
        [JsonProperty("alias")]
        public string Alias { get; set; }
    }

    internal sealed class RefDbTableDto
    {
        [JsonProperty("id")]
        public int? Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; }

        [JsonProperty("fx_min")]
        public object FractionCountMinimum { get; set; }

        [JsonProperty("fx_max")]
        public object FractionCountMaximum { get; set; }

        [JsonProperty("dpf_min")]
        public object DosePerFractionMinimum { get; set; }

        [JsonProperty("dpf_max")]
        public object DosePerFractionMaximum { get; set; }

        [JsonProperty("td_min")]
        public object TotalDoseMinimum { get; set; }

        [JsonProperty("td_max")]
        public object TotalDoseMaximum { get; set; }

        [JsonProperty("site")]
        public string Site { get; set; }

        [JsonProperty("regime")]
        public string Regime { get; set; }

        [JsonProperty("source_note")]
        public string SourceNote { get; set; }

        [JsonProperty("prescriptions")]
        public List<RefDbPrescriptionDto> Prescriptions { get; set; }

        [JsonProperty("constraints")]
        public List<RefDbConstraintDto> Constraints { get; set; }
    }

    internal sealed class RefDbPrescriptionDto
    {
        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; }
    }

    internal sealed class RefDbConstraintDto
    {
        [JsonProperty("id")]
        public int? Id { get; set; }

        [JsonProperty("constraint_table_id")]
        public int? TableId { get; set; }

        [JsonProperty("metric")]
        public string Metric { get; set; }

        [JsonProperty("unit")]
        public string Unit { get; set; }

        [JsonProperty("comparator")]
        public string Comparator { get; set; }

        [JsonProperty("limit_optimal")]
        public object LimitOptimal { get; set; }

        [JsonProperty("limit_maximal")]
        public object LimitMaximal { get; set; }

        [JsonProperty("priority")]
        public string Priority { get; set; }

        [JsonProperty("source")]
        public string Source { get; set; }

        [JsonProperty("comment")]
        public string Comment { get; set; }

        [JsonProperty("structure_id")]
        public int? StructureId { get; set; }

        [JsonProperty("structure_name")]
        public string StructureName { get; set; }

        [JsonProperty("oar_raw")]
        public string RawStructureName { get; set; }
    }
}
