using System.Collections.Generic;

namespace ClearPlan.Core.Simulation
{
    public sealed class SyntheticScenarioExpectation
    {
        public SyntheticScenarioExpectation()
        {
            RequiredStructureIds = new List<string>();
        }

        public string ScenarioId { get; set; }

        public int Seed { get; set; }

        public int SourceCount { get; set; }

        public int PqmRowCount { get; set; }

        public int PlanCheckRowCount { get; set; }

        public int FieldRowCount { get; set; }

        public int StructureMappingCount { get; set; }

        public int DvhSeriesCount { get; set; }

        public IList<string> RequiredStructureIds { get; set; }
    }
}
