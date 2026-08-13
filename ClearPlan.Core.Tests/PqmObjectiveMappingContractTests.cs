using System;
using System.Collections.Generic;
using ClearPlan.Core.Constraints;

namespace ClearPlan.Core.Tests
{
    internal static class PqmObjectiveMappingContractTests
    {
        public static void MapsNormalizedConstraintToLegacyPqmContract()
        {
            var constraint = new ConstraintDefinition
            {
                ConstraintId = "cord-dmax",
                TableId = "conventional",
                StructureId = "spinal-cord",
                StructureName = "SpinalCord",
                RawStructureName = "Myelon",
                DvhObjective = "Max[Gy]",
                EvaluationPoint = "<=45",
                Comparator = "<=",
                Goal = 45m,
                Variation = 50m,
                Priority = 1,
                Source = "RefDB",
                Comment = "Serial organ"
            };
            constraint.Aliases = new List<string> { "Myelon", "Spinalkanal" };
            constraint.Codes = new List<string> { "T-A7010" };
            constraint.DicomTypes = new List<string> { "ORGAN" };

            PqmObjectiveDefinition objective = PqmObjectiveMapper.Map(constraint);

            TestAssert.Equal("SpinalCord", objective.TemplateId);
            TestAssert.Equal("Max[Gy]", objective.DvhObjective);
            TestAssert.Equal("<=45", objective.Goal);
            TestAssert.Equal("50", objective.Variation);
            TestAssert.Equal("1", objective.Priority);
            TestAssert.Equal("Myelon", objective.Aliases[0]);
            TestAssert.Equal("Spinalkanal", objective.Aliases[1]);
            TestAssert.Equal("T-A7010", objective.Codes[0]);
            TestAssert.Equal("ORGAN", objective.DicomTypes[0]);
            TestAssert.Equal("RefDB", objective.Source);
            TestAssert.Equal("Serial organ", objective.Comment);
        }

        public static void SuppliesSafeIdentityFallbacks()
        {
            var constraint = new ConstraintDefinition
            {
                ConstraintId = "target-coverage",
                StructureId = "target",
                RawStructureName = "Zielvolumen",
                Metric = "D95%",
                Unit = "%",
                Comparator = ">=",
                Goal = 95m
            };

            PqmObjectiveDefinition objective = PqmObjectiveMapper.Map(constraint);

            TestAssert.Equal("Zielvolumen", objective.TemplateId);
            TestAssert.Equal("D95%[%]", objective.DvhObjective);
            TestAssert.Equal(">=95", objective.Goal);
            TestAssert.Equal("Zielvolumen", objective.Aliases[0]);
            TestAssert.Equal("Zielvolumen", objective.Codes[0]);
            TestAssert.Equal("Zielvolumen", objective.DicomTypes[0]);
        }

        public static void RejectsMissingConstraint()
        {
            TestAssert.Throws<ArgumentNullException>(
                () => PqmObjectiveMapper.Map(null));
        }
    }
}
