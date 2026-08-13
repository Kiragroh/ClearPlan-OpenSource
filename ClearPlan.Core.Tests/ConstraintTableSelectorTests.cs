using System.Collections.Generic;
using System.Linq;
using ClearPlan.Core.Constraints;

namespace ClearPlan.Core.Tests
{
    internal static class ConstraintTableSelectorTests
    {
        public static void ExactFractionationOutranksGenericTables()
        {
            ConstraintTableSelection result = ConstraintTableSelector.Select(
                new[]
                {
                    Table("T_Conv"),
                    Table("T_5Fx", 5, 5)
                },
                Context(5));

            TestAssert.Equal("T_5Fx", result.SelectedTable.TableId);
            TestAssert.False(result.RequiresConfirmation);
            TestAssert.True(result.Candidates[0].Reasons.Any(
                reason => reason.Contains("exact fraction")));
        }

        public static void DoseRangesRefineCompatibleCandidates()
        {
            ConstraintTableDefinition lowDose = Table("low", 5, 5);
            lowDose.DosePerFractionMinimumGy = 4m;
            lowDose.DosePerFractionMaximumGy = 5m;
            ConstraintTableDefinition highDose = Table("high", 5, 5);
            highDose.DosePerFractionMinimumGy = 7m;
            highDose.DosePerFractionMaximumGy = 8m;

            ConstraintTableSelection result = ConstraintTableSelector.Select(
                new[] { lowDose, highDose },
                new PlanConstraintContext
                {
                    FractionCount = 5,
                    DosePerFractionGy = 7.5m,
                    TotalDoseGy = 37.5m
                });

            TestAssert.Equal("high", result.SelectedTable.TableId);
            TestAssert.True(result.Candidates.All(candidate =>
                candidate.Table.TableId != "low"));
        }

        public static void PlanSumRequiresPlanSumTable()
        {
            ConstraintTableDefinition plan = Table("plan");
            ConstraintTableDefinition sum = Table("sum");
            sum.IsPlanSum = true;

            ConstraintTableSelection result = ConstraintTableSelector.Select(
                new[] { plan, sum },
                new PlanConstraintContext { IsPlanSum = true });

            TestAssert.Equal("sum", result.SelectedTable.TableId);
            TestAssert.Equal(1, result.Candidates.Count);
        }

        public static void StructureCoverageCannotRescueIncompatibleFractionation()
        {
            ConstraintTableDefinition wrong = Table("wrong", 3, 3);
            wrong.Constraints.Add(new ConstraintDefinition { StructureId = "A" });
            wrong.Constraints.Add(new ConstraintDefinition { StructureId = "B" });
            ConstraintTableDefinition compatible = Table("compatible", 5, 5);
            compatible.Constraints.Add(new ConstraintDefinition { StructureId = "A" });

            PlanConstraintContext context = Context(5);
            context.StructureIds.Add("A");
            context.StructureIds.Add("B");
            ConstraintTableSelection result = ConstraintTableSelector.Select(
                new[] { wrong, compatible },
                context);

            TestAssert.Equal("compatible", result.SelectedTable.TableId);
            TestAssert.False(result.Candidates.Any(candidate =>
                candidate.Table.TableId == "wrong"));
        }

        public static void TiesRequireConfirmation()
        {
            ConstraintTableSelection result = ConstraintTableSelector.Select(
                new[]
                {
                    Table("a", 5, 5),
                    Table("b", 5, 5)
                },
                Context(5));

            TestAssert.Equal(null, result.SelectedTable);
            TestAssert.True(result.RequiresConfirmation);
            TestAssert.Equal(2, result.Candidates.Count);
        }

        public static void SelectsAllNineLegacyDefaults()
        {
            var tables = new List<ConstraintTableDefinition>
            {
                Table("T_1Fx", 1, 1),
                Table("T_3Fx", 3, 3),
                Table("T_5Fx", 5, 5),
                Table("T_8Fx", 8, 8),
                Table("T_10Fx", 10, 10),
                Table("T_15Fx", 15, 15),
                Table("T_20Fx", 20, 20),
                Table("T_Conv")
            };
            ConstraintTableDefinition sum = Table("T_Conv_Sum");
            sum.IsPlanSum = true;
            tables.Add(sum);

            foreach (int fractions in new[] { 1, 3, 5, 8, 10, 15, 20 })
            {
                ConstraintTableSelection selection = ConstraintTableSelector.Select(
                    tables,
                    Context(fractions));
                TestAssert.Equal(
                    "T_" + fractions + "Fx",
                    selection.SelectedTable.TableId);
            }

            TestAssert.Equal(
                "T_Conv",
                ConstraintTableSelector.Select(tables, Context(28)).SelectedTable.TableId);
            TestAssert.Equal(
                "T_Conv_Sum",
                ConstraintTableSelector.Select(
                    tables,
                    new PlanConstraintContext { IsPlanSum = true }).SelectedTable.TableId);
        }

        private static PlanConstraintContext Context(int fractions)
        {
            return new PlanConstraintContext
            {
                FractionCount = fractions,
                IsPlanSum = false
            };
        }

        private static ConstraintTableDefinition Table(
            string id,
            int? minimumFractions = null,
            int? maximumFractions = null)
        {
            return new ConstraintTableDefinition
            {
                TableId = id,
                DisplayName = id,
                Active = true,
                FractionCountMinimum = minimumFractions,
                FractionCountMaximum = maximumFractions
            };
        }
    }
}
