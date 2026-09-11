using System.Collections.Generic;
using System.Linq;
using ClearPlan.Core.Constraints;

namespace ClearPlan.Core.Tests
{
    internal static class ConstraintTableSelectorTests
    {
        public static void MostDistinctHitsWinWithinMatchingFractions()
        {
            var sparse = Table("sparse", 5, 5);
            sparse.Constraints.Add(new ConstraintDefinition { StructureName = "A" });
            var broad = Table("broad", 5, 5);
            foreach (var id in new[] { "A", "B", "C", "D", "E", "F", "G", "H" })
                broad.Constraints.Add(new ConstraintDefinition { StructureName = id });
            var context = Context(5); context.StructureIds = new List<string> { "A", "B", "C" };
            var result = ConstraintTableSelector.Select(new[] { sparse, broad }, context);
            TestAssert.Equal("broad", result.Candidates[0].Table.TableId,
                "Three distinct hits must outrank one hit, even when that one covers 100% of a smaller table.");
            TestAssert.Equal("broad", result.SelectedTable.TableId);
            for (int i = 0; i < 20; i++) sparse.Constraints.Add(new ConstraintDefinition { StructureName = "A" });
            result = ConstraintTableSelector.Select(new[] { sparse, broad }, context);
            TestAssert.Equal("broad", result.Candidates[0].Table.TableId, "Repeated metrics for one organ are not additional hits.");
        }

        public static void AliasesCountWithoutOverridingFractionOrConfirmation()
        {
            var one = Table("one", 5, 5); one.Constraints.Add(new ConstraintDefinition { StructureName = "A" });
            one.PrescriptionLabels.Add("Linked name");
            var two = Table("two", 5, 5);
            two.Constraints.Add(new ConstraintDefinition { StructureId = "organ-b", StructureName = "B" });
            two.Constraints.Add(new ConstraintDefinition { StructureName = "C" });
            var wrong = Table("wrong", 3, 3);
            foreach (var id in new[] { "A", "B_local", "C", "D" }) wrong.Constraints.Add(new ConstraintDefinition { StructureName = id });
            var context = Context(5); context.StructureIds = new List<string> { "A", "B_local", "C", "D" };
            context.PrescriptionLabels.Add("Linked name");
            context.StructureDefinitions.Add(new StructureDefinition { StructureId = "organ-b", CanonicalName = "B", Active = true, Aliases = new List<string> { "B_local" } });
            var result = ConstraintTableSelector.Select(new[] { one, two, wrong }, context);
            TestAssert.Equal("two", result.SelectedTable.TableId, "Most matching structures take priority over a label bonus, within valid Fx scope.");
            TestAssert.Equal(2, result.Candidates[0].StructureHits);
            TestAssert.False(result.Candidates.Any(c => c.Table == wrong));
            two.RequiresConfirmation = true;
            result = ConstraintTableSelector.Select(new[] { one, two }, context);
            TestAssert.True(result.RequiresConfirmation && result.SelectedTable == null);
            TestAssert.Equal("two", result.Candidates[0].Table.TableId);
        }

        public static void AConstraintCannotCountCanonicalAndLegacyIdTwice()
        {
            var table = Table("scope", 5, 5);
            table.Constraints.Add(new ConstraintDefinition { StructureId = "Cord" });
            var context = Context(5); context.StructureIds = new List<string> { "Cord", "SpinalCord" };
            context.StructureDefinitions.Add(new StructureDefinition { StructureId = "Cord", CanonicalName = "SpinalCord", Active = true });
            var result = ConstraintTableSelector.Select(new[] { table }, context);
            TestAssert.Equal(1, result.Candidates[0].StructureHits, "One constraint request cannot count canonical and legacy names twice.");
            table.Constraints[0].StructureName = ""; table.Constraints[0].RawStructureName = "SpinalCord";
            result = ConstraintTableSelector.Select(new[] { table }, context);
            TestAssert.Equal(1, result.Candidates[0].StructureHits, "Empty display names must not hide a valid raw requested name.");
            context.FractionCount = null;
            result = ConstraintTableSelector.Select(new[] { table }, context);
            TestAssert.True(result.RequiresConfirmation && result.SelectedTable == null, "Unknown fractions cannot establish a bounded Fx match.");
        }

        public static void StockDefinitionsCanRequireExplicitConfirmation()
        {
            var stock = Table("stock-5", 5, 5);
            var property = typeof(ConstraintTableDefinition).GetProperty("RequiresConfirmation");
            TestAssert.NotNull(property, "Stock catalogs must be able to require review before evaluation.");
            property.SetValue(stock, true, null);
            var result = ConstraintTableSelector.Select(new[] { stock }, Context(5));
            TestAssert.True(result.RequiresConfirmation && result.SelectedTable == null);
            TestAssert.Equal("stock-5", result.Candidates[0].Table.TableId);
            TestAssert.True(result.Reason.Contains("explicit confirmation"));
        }
        public static void PrescriptionLabelsDisambiguateButDoNotOverrideDose()
        {
            var linked = Table("linked", 18, 20); linked.PrescriptionLabels.Add("Synthetic linked Rx");
            var other = Table("other", 18, 20);
            var context = Context(18);
            typeof(PlanConstraintContext).GetProperty("PrescriptionLabels").SetValue(context,
                new List<string> { "synthetic linked rx" }, null);
            var result = ConstraintTableSelector.Select(new[] { linked, other }, context);
            TestAssert.Equal("linked", result.SelectedTable.TableId);
            context.FractionCount = 9;
            result = ConstraintTableSelector.Select(new[] { linked, Table("generic") }, context);
            TestAssert.True(result.SelectedTable == null, "A linked full-course prescription must not silently evaluate a partial plan.");
            TestAssert.True(result.RequiresConfirmation);
        }

        public static void AnchoredPrescriptionCodeMatchesConfiguredTableOnly()
        {
            var linked = Table("configured"); linked.DisplayName = "EX18:Example course";
            var context = Context(18); context.PrescriptionLabels.Add("EX18: Synthetic prescription description");
            var result = ConstraintTableSelector.Select(new[] { linked, Table("other") }, context);
            TestAssert.Equal("configured", result.SelectedTable.TableId);
            context.PrescriptionFractionCount = 18; context.FractionCount = 9;
            result = ConstraintTableSelector.Select(new[] { linked, Table("other") }, context);
            TestAssert.True(result.RequiresConfirmation && result.SelectedTable == null);
            TestAssert.Equal("configured", result.Candidates[0].Table.TableId);
            context.FractionCount = 18;
            context.PrescriptionLabels[0] = "prefix EX18: should not match";
            result = ConstraintTableSelector.Select(new[] { linked, Table("other") }, context);
            TestAssert.True(result.RequiresConfirmation);
        }

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
