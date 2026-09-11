using System;
using System.Collections.Generic;
using System.IO;
using ClearPlan.Core.Constraints;
using ClearPlan.Core.Review;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Core.Tests
{
    internal static class StructureMappingSafetyTests
    {
        public static void UnmatchedOrgansNeverUseBroadDicomTypes()
        {
            foreach (string requested in new[] { "A_Carotid_L/R", "Cochlea", "Lens" })
            {
                foreach (string candidateId in new[] { "Brain-PTV", "PTV03", "auxiliaryPTV" })
                {
                    var definition = new StructureDefinition
                    {
                        StructureId = "synthetic-oar", CanonicalName = requested,
                        DicomTypes = { "ORGAN" }, Active = true
                    };
                    StructureMatch match = StructureAliasResolver.Resolve(
                        new ConstraintDefinition { StructureId = definition.StructureId, StructureName = requested },
                        new[] { definition },
                        new[] { new StructureCandidate { Id = candidateId, DicomType = "ORGAN" } });
                    TestAssert.Equal(null, match.CandidateId,
                        "A broad DICOM category must not assign " + requested + " to " + candidateId + ".");
                    TestAssert.Equal(StructureMatchRule.None, match.Rule);
                }
            }
        }

        public static void AmbiguousAliasesNeverChooseTheFirstCandidate()
        {
            var definition = new StructureDefinition
            {
                CanonicalName = "SpinalCord", Aliases = { "Cord", "Myelon" }, Active = true
            };
            var candidates = new[]
            {
                new StructureCandidate { Id = "Cord" }, new StructureCandidate { Id = "Myelon" }
            };
            foreach (var ordered in new[] { candidates, new[] { candidates[1], candidates[0] } })
            {
                StructureMatch match = StructureAliasResolver.Resolve(
                    new ConstraintDefinition { StructureName = "SpinalCord" }, new[] { definition }, ordered);
                TestAssert.True(match.IsAmbiguous);
                TestAssert.Equal(null, match.CandidateId);
            }
        }

        public static void AmbiguousDefinitionsNeverChooseTheFirstDefinition()
        {
            var definitions = new[]
            {
                new StructureDefinition { StructureId = "1", CanonicalName = "SpinalCord", Aliases = { "SharedAlias" } },
                new StructureDefinition { StructureId = "2", CanonicalName = "Brainstem", Aliases = { "SharedAlias" } }
            };
            StructureMatch match = StructureAliasResolver.Resolve(
                new ConstraintDefinition { StructureName = "SharedAlias" }, definitions,
                new[] { new StructureCandidate { Id = "SpinalCord" }, new StructureCandidate { Id = "Brainstem" } });
            TestAssert.Equal(null, match.CandidateId,
                "Conflicting alias definitions require review rather than an alphabetical winner.");
        }

        public static void ExplicitAliasesAndCodesStillResolve()
        {
            var definition = new StructureDefinition
            {
                StructureId = "cord", CanonicalName = "SpinalCord", Aliases = { "Myelon" },
                Codes = { "synthetic-specific-cord-code" }, Active = true
            };
            var constraint = new ConstraintDefinition { StructureId = "cord", StructureName = "SpinalCord" };
            StructureMatch alias = StructureAliasResolver.Resolve(constraint, new[] { definition },
                new[] { new StructureCandidate { Id = "Myelon", DicomType = "ORGAN" } });
            TestAssert.Equal("Myelon", alias.CandidateId);
            TestAssert.Equal(StructureMatchRule.Alias, alias.Rule);
            StructureMatch code = StructureAliasResolver.Resolve(constraint, new[] { definition },
                new[] { new StructureCandidate { Id = "LocalCord", Codes = { "synthetic-specific-cord-code" } } });
            TestAssert.Equal("LocalCord", code.CandidateId);
            TestAssert.Equal(StructureMatchRule.Code, code.Rule);
            StructureMatch unrelated = StructureAliasResolver.Resolve(constraint, new[] { definition },
                new[] { new StructureCandidate { Id = "Myelon-PTV" } });
            TestAssert.Equal(null, unrelated.CandidateId, "A cropped structure is not an exact configured alias.");
        }

        public static void MappingAvailabilityIsNotGoalAchievement()
        {
            var mapped = new ReviewStructureMappingViewModel(new ReviewStructureMapping
            {
                SelectedStructureId = "Cord", Status = ReviewStatusCodes.Pass,
                AvailableStructureIds = new List<string> { "Cord" }
            });
            TestAssert.Equal("Zugeordnet", mapped.StatusText);
            var unresolved = new ReviewStructureMappingViewModel(new ReviewStructureMapping
            {
                SelectedStructureId = string.Empty, Status = ReviewStatusCodes.NotEvaluated,
                AvailableStructureIds = new List<string> { "Unrelated" }
            });
            TestAssert.Equal(string.Empty, unresolved.SelectedStructureId,
                "The editor must not select its first available structure by default.");
            TestAssert.Equal("Nicht zugeordnet", unresolved.StatusText);
            unresolved.MarkApplied("Unrelated", "Explicit review mapping");
            TestAssert.Equal("Zugeordnet", unresolved.StatusText);
        }

        public static void LegacyMappingUsesTheSameFailClosedResolver()
        {
            string root = FindRepositoryRoot();
            string calculator = File.ReadAllText(Path.Combine(root, "ClearPlan.Script", "Calculators", "PQMSummaryCalculator.cs"));
            int begin = calculator.IndexOf("public Structure FindStructureFromAlias", StringComparison.Ordinal);
            int end = calculator.IndexOf("void ConvertUnitToGy", begin, StringComparison.Ordinal);
            string legacyResolver = calculator.Substring(begin, end - begin);
            TestAssert.True(legacyResolver.Contains("StructureAliasResolver.Resolve("),
                "Every legacy mapping entry point must use the deterministic resolver.");
            TestAssert.False(legacyResolver.Contains("StartsWith("), "Prefix and empty-alias matching are unsafe.");
            TestAssert.False(legacyResolver.Contains("PlanningItemTargetId"), "Plan target is not a substitute for a requested structure.");
            string main = File.ReadAllText(Path.Combine(root, "ClearPlan.Script", "ViewModels", "MainViewModel.cs"));
            begin = main.IndexOf(" AddPQMSummary(", StringComparison.Ordinal);
            end = main.IndexOf(" GetPlanningItemSummary(", begin, StringComparison.Ordinal);
            string append = main.Substring(begin, end - begin);
            TestAssert.False(append.Contains("TargetVolumeID"), "An exact PTV selection must not be replaced by the plan target.");
        }

        private static string FindRepositoryRoot()
        {
            var directory = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (directory != null)
            {
                if (Directory.Exists(Path.Combine(directory.FullName, "ClearPlan.Script"))) return directory.FullName;
                directory = directory.Parent;
            }
            throw new InvalidOperationException("Repository root not found for mapping contract check.");
        }
    }
}
