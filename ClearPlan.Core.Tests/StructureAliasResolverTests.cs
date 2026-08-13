using System.Collections.Generic;
using System.Linq;
using ClearPlan.Core.Constraints;

namespace ClearPlan.Core.Tests
{
    internal static class StructureAliasResolverTests
    {
        public static void UsesDeterministicRuleOrder()
        {
            StructureDefinition definition = KidneyPair();
            var candidates = new[]
            {
                new StructureCandidate { Id = "Niere_L", DicomType = "ORGAN" },
                new StructureCandidate { Id = "Kidney_L/R", DicomType = "ORGAN" }
            };
            var constraint = new ConstraintDefinition
            {
                StructureId = "229",
                StructureName = "Kidney_L/R"
            };

            StructureMatch match = StructureAliasResolver.Resolve(
                constraint,
                new[] { definition },
                candidates);

            TestAssert.Equal("Kidney_L/R", match.CandidateId);
            TestAssert.Equal(StructureMatchRule.ExactRequestedName, match.Rule);
            TestAssert.False(match.IsAmbiguous);
        }

        public static void ProtectsLeftAndRightLaterality()
        {
            StructureDefinition definition = KidneyPair();
            var constraint = new ConstraintDefinition
            {
                StructureId = "229",
                StructureName = "Kidney_L"
            };

            StructureMatch match = StructureAliasResolver.Resolve(
                constraint,
                new[] { definition },
                new[] { new StructureCandidate { Id = "Niere_R", DicomType = "ORGAN" } });

            TestAssert.Equal(null, match.CandidateId);
            TestAssert.Equal(StructureMatchRule.None, match.Rule);
        }

        public static void ExpandsPairedButNotCombinedStructures()
        {
            StructureDefinition paired = KidneyPair();
            var pairedConstraint = new ConstraintDefinition
            {
                StructureId = "229",
                StructureName = "Kidney_L/R"
            };
            var pairCandidates = new[]
            {
                new StructureCandidate { Id = "Niere_L" },
                new StructureCandidate { Id = "Niere_R" }
            };

            IList<StructureMatch> pairMatches = StructureAliasResolver.ResolveAll(
                pairedConstraint,
                new[] { paired },
                pairCandidates);

            TestAssert.Equal(2, pairMatches.Count);
            TestAssert.True(pairMatches.Any(match => match.CandidateId == "Niere_L"));
            TestAssert.True(pairMatches.Any(match => match.CandidateId == "Niere_R"));

            var combined = new StructureDefinition
            {
                StructureId = "146",
                CanonicalName = "Lung_R+L",
                Laterality = "R+L",
                Active = true
            };
            var combinedConstraint = new ConstraintDefinition
            {
                StructureId = "146",
                StructureName = "Lung_R+L"
            };
            IList<StructureMatch> combinedMatches = StructureAliasResolver.ResolveAll(
                combinedConstraint,
                new[] { combined },
                new[]
                {
                    new StructureCandidate { Id = "Lung_L" },
                    new StructureCandidate { Id = "Lung_R" }
                });

            TestAssert.Equal(0, combinedMatches.Count);
        }

        public static void ResolvesMissingStructureIdFromRawName()
        {
            var liver = new StructureDefinition
            {
                StructureId = "95",
                CanonicalName = "Liver",
                Active = true,
                Aliases = { "Leber" }
            };
            var constraint = new ConstraintDefinition
            {
                StructureId = null,
                StructureName = null,
                RawStructureName = "Leber"
            };

            StructureMatch match = StructureAliasResolver.Resolve(
                constraint,
                new[] { liver },
                new[] { new StructureCandidate { Id = "Liver" } });

            TestAssert.Equal("Liver", match.CandidateId);
            TestAssert.Equal(StructureMatchRule.CanonicalName, match.Rule);
        }

        public static void ReportsAmbiguityInsteadOfChoosingArbitrarily()
        {
            StructureDefinition definition = KidneyPair();
            var constraint = new ConstraintDefinition
            {
                StructureId = "229",
                StructureName = "Kidney"
            };
            var candidates = new[]
            {
                new StructureCandidate { Id = "Kidney L" },
                new StructureCandidate { Id = "Kidney_L" }
            };

            StructureMatch match = StructureAliasResolver.Resolve(
                constraint,
                new[] { definition },
                candidates);

            TestAssert.True(match.IsAmbiguous);
            TestAssert.Equal(null, match.CandidateId);
            TestAssert.Equal(2, match.CandidateIds.Count);
        }

        public static void UsesCodeBeforeDicomTypeFallback()
        {
            StructureDefinition definition = KidneyPair();
            var constraint = new ConstraintDefinition
            {
                StructureId = "229",
                StructureName = "Unmatched renal name"
            };
            var candidates = new[]
            {
                new StructureCandidate
                {
                    Id = "Renal-X",
                    DicomType = "ORGAN",
                    Codes = { "T-71000" }
                },
                new StructureCandidate
                {
                    Id = "Other organ",
                    DicomType = "ORGAN"
                }
            };

            StructureMatch match = StructureAliasResolver.Resolve(
                constraint,
                new[] { definition },
                candidates);

            TestAssert.Equal("Renal-X", match.CandidateId);
            TestAssert.Equal(StructureMatchRule.Code, match.Rule);
        }

        private static StructureDefinition KidneyPair()
        {
            return new StructureDefinition
            {
                StructureId = "229",
                CanonicalName = "Kidney_L/R",
                Laterality = "L/R",
                Active = true,
                Aliases = { "Kidney" },
                SideAliasesLeft = { "Kidney_L", "Niere_L" },
                SideAliasesRight = { "Kidney_R", "Niere_R" },
                Codes = { "T-71000" },
                DicomTypes = { "ORGAN" }
            };
        }
    }
}
