using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ClearPlan.Core.Constraints
{
    public static class StructureAliasResolver
    {
        public static StructureMatch Resolve(
            ConstraintDefinition constraint,
            IEnumerable<StructureDefinition> definitions,
            IEnumerable<StructureCandidate> candidates)
        {
            IList<StructureMatch> matches = ResolveAll(constraint, definitions, candidates);
            if (matches.Count == 0)
            {
                return new StructureMatch
                {
                    Rule = StructureMatchRule.None,
                    Confidence = 0,
                    Message = "No deterministic structure match was found."
                };
            }

            if (matches.Count == 1)
            {
                return matches[0];
            }

            return new StructureMatch
            {
                Rule = matches[0].Rule,
                Confidence = matches[0].Confidence,
                IsAmbiguous = true,
                CandidateIds = matches
                    .Select(match => match.CandidateId)
                    .OrderBy(id => id, StringComparer.OrdinalIgnoreCase)
                    .ToList(),
                Message = "Multiple structures match with equal confidence."
            };
        }

        public static IList<StructureMatch> ResolveAll(
            ConstraintDefinition constraint,
            IEnumerable<StructureDefinition> definitions,
            IEnumerable<StructureCandidate> candidates)
        {
            if (constraint == null)
            {
                return new List<StructureMatch>();
            }

            List<StructureDefinition> definitionList = (definitions ??
                Enumerable.Empty<StructureDefinition>()).Where(item => item != null).ToList();
            List<StructureCandidate> candidateList = (candidates ??
                Enumerable.Empty<StructureCandidate>())
                .Where(item => item != null && !string.IsNullOrWhiteSpace(item.Id))
                .ToList();
            StructureDefinition definition = FindDefinition(constraint, definitionList);
            string requestedName = FirstNonBlank(
                constraint.StructureName,
                constraint.RawStructureName,
                definition == null ? null : definition.CanonicalName);

            var ranked = new List<RankedCandidate>();
            AddNameMatches(
                ranked,
                candidateList,
                requestedName == null ? new string[0] : new[] { requestedName },
                StructureMatchRule.ExactRequestedName,
                100);

            if (definition != null)
            {
                AddNameMatches(
                    ranked,
                    candidateList,
                    new[] { definition.CanonicalName },
                    StructureMatchRule.CanonicalName,
                    90);
                AddNameMatches(
                    ranked,
                    candidateList,
                    definition.Aliases,
                    StructureMatchRule.Alias,
                    80);

                IEnumerable<string> sideAliases = SelectSideAliases(definition, requestedName);
                AddNameMatches(
                    ranked,
                    candidateList,
                    sideAliases,
                    StructureMatchRule.SideAlias,
                    70);

                bool sideSpecificRequest = RequestsOneSide(definition, requestedName);
                if (!sideSpecificRequest)
                {
                    AddCodeMatches(ranked, candidateList, definition.Codes);
                    // ORGAN, PTV, etc. describe a category, not anatomical identity.
                    // They must never resolve a missing named structure automatically.
                }
            }

            if (ranked.Count == 0)
            {
                return new List<StructureMatch>();
            }

            int bestConfidence = ranked.Max(item => item.Confidence);
            return ranked
                .Where(item => item.Confidence == bestConfidence)
                .GroupBy(item => item.Candidate.Id, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .OrderBy(item => item.Candidate.Id, StringComparer.OrdinalIgnoreCase)
                .Select(item => new StructureMatch
                {
                    CandidateId = item.Candidate.Id,
                    CandidateIds = new List<string> { item.Candidate.Id },
                    Rule = item.Rule,
                    Confidence = item.Confidence,
                    IsAmbiguous = false,
                    Message = item.Message
                })
                .ToList();
        }

        public static string NormalizeName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var builder = new StringBuilder();
            foreach (char character in value.Normalize(NormalizationForm.FormKC))
            {
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(character);
                if (category == UnicodeCategory.UppercaseLetter ||
                    category == UnicodeCategory.LowercaseLetter ||
                    category == UnicodeCategory.TitlecaseLetter ||
                    category == UnicodeCategory.DecimalDigitNumber ||
                    "~-+^!=/".IndexOf(character) >= 0)
                {
                    // TG-263 markers encode partial, cropped, combined and qualified structures.
                    // Dropping them can turn a different anatomical/evaluation volume into an exact match.
                    builder.Append(char.ToLowerInvariant(character));
                }
            }

            return builder.ToString();
        }

        private static StructureDefinition FindDefinition(
            ConstraintDefinition constraint,
            IList<StructureDefinition> definitions)
        {
            if (!string.IsNullOrWhiteSpace(constraint.StructureId))
            {
                List<StructureDefinition> byId = definitions.Where(
                    item => string.Equals(
                        item.StructureId,
                        constraint.StructureId,
                        StringComparison.OrdinalIgnoreCase)).ToList();
                if (byId.Count > 0)
                {
                    return byId.Count == 1 ? byId[0] : null;
                }
            }

            string requested = NormalizeName(FirstNonBlank(
                constraint.StructureName,
                constraint.RawStructureName));
            if (string.IsNullOrWhiteSpace(requested))
            {
                return null;
            }

            var ranked = definitions
                .Select(item => new
                {
                    Definition = item,
                    Rank = DefinitionRank(item, requested)
                })
                .Where(item => item.Rank > 0)
                .OrderByDescending(item => item.Rank)
                .ToList();
            if (ranked.Count == 0 ||
                (ranked.Count > 1 && ranked[0].Rank == ranked[1].Rank))
            {
                return null;
            }

            return ranked[0].Definition;
        }

        private static int DefinitionRank(StructureDefinition definition, string requested)
        {
            if (NormalizeName(definition.CanonicalName) == requested)
            {
                return 4;
            }

            if (ContainsNormalized(definition.Aliases, requested))
            {
                return 3;
            }

            if (ContainsNormalized(definition.SideAliasesLeft, requested) ||
                ContainsNormalized(definition.SideAliasesRight, requested))
            {
                return 2;
            }

            return NormalizeName(definition.StructureId) == requested ? 1 : 0;
        }

        private static IEnumerable<string> SelectSideAliases(
            StructureDefinition definition,
            string requestedName)
        {
            if (string.Equals(definition.Laterality, "R+L", StringComparison.OrdinalIgnoreCase))
            {
                return Enumerable.Empty<string>();
            }

            string requested = NormalizeName(requestedName);
            bool requestsLeft = ContainsNormalized(definition.SideAliasesLeft, requested);
            bool requestsRight = ContainsNormalized(definition.SideAliasesRight, requested);

            if (requestsLeft && !requestsRight)
            {
                return definition.SideAliasesLeft;
            }

            if (requestsRight && !requestsLeft)
            {
                return definition.SideAliasesRight;
            }

            if (string.Equals(definition.Laterality, "L/R", StringComparison.OrdinalIgnoreCase))
            {
                return definition.SideAliasesLeft.Concat(definition.SideAliasesRight);
            }

            return Enumerable.Empty<string>();
        }

        private static bool RequestsOneSide(
            StructureDefinition definition,
            string requestedName)
        {
            string requested = NormalizeName(requestedName);
            bool requestsLeft = ContainsNormalized(definition.SideAliasesLeft, requested);
            bool requestsRight = ContainsNormalized(definition.SideAliasesRight, requested);
            return requestsLeft ^ requestsRight;
        }

        private static void AddNameMatches(
            IList<RankedCandidate> matches,
            IEnumerable<StructureCandidate> candidates,
            IEnumerable<string> expectedNames,
            StructureMatchRule rule,
            int confidence)
        {
            var normalizedNames = new HashSet<string>(
                (expectedNames ?? Enumerable.Empty<string>())
                    .Select(NormalizeName)
                    .Where(item => !string.IsNullOrWhiteSpace(item)),
                StringComparer.Ordinal);
            if (normalizedNames.Count == 0)
            {
                return;
            }

            foreach (StructureCandidate candidate in candidates)
            {
                if (normalizedNames.Contains(NormalizeName(candidate.Id)))
                {
                    matches.Add(new RankedCandidate
                    {
                        Candidate = candidate,
                        Rule = rule,
                        Confidence = confidence,
                        Message = string.Format("Matched by {0}.", rule)
                    });
                }
            }
        }

        private static void AddCodeMatches(
            IList<RankedCandidate> matches,
            IEnumerable<StructureCandidate> candidates,
            IEnumerable<string> definitionCodes)
        {
            var expected = new HashSet<string>(
                (definitionCodes ?? Enumerable.Empty<string>())
                    .Where(item => !string.IsNullOrWhiteSpace(item))
                    .Select(item => item.Trim()),
                StringComparer.OrdinalIgnoreCase);
            if (expected.Count == 0)
            {
                return;
            }

            foreach (StructureCandidate candidate in candidates)
            {
                if ((candidate.Codes ?? new List<string>()).Any(expected.Contains))
                {
                    matches.Add(new RankedCandidate
                    {
                        Candidate = candidate,
                        Rule = StructureMatchRule.Code,
                        Confidence = 60,
                        Message = "Matched by structure code."
                    });
                }
            }
        }

        private static bool ContainsNormalized(IEnumerable<string> values, string expected)
        {
            return !string.IsNullOrWhiteSpace(expected) &&
                   (values ?? Enumerable.Empty<string>())
                   .Any(value => NormalizeName(value) == expected);
        }

        private static string FirstNonBlank(params string[] values)
        {
            return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        }

        private sealed class RankedCandidate
        {
            public StructureCandidate Candidate { get; set; }
            public StructureMatchRule Rule { get; set; }
            public int Confidence { get; set; }
            public string Message { get; set; }
        }
    }
}
