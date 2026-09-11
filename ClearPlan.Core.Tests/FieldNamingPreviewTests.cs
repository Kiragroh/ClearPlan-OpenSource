using System.Collections.Generic;
using System.Linq;
using System;
using System.IO;
using ClearPlan.Core.Fields;

namespace ClearPlan.Core.Tests
{
    internal static class FieldNamingPreviewTests
    {
        public static void ConfiguredDefaultsPreserveEstablishedNamesAndIds()
        {
            var rules = FieldNamingRuleConfiguration.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(FieldNamingRules.CreateDefault()));
            TestAssert.Equal("available", rules.Status);
            var first = BeamNamingInput.Arc("7GA01", 1, 120, 30, "CounterClockwise");
            first.PatientSupportAngle = 300;
            var second = BeamNamingInput.Arc("7GA02", 2, 120, 30, "CounterClockwise");
            second.PatientSupportAngle = 300;
            var inputs = new[] { first, second, BeamNamingInput.Static("7GA03", 3, 90), BeamNamingInput.Static("7GA04", 4, 90) };
            var existing = FieldNameSuggester.Suggest("7GA_plan", inputs);
            var configured = FieldNameSuggester.Suggest("7GA_plan", inputs, rules.Rules);
            TestAssert.Equal("120-30 T300 GUZa", configured[0].SuggestedName);
            TestAssert.Equal("90 UZa", configured[2].SuggestedName);
            for (int index = 0; index < inputs.Length; index++)
            {
                TestAssert.Equal(existing[index].SuggestedName, configured[index].SuggestedName);
                TestAssert.Equal(existing[index].ExpectedId, configured[index].ExpectedId);
                TestAssert.True(configured[index].IsEvaluated);
            }
        }

        public static void ConfiguredOrderTokensAndDuplicateDirectionAreApplied()
        {
            var rules = FieldNamingRules.CreateDefault();
            rules.ArcOrder = new List<string> { "direction", "angles", "table" };
            rules.StaticOrder = new List<string> { "table", "angles" };
            rules.PartSeparator = "_";
            rules.AngleSeparator = ":";
            rules.TablePrefix = "Couch";
            rules.ClockwiseToken = "CW";
            rules.CounterClockwiseToken = "CCW";
            rules.StaticDuplicateToken = "DUP";
            var inputs = new[] { BeamNamingInput.Arc("A01", 1, 179, 181, "Clockwise"),
                BeamNamingInput.Arc("A02", 2, 179, 181, "Clockwise"),
                BeamNamingInput.Arc("A03", 3, 181, 179, "CounterClockwise"),
                BeamNamingInput.Static("A04", 4, 90), BeamNamingInput.Static("A05", 5, 90) };
            foreach (var beam in inputs) beam.PatientSupportAngle = 30;
            var names = FieldNameSuggester.Suggest("A_plan", inputs, rules);
            TestAssert.Equal("CWa_179:181_Couch30", names[0].SuggestedName);
            TestAssert.Equal("CWb_179:181_Couch30", names[1].SuggestedName);
            TestAssert.Equal("CCW_181:179_Couch30", names[2].SuggestedName);
            TestAssert.Equal("Couch30_90_DUPa", names[3].SuggestedName);
            TestAssert.Equal("Couch30_90_DUPb", names[4].SuggestedName);
            TestAssert.Equal("A01", names[0].ExpectedId);
        }

        public static void InvalidMissingAndDisabledRulesNeverInferConformance()
        {
            var invalidOrders = new[] {
                new List<string> { "angles", "direction" },
                new List<string> { "angles", "table", "table" },
                new List<string> { "angles", "table", "unknown" }
            };
            foreach (var order in invalidOrders)
            {
                var rules = FieldNamingRules.CreateDefault();
                rules.ArcOrder = order;
                var invalid = FieldNamingRuleConfiguration.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(rules));
                TestAssert.Equal("unavailable", invalid.Status);
                AssertNotEvaluated(invalid.Rules);
            }
            var duplicateTokens = FieldNamingRules.CreateDefault();
            duplicateTokens.CounterClockwiseToken = duplicateTokens.ClockwiseToken;
            TestAssert.False(duplicateTokens.IsValid());
            AssertNotEvaluated(duplicateTokens);
            var disabled = FieldNamingRules.CreateDefault();
            disabled.Enabled = false;
            var configuration = FieldNamingRuleConfiguration.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(disabled));
            TestAssert.Equal("not-configured", configuration.Status);
            AssertNotEvaluated(configuration.Rules);
            foreach (string text in new[] { "{}", "not-json", "null" })
            {
                var invalid = FieldNamingRuleConfiguration.Parse(text);
                TestAssert.Equal("unavailable", invalid.Status);
                AssertNotEvaluated(invalid.Rules);
            }
            string missingPath = Path.Combine(Path.GetTempPath(), "clearplan-naming-missing-" + Guid.NewGuid().ToString("N") + ".json");
            var missing = FieldNamingRuleConfiguration.Load(missingPath);
            TestAssert.Equal("not-configured", missing.Status);
            TestAssert.False(File.Exists(missingPath));
            AssertNotEvaluated(missing.Rules);
            TestAssert.Equal("not-configured", FieldNamingRuleConfiguration.Load("").Status);
        }

        public static void FieldNamingConfigurationPathRoundTripsIni()
        {
            var settings = new ClearPlan.Core.Settings.ClearPlanSettingsModel();
            settings.Paths.FieldNamingRulesJsonPath = "Config\\FieldNamingRules.json";
            var copy = new ClearPlan.Core.Settings.ClearPlanSettingsModel();
            ClearPlan.Core.Settings.PathSettingsIni.Apply(copy, ClearPlan.Core.Settings.PathSettingsIni.Serialize(settings));
            TestAssert.Equal(settings.Paths.FieldNamingRulesJsonPath, copy.Paths.FieldNamingRulesJsonPath);
        }

        private static void AssertNotEvaluated(FieldNamingRules rules)
        {
            var beam = BeamNamingInput.Static("7GA01", 1, 90);
            beam.CurrentName = "90";
            var suggestion = FieldNameSuggester.Suggest("7GA_plan", new[] { beam }, rules).Single();
            TestAssert.False(suggestion.IsEvaluated);
            TestAssert.Equal("", suggestion.ExpectedId);
            TestAssert.Equal("", suggestion.SuggestedName);
            TestAssert.False(suggestion.WouldChange);
            TestAssert.True(suggestion.EvaluationMessage.Contains("no fallback"));
        }

        public static void NamesStaticAndArcFields()
        {
            IList<FieldNameSuggestion> suggestions = FieldNameSuggester.Suggest(
                new[]
                {
                    BeamNamingInput.Static("F1", 1, 90.2),
                    BeamNamingInput.Arc("F2", 2, 179.0, 181.0, "Clockwise"),
                    BeamNamingInput.Arc("F3", 3, 181.0, 179.0, "CounterClockwise")
                });

            TestAssert.Equal("90", suggestions[0].SuggestedName);
            TestAssert.Equal("179-181 UZ", suggestions[1].SuggestedName);
            TestAssert.Equal("181-179 GUZ", suggestions[2].SuggestedName);
        }

        public static void InsertsCouchAngleBeforeDirection()
        {
            BeamNamingInput input = BeamNamingInput.Arc(
                "F1",
                1,
                179,
                181,
                "Clockwise");
            input.PatientSupportAngle = 30;

            FieldNameSuggestion suggestion = FieldNameSuggester
                .Suggest(new[] { input })
                .Single();

            TestAssert.Equal("179-181 T30 UZ", suggestion.SuggestedName);
        }

        public static void KeepsCouchAnglePenultimateForDuplicateArcs()
        {
            BeamNamingInput first = BeamNamingInput.Arc(
                "F1",
                1,
                179,
                181,
                "Clockwise");
            first.PatientSupportAngle = 30;
            BeamNamingInput second = BeamNamingInput.Arc(
                "F2",
                2,
                179,
                181,
                "Clockwise");
            second.PatientSupportAngle = 30;

            IList<FieldNameSuggestion> suggestions =
                FieldNameSuggester.Suggest(new[] { first, second });

            TestAssert.Equal("179-181 T30 UZa", suggestions[0].SuggestedName);
            TestAssert.Equal("179-181 T30 UZb", suggestions[1].SuggestedName);
        }

        public static void AddsAttachedDuplicateSuffixes()
        {
            IList<FieldNameSuggestion> suggestions = FieldNameSuggester.Suggest(
                new[]
                {
                    BeamNamingInput.Arc("A", 10, 179, 181, "Clockwise"),
                    BeamNamingInput.Arc("B", 20, 179, 181, "Clockwise"),
                    BeamNamingInput.Static("C", 30, 90),
                    BeamNamingInput.Static("D", 40, 90)
                });

            TestAssert.Equal("179-181 UZa", suggestions[0].SuggestedName);
            TestAssert.Equal("179-181 UZb", suggestions[1].SuggestedName);
            TestAssert.Equal("90 UZa", suggestions[2].SuggestedName);
            TestAssert.Equal("90 UZb", suggestions[3].SuggestedName);
        }

        public static void UsesTreatmentOrderThenBeamNumber()
        {
            BeamNamingInput first = BeamNamingInput.Static("First", 90, 90);
            first.TreatmentOrderIndex = 0;
            BeamNamingInput second = BeamNamingInput.Static("Second", 2, 90);
            second.TreatmentOrderIndex = 1;
            BeamNamingInput fallback = BeamNamingInput.Static("Fallback", 1, 180);

            IList<FieldNameSuggestion> suggestions = FieldNameSuggester.Suggest(
                new[] { fallback, second, first });

            TestAssert.Equal("First", suggestions[0].CurrentId);
            TestAssert.Equal("Second", suggestions[1].CurrentId);
            TestAssert.Equal("Fallback", suggestions[2].CurrentId);
            TestAssert.Equal("90 UZa", suggestions[0].SuggestedName);
            TestAssert.Equal("90 UZb", suggestions[1].SuggestedName);
        }

        public static void UsesHalfDegreeTolerance()
        {
            BeamNamingInput nearStatic = BeamNamingInput.Arc(
                "Near",
                1,
                0,
                0.5,
                "Clockwise");
            BeamNamingInput arc = BeamNamingInput.Arc(
                "Arc",
                2,
                0,
                0.51,
                "Clockwise");

            IList<FieldNameSuggestion> suggestions =
                FieldNameSuggester.Suggest(new[] { nearStatic, arc });

            TestAssert.Equal("0", suggestions[0].SuggestedName);
            TestAssert.Equal("0-1 UZ", suggestions[1].SuggestedName);
        }

        public static void UsesPlanIdForExpectedFieldIds()
        {
            BeamNamingInput first = BeamNamingInput.Static("TMP 1", 8, 120);
            first.CurrentName = "120";
            first.TreatmentOrderIndex = 0;
            BeamNamingInput second = BeamNamingInput.Static("TMP 2", 9, 240);
            second.CurrentName = "240";
            second.TreatmentOrderIndex = 1;

            IList<FieldNameSuggestion> suggestions = FieldNameSuggester.Suggest(
                "7GA_HI9-15_20",
                new[] { first, second });

            TestAssert.Equal("7GA01", suggestions[0].ExpectedId);
            TestAssert.Equal("7GA02", suggestions[1].ExpectedId);
            TestAssert.True(suggestions[0].IdWouldChange);
            TestAssert.False(suggestions[0].NameWouldChange);
        }

        public static void PreservesConformingPrefixFromCurrentFields()
        {
            BeamNamingInput first = BeamNamingInput.Static("7GA01", 1, 120);
            first.CurrentName = "120";
            first.TreatmentOrderIndex = 0;
            BeamNamingInput second = BeamNamingInput.Static("7GA02", 2, 240);
            second.CurrentName = "240";
            second.TreatmentOrderIndex = 1;

            IList<FieldNameSuggestion> suggestions = FieldNameSuggester.Suggest(
                "Different_Plan",
                new[] { first, second });

            TestAssert.Equal("7GA01", suggestions[0].ExpectedId);
            TestAssert.Equal("7GA02", suggestions[1].ExpectedId);
            TestAssert.False(suggestions[0].IdWouldChange);
            TestAssert.False(suggestions[1].IdWouldChange);
        }

        public static void ReportsIdAndNameConformanceSeparately()
        {
            BeamNamingInput input = BeamNamingInput.Arc(
                "7GA01",
                1,
                120,
                30,
                "CounterClockwise");
            input.CurrentName = "120-30 GUZ T300";
            input.PatientSupportAngle = 300;

            FieldNameSuggestion suggestion = FieldNameSuggester.Suggest(
                "7GA_HI9-15_20",
                new[] { input })
                .Single();

            TestAssert.False(suggestion.IdWouldChange);
            TestAssert.True(suggestion.NameWouldChange);
            TestAssert.True(suggestion.WouldChange);
            TestAssert.Equal("120-30 T300 GUZ", suggestion.SuggestedName);
        }
    }
}
