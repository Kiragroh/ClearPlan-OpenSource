using System.Collections.Generic;
using System.Linq;
using ClearPlan.Core.Fields;

namespace ClearPlan.Core.Tests
{
    internal static class FieldNamingPreviewTests
    {
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
