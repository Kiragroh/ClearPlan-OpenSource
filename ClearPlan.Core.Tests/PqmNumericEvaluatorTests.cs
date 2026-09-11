using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using ClearPlan.Core.Constraints;

namespace ClearPlan.Core.Tests
{
    internal static class PqmNumericEvaluatorTests
    {
        public static void RunAll()
        {
            MinMaxMeanUsesValidatedNativeDoseUnits();
            CultureInfo original = Thread.CurrentThread.CurrentCulture;
            try
            {
                foreach (string culture in new[] { "de-DE", "en-US" })
                {
                    Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);
                    ParsesDecimalParameters();
                    EvaluatesGoalsAndVariation();
                    SeparatesMeasurementFromUnconfirmedAssessment();
                    RejectsInvalidValues();
                    ConvertsOnlyExplicitCompatibleDoseUnits();
                    PreservesDecimalColorRatio();
                    ComplementsVolumeInItsOwnUnit();
                    UsesSameUnitForComplementColorRatio();
                }
            }
            finally { Thread.CurrentThread.CurrentCulture = original; }
        }

        private static void MinMaxMeanUsesValidatedNativeDoseUnits()
        {
            string path = null;
            foreach (string start in new[] { Directory.GetCurrentDirectory(), AppDomain.CurrentDomain.BaseDirectory })
            {
                for (var directory = new DirectoryInfo(start); directory != null; directory = directory.Parent)
                {
                    string candidate = Path.Combine(directory.FullName, "ClearPlan.Script", "Calculators", "PQMMinMaxMean.cs");
                    if (File.Exists(candidate)) { path = candidate; break; }
                }
                if (path != null) break;
            }
            TestAssert.NotNull(path, "PQM source contract requires the repository checkout.");
            string source = File.ReadAllText(path);
            TestAssert.True(source.Contains("PQMDoseAtVolume.FormatDoseValue"), "Min/Max/Mean must use the validated native DoseValue formatter.");
            TestAssert.True(source.Contains("dvh.MinDose") && source.Contains("dvh.MaxDose") && source.Contains("dvh.MeanDose"));
            TestAssert.False(source.Contains("TotalDose") || source.Contains("planSumRxDose"), "Do not invent a relative PlanSum dose normalization.");
            TestAssert.True(source.Contains("is PlanSum") && source.Contains("relative dose for plan sum is unsupported"));
            TestAssert.True(source.Contains("evalunit.Value != \"cc\""), "Structure.Volume may only be labelled cc.");
            TestAssert.True(source.Contains("dvh == null") && source.Contains("double.IsNaN(dvh.SamplingCoverage)") &&
                source.Contains("double.IsInfinity(dvh.Coverage)"), "Missing or non-finite coverage must not be evaluated.");
            string tableSource = File.ReadAllText(Path.Combine(Path.GetDirectoryName(path), "PQMSummaryCalculator.cs"));
            TestAssert.False(tableSource.Contains("if (selection.RequiresConfirmation ||"),
                "An unconfirmed table must not prevent measurement of the current plan dose.");
            TestAssert.True(tableSource.Contains("objective.EvaluationRequiresConfirmation = selection.RequiresConfirmation"),
                "Constraint scope confirmation must be carried into row evaluation, including manual remapping.");
            string viewModel = File.ReadAllText(Path.Combine(Path.GetDirectoryName(path), "..", "ViewModels", "PQMSummaryViewModel.cs"));
            TestAssert.True(viewModel.Contains("EvaluateForConfirmedScope(") && viewModel.Contains("!EvaluationRequiresConfirmation"),
                "The measurement must never silently turn into a passed assessment while scope is unconfirmed.");
        }

        private static void ParsesDecimalParameters()
        {
            foreach (string value in new[] { "0.03", "0,03", ".03" })
            {
                object[] args = { value, 0.0 };
                TestAssert.Equal(true, Call("TryParseNumber", args));
                Near(0.03, (double)args[1]);
            }
        }

        private static void EvaluatesGoalsAndVariation()
        {
            TestAssert.Equal("Not met", Call("Evaluate", "21 Gy", "<=20.5", ""));
            TestAssert.Equal("Variation", Call("Evaluate", "20.6 Gy", "<=20.5", "21"));
            TestAssert.Equal("Variation", Call("Evaluate", "20,6 Gy", "<=20.5", "21"));
            TestAssert.Equal("Goal", Call("Evaluate", "20.5 Gy", "<=20.5", null));
            TestAssert.Equal("Not met", Call("Evaluate", "20.5 Gy", "<20.5", null));
            TestAssert.Equal("Goal", Call("Evaluate", "21 %", ">20.5", ""));
            TestAssert.Equal("Goal", Call("Evaluate", "20.5 cc", ">=20.5", ""));
            TestAssert.Equal("Goal", Call("Evaluate", "20.5", "=20.5", ""));
            TestAssert.Equal("Variation", Call("Evaluate", "20.4", ">=20.5", "20"));
        }

        private static void SeparatesMeasurementFromUnconfirmedAssessment()
        {
            // A partial plan still has measurable dose. Unknown assessment scope must
            // suppress both reassuring passes and misleading failures, not the value.
            TestAssert.Equal("Not evaluated", Call("EvaluateForConfirmedScope", "10 Gy", "<=20", "", false));
            TestAssert.Equal("Not evaluated", Call("EvaluateForConfirmedScope", "30 Gy", "<=20", "", false));
            TestAssert.Equal("Goal", Call("EvaluateForConfirmedScope", "10 Gy", "<=20", "", true));
            TestAssert.Equal("Not met", Call("EvaluateForConfirmedScope", "30 Gy", "<=20", "", true));
            TestAssert.Equal("Not evaluated", Call("EvaluateForConfirmedScope", "Unavailable", "<=20", "", true));
            double actual;
            TestAssert.True(PqmNumericEvaluator.TryParseAchieved("10 Gy", out actual));
            Near(10, actual);
        }

        private static void RejectsInvalidValues()
        {
            foreach (string value in new[] { null, "", "NaN", "Infinity", "-1 Gy", "1.2.3", "1,234.5", "Unable to calculate D0.03cc", "3 errors", "21 Unknown", "21 Gy extra" })
            {
                TestAssert.Equal("Not evaluated", Call("Evaluate", value, "<=20.5", "21"));
            }
            TestAssert.Equal("Not evaluated", Call("Evaluate", "20 Gy", null, "21"));
            TestAssert.Equal("Not evaluated", Call("Evaluate", "20 Gy", "<=NaN", "21"));
            TestAssert.Equal("Not evaluated", Call("Evaluate", "20 Gy", "<=20.5", "nan"));
            TestAssert.Equal("Not evaluated", Call("Evaluate", "21 Gy", "<=20.5", "bad21"));
            TestAssert.Equal("Not evaluated", Call("Evaluate", "20 Gy", ">=-1", ""));
            TestAssert.Equal("Not evaluated", Call("Evaluate", "20 Gy", "<=20.5", "-1"));
        }

        private static void ConvertsOnlyExplicitCompatibleDoseUnits()
        {
            ConvertDose(2050, "cGy", "Gy", true, 20.5);
            ConvertDose(20.5, "Gy", "cGy", true, 2050);
            ConvertDose(20.5, "Gy", "Gy", true, 20.5);
            ConvertDose(95, "%", "%", true, 95);
            ConvertDose(95, "%", "Gy", false, 0);
            ConvertDose(20.5, "Gy", "%", false, 0);
            ConvertDose(20.5, "Unknown", "Gy", false, 0);
            ConvertDose(20.5, "Unknown", "Unknown", false, 0);
            ConvertDose(double.NaN, "Gy", "Gy", false, 0);
            ConvertDose(double.PositiveInfinity, "Gy", "Gy", false, 0);
            ConvertDose(-1, "Gy", "Gy", false, 0);
        }

        private static void ComplementsVolumeInItsOwnUnit()
        {
            Complement(30, 500, "%", 90, true, 70);
            Complement(30, 500, "cc", 400, true, 470);
            Complement(30, 50, "cc", 90, false, 0);
            Complement(30, 50, "%", 90, true, 70);
            Complement(101, 500, "%", null, false, 0);
            Complement(501, 500, "cc", null, false, 0);
            Complement(-1, 500, "%", null, false, 0);
            Complement(double.NaN, 500, "cc", null, false, 0);
            Complement(30, double.NaN, "cc", null, false, 0);
            Complement(30, 500, "Gy", null, false, 0);
            Complement(30, 500, "cc", -1, false, 0);
        }

        private static void Complement(double actual, double organ, string unit, double? minimum, bool expected, double expectedValue)
        {
            object[] args = { actual, organ, unit, minimum, 0.0 };
            TestAssert.Equal(expected, Call("TryComplementVolume", args));
            if (expected) Near(expectedValue, (double)args[4]);
        }

        private static void UsesSameUnitForComplementColorRatio()
        {
            object[] relative = { "70 %", ">=60", 500.0, "%", 0.0 };
            TestAssert.Equal(true, Call("TryGetComplementGoalRatio", relative));
            Near(75, (double)relative[4]);
            object[] absolute = { "470 cc", ">=460", 500.0, "cc", 0.0 };
            TestAssert.Equal(true, Call("TryGetComplementGoalRatio", absolute));
            Near(75, (double)absolute[4]);
            TestAssert.Equal(false, Call("TryGetComplementGoalRatio", "70 %", ">=100", 500.0, "%", 0.0));
            TestAssert.Equal(false, Call("TryGetComplementGoalRatio", "70 %", ">=60", 500.0, "Gy", 0.0));
        }

        private static void PreservesDecimalColorRatio()
        {
            object[] args = { "20.6 Gy", "<=20.5", 0.0 };
            TestAssert.Equal(true, Call("TryGetGoalRatio", args));
            Near(20.6 / 20.5 * 100, (double)args[2]);
            TestAssert.Equal(false, Call("TryGetGoalRatio", "Unable to calculate D0.03cc", "<=20.5", 0.0));
            TestAssert.Equal(false, Call("TryGetGoalRatio", "1 Gy", "<=0", 0.0));
        }

        private static void ConvertDose(double value, string from, string to, bool expected, double expectedValue)
        {
            object[] args = { value, from, to, 0.0 };
            TestAssert.Equal(expected, Call("TryConvertDose", args));
            if (expected) Near(expectedValue, (double)args[3]);
        }

        private static object Call(string method, params object[] args)
        {
            Type type = typeof(ConstraintValueNormalizer).Assembly.GetType("ClearPlan.Core.Constraints.PqmNumericEvaluator");
            TestAssert.NotNull(type, "PQM must have one strict, culture-independent numeric evaluator.");
            MethodInfo target = type.GetMethod(method, BindingFlags.Public | BindingFlags.Static);
            TestAssert.NotNull(target, "Missing numeric evaluator method: " + method);
            return target.Invoke(null, args);
        }

        private static void Near(double expected, double actual)
        {
            TestAssert.True(Math.Abs(expected - actual) < 1e-9, "Expected " + expected + "; actual " + actual);
        }
    }
}
