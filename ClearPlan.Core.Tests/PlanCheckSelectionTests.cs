using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class PlanCheckSelectionTests
    {
        public static int Main()
        {
            int failures = 0;
            foreach (var test in new Action[] { RepeatedLegacyCodesAreGroupedWithoutHidingWarnings,
                MissingAndInvalidConfigurationPreserveFindings, ConfigurationContainsOnlyCodesAndFlags,
                SettingsAndVersionedReviewIntegrationExist })
            {
                try { test(); Console.WriteLine("PASS " + test.Method.Name); }
                catch (Exception error) { failures++; Console.Error.WriteLine("FAIL " + test.Method.Name + ": " + error.Message); }
            }
            return failures == 0 ? 0 : 1;
        }
        public static void RepeatedLegacyCodesAreGroupedWithoutHidingWarnings()
        {
            object config = Parse("{\"schemaVersion\":1,\"checks\":[{\"code\":\"74\",\"enabled\":false},{\"code\":\"PC001\",\"enabled\":false}]}");
            var rows = new[] { Row("74", "Default"), Row("74-2", "Default"), Row("74-10", "Default"),
                Row("75", "Default"), Row("PC001", "Metadata"), Row("native-warning-74", "Eclipse warning"),
                Row("74", "Reference point"), Row("plan-check-1", "Default") };
            int disabled;
            var kept = Filter(config, rows, out disabled);
            TestAssert.Equal(4, disabled);
            TestAssert.Equal(4, kept.Count);
            TestAssert.True(kept.Any(r => r.CheckCode == "native-warning-74"));
            TestAssert.True(kept.Any(r => r.Category == "Reference point"));
            TestAssert.True(kept.Any(r => r.CheckCode == "plan-check-1"), "Per-run fallback IDs are not stable check families.");
        }
        public static void MissingAndInvalidConfigurationPreserveFindings()
        {
            var type = ConfigurationType();
            string missing = Path.Combine(Path.GetTempPath(), "clearplan-missing-checks-" + Guid.NewGuid().ToString("N") + ".json");
            object config = type.GetMethod("Load").Invoke(null, new object[] { missing });
            TestAssert.Equal("not-configured", Read<string>(config, "Status"));
            TestAssert.False(File.Exists(missing));
            int disabled;
            TestAssert.Equal(1, Filter(config, new[] { Row("74", "Default") }, out disabled).Count);
            foreach (string invalid in new[] { "bad-json", "{}", "{\"schemaVersion\":1,\"checks\":null}",
                "{\"schemaVersion\":1,\"checks\":[{\"code\":\"74\"}]}",
                "{\"schemaVersion\":1,\"checks\":[{\"code\":\"74-2\",\"enabled\":false}]}",
                "{\"schemaVersion\":1,\"checks\":[{\"code\":\"eclipse-warning\",\"enabled\":false}]}",
                "{\"schemaVersion\":1,\"checks\":[{\"code\":\"74\",\"enabled\":false},{\"code\":\"74\",\"enabled\":true}]}" })
            {
                config = Parse(invalid);
                TestAssert.Equal("unavailable", Read<string>(config, "Status"));
                TestAssert.Equal(1, Filter(config, new[] { Row("74", "Default") }, out disabled).Count);
                TestAssert.Equal(0, disabled);
            }
        }
        public static void ConfigurationContainsOnlyCodesAndFlags()
        {
            var config = Parse("{\"schemaVersion\":1,\"checks\":[{\"code\":\"74\",\"enabled\":false,\"description\":\"DO_NOT_PERSIST_CLINICAL_TEXT\"}]}");
            TestAssert.Equal("unavailable", Read<string>(config, "Status"));
            TestAssert.NotNull(typeof(ClearPlan.Core.Settings.ClearPlanPathOptions).GetProperty("PlanCheckSelectionJsonPath"));
            TestAssert.NotNull(typeof(ReviewSnapshot).GetProperty("DisabledCheckCount"));
            var settings = new ClearPlan.Core.Settings.ClearPlanSettingsModel();
            settings.Paths.PlanCheckSelectionJsonPath = "Configuration/checks.json";
            var copy = new ClearPlan.Core.Settings.ClearPlanSettingsModel();
            ClearPlan.Core.Settings.PathSettingsIni.Apply(copy, ClearPlan.Core.Settings.PathSettingsIni.Serialize(settings));
            TestAssert.Equal(settings.Paths.PlanCheckSelectionJsonPath, copy.Paths.PlanCheckSelectionJsonPath);
            string json = PlanCheckSelectionConfiguration.Serialize(new[] { new PlanCheckSelectionEntry { Code = "74", Enabled = false } });
            TestAssert.False(json.Contains("description") || json.Contains("SYNTHETIC MEMORY ONLY"));
            TestAssert.Equal(false, PlanCheckSelectionConfiguration.Parse(json).IsEnabled("74"));
            var snapshot = new ReviewSnapshot { DisabledCheckCount = 3 };
            TestAssert.Equal(3, Newtonsoft.Json.JsonConvert.DeserializeObject<ReviewSnapshot>(
                Newtonsoft.Json.JsonConvert.SerializeObject(snapshot)).DisabledCheckCount);
        }
        public static void SettingsAndVersionedReviewIntegrationExist()
        {
            string vm = File.ReadAllText("ClearPlan.Script/ViewModels/SettingsViewModel.cs");
            TestAssert.True(vm.Contains("InitializeCheckSelection") && vm.Contains("plancheck-selection"));
            string view = File.ReadAllText("ClearPlan.Script/Views/SettingsView.xaml");
            TestAssert.True(view.Contains("PlanCheckSelectionRows") && view.Contains("DataGridCheckBoxColumn"));
            string source = File.ReadAllText("ClearPlan.Script/Review/EsapiReviewSnapshotBuilder.cs");
            TestAssert.True(source.Contains("DisabledCheckCount") && source.Contains("PlanCheckSelectionConfiguration.Load"));
            TestAssert.True(source.Contains("plancheck-selection-config") && source.Contains("ReviewStatusCodes.NotEvaluated"));
        }
        private static ReviewCheckRow Row(string code, string category)
        { return new ReviewCheckRow { CheckCode = code, Category = category, Status = "fail", Message = "SYNTHETIC MEMORY ONLY" }; }
        private static Type ConfigurationType()
        {
            var type = typeof(ReviewSnapshot).Assembly.GetType("ClearPlan.Core.Review.PlanCheckSelectionConfiguration");
            TestAssert.NotNull(type, "Missing bounded review-inclusion policy.");
            return type;
        }
        private static object Parse(string json) { return ConfigurationType().GetMethod("Parse").Invoke(null, new object[] { json }); }
        private static List<ReviewCheckRow> Filter(object config, IEnumerable<ReviewCheckRow> rows, out int disabled)
        {
            var args = new object[] { rows, 0 };
            var result = (List<ReviewCheckRow>)config.GetType().GetMethod("Filter").Invoke(config, args);
            disabled = (int)args[1]; return result;
        }
        private static T Read<T>(object source, string property)
        { return (T)source.GetType().GetProperty(property).GetValue(source, null); }
    }
}
