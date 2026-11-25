using System;
using System.IO;
using System.Linq;
using System.Text;
using ClearPlan.Core.Review;
using ClearPlan.Core.Settings;

namespace ClearPlan.Core.Tests
{
    public static class CtCompatibilityTests
    {
        private const string Approved = "{\"schemaVersion\":1,\"combinations\":[{\"id\":\"synthetic-approved\",\"enabled\":true,\"manufacturer\":\"SYNTHETIC\",\"model\":\"CT Test\",\"serialNumber\":\"000123\",\"calibration\":\"HU Test\"}]}";
        private const string Legacy = "ImagingDevice-SerialNo ('000123'), -Model ('CT Test'), -Manufacturer ('SYNTHETIC') and HU-Table ('HU Test') match?";

        private static dynamic Policy(string json)
        {
            var type = typeof(ReviewSnapshot).Assembly.GetType("ClearPlan.Core.Review.CtCompatibilityConfiguration");
            TestAssert.NotNull(type, "Exact CT compatibility policy has not been implemented.");
            return type.GetMethod("Parse").Invoke(null, new object[] { json });
        }

        public static void ApprovedExactTuple()
        {
            dynamic policy = Policy(Approved);
            ReviewCheckRow row = policy.Evaluate(" synthetic ", " CT Test ", " 000123 ", " hu test ");
            TestAssert.Equal(ReviewStatusCodes.Pass, row.Status);
            TestAssert.True(row.ExpectedValue.Contains("synthetic-approved"));
            TestAssert.True(row.ObservedValue.Contains("000123"));
        }

        public static void WrongTupleNeverPasses()
        {
            dynamic policy = Policy(Approved);
            TestAssert.Equal(ReviewStatusCodes.Fail, ((ReviewCheckRow)policy.Evaluate("SYNTHETIC", "CT Test", "123", "HU Test")).Status);
            TestAssert.Equal(ReviewStatusCodes.Fail, ((ReviewCheckRow)policy.Evaluate("SYNTHETIC", "CT Test", "000123", "HU Test-extra")).Status);
            TestAssert.Equal(ReviewStatusCodes.Fail, ((ReviewCheckRow)policy.Evaluate("OTHER", "CT Test", "000123", "HU Test")).Status);
            TestAssert.Equal(ReviewStatusCodes.Fail, ((ReviewCheckRow)policy.Evaluate("SYNTHETIC", "CT  Test", "000123", "HU Test")).Status);
        }

        public static void MissingDisabledAndUnconfiguredNeverPass()
        {
            foreach (string json in new[] { null, "{}", "{\"schemaVersion\":1,\"combinations\":[]}", Approved.Replace("true", "false") })
            {
                dynamic policy = Policy(json);
                TestAssert.Equal(ReviewStatusCodes.NotEvaluated, ((ReviewCheckRow)policy.Evaluate("SYNTHETIC", "CT Test", "000123", "HU Test")).Status);
            }
            dynamic valid = Policy(Approved);
            TestAssert.Equal(ReviewStatusCodes.NotEvaluated, ((ReviewCheckRow)valid.Evaluate("SYNTHETIC", null, "000123", "HU Test")).Status);
            TestAssert.Equal(ReviewStatusCodes.NotEvaluated, ((ReviewCheckRow)valid.Evaluate("SYNTHETIC", "CT Test", "", "HU Test")).Status);
            TestAssert.Equal(ReviewStatusCodes.NotEvaluated, ((ReviewCheckRow)valid.Evaluate("SYNTHETIC", "CT Test", "000123", " ")).Status);
        }

        public static void InvalidConfigurationNeverPasses()
        {
            foreach (string json in new[] { "{", Approved.Replace("\"enabled\":true,", ""),
                Approved.Replace("\"000123\"", "\"*\""), Approved.Replace("\"HU Test\"", "\"\""),
                Approved.Replace("\"schemaVersion\":1", "\"schemaVersion\":2"), Approved.Replace("\"enabled\":true", "\"enabled\":\"true\""),
                Approved.Replace("\"serialNumber\":\"000123\"", "\"serialNumber\":123"),
                Approved.Replace("\"schemaVersion\":1", "\"schemaVersion\":1,\"schemaVersion\":1"),
                Approved.Replace("\"enabled\":true", "\"enabled\":true,\"wildcard\":true") })
            {
                dynamic policy = Policy(json);
                TestAssert.Equal(ReviewStatusCodes.Unavailable, (string)policy.Status, "Invalid configuration was accepted: " + json);
                TestAssert.Equal(ReviewStatusCodes.NotEvaluated, ((ReviewCheckRow)policy.Evaluate("SYNTHETIC", "CT Test", "000123", "HU Test")).Status);
            }
        }

        public static void LegacyBoundaryIsExact()
        {
            dynamic policy = Policy(Approved);
            ReviewCheckRow passed = policy.EvaluateLegacy("19", Legacy);
            TestAssert.Equal(ReviewStatusCodes.Pass, passed.Status);
            TestAssert.Equal(ReviewStatusCodes.Fail, ((ReviewCheckRow)policy.EvaluateLegacy("19", Legacy.Replace("000123", "000124"))).Status);
            TestAssert.Equal(ReviewStatusCodes.NotEvaluated, ((ReviewCheckRow)policy.EvaluateLegacy("19", Legacy + " extra")).Status);
            TestAssert.Equal(ReviewStatusCodes.NotEvaluated, ((ReviewCheckRow)policy.EvaluateLegacy("19", "CT appears correct")).Status);
            TestAssert.True(policy.EvaluateLegacy("20", Legacy) == null, "Unrelated check must remain untouched.");
            TestAssert.True(policy.EvaluateLegacy("19-2", Legacy) == null, "Adapter accepts raw algorithm code, not generated row IDs.");
        }

        public static void SettingsIntegration()
        {
            var property = typeof(ClearPlanPathOptions).GetProperty("CtCompatibilityJsonPath");
            TestAssert.NotNull(property, "CT compatibility path must round-trip through existing settings.");
            var settings = new ClearPlanSettingsModel();
            property.SetValue(settings.Paths, "Configuration/CtCompatibility.json");
            var roundtrip = new ClearPlanSettingsModel();
            PathSettingsIni.Apply(roundtrip, PathSettingsIni.Serialize(settings));
            TestAssert.Equal("Configuration/CtCompatibility.json", (string)property.GetValue(roundtrip.Paths));
            ClearPlan.ConfigurationContentValidator.Validate("ct-compatibility", Encoding.UTF8.GetBytes(Approved), null);
            TestAssert.Throws<FormatException>(() => ClearPlan.ConfigurationContentValidator.Validate("ct-compatibility", Encoding.UTF8.GetBytes("{}"), null));
        }

        public static void NativeAndGuiIntegration()
        {
            string root = AppDomain.CurrentDomain.BaseDirectory;
            while (!File.Exists(Path.Combine(root, "ClearPlan.sln"))) root = Directory.GetParent(root).FullName;
            string builder = File.ReadAllText(Path.Combine(root, "ClearPlan.Script/Review/EsapiReviewSnapshotBuilder.cs"));
            TestAssert.True(builder.Contains("ctCompatibility.EvaluateLegacy(item.Severity, item.Description)"), "Native projection must evaluate the raw exact legacy tuple before sanitization.");
            TestAssert.True(builder.Contains("CtCompatibilityConfiguration.Load") && builder.Contains("source-ct-compatibility"));
            string settings = File.ReadAllText(Path.Combine(root, "ClearPlan.Script/ViewModels/SettingsViewModel.cs"));
            TestAssert.True(settings.Contains("Entry(\"ct-compatibility\"") && settings.Contains("CtCompatibilityJsonPath = path"), "CT configuration must use the versioned existing GUI editor.");
            string view = File.ReadAllText(Path.Combine(root, "ClearPlan.Script/Views/SettingsView.xaml"));
            TestAssert.True(view.Contains("{Binding CtCompatibilityJsonPath"));
            dynamic generic = Policy(File.ReadAllText(Path.Combine(root, "ClearPlan.Script/Distribution/CtCompatibility.json")));
            TestAssert.Equal(ReviewStatusCodes.NotEvaluated, ((ReviewCheckRow)generic.Evaluate("SYNTHETIC", "CT Test", "000123", "HU Test")).Status);
        }

        public static void VersionHistoryAndReload()
        {
            string root = Path.Combine(Path.GetTempPath(), "clearplan-ct-tests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var entry = new ClearPlan.ConfigurationEntryViewModel { Key = "ct-compatibility", Title = "CT", Extension = ".json",
                    SourcePath = () => string.Empty, Validate = bytes => ClearPlan.ConfigurationContentValidator.Validate("ct-compatibility", bytes, root) };
                string activated = null;
                var vm = new ClearPlan.ConfigurationWorkspaceViewModel(root, new[] { entry }, (key, path) => activated = path);
                vm.StageJson("ct-compatibility", Approved);
                vm.ChangeReason = "Synthetic exact tuple approved";
                TestAssert.True(vm.SaveJson());
                string first = activated;
                dynamic policy = Load(activated);
                TestAssert.Equal(ReviewStatusCodes.Pass, ((ReviewCheckRow)policy.EvaluateLegacy("19", Legacy)).Status);
                vm.EditorText = Approved.Replace("true", "false"); vm.ChangeReason = "Disable test tuple";
                TestAssert.True(vm.SaveJson());
                policy = Load(activated);
                TestAssert.Equal(ReviewStatusCodes.NotEvaluated, ((ReviewCheckRow)policy.EvaluateLegacy("19", Legacy)).Status);
                TestAssert.Equal(Approved, File.ReadAllText(first));
                vm.SelectedVersion = vm.Versions.Last(); vm.ChangeReason = "Restore original approval";
                TestAssert.True(vm.RestoreSelected());
                policy = Load(activated);
                TestAssert.Equal(ReviewStatusCodes.Pass, ((ReviewCheckRow)policy.EvaluateLegacy("19", Legacy)).Status);
                TestAssert.Equal(3, vm.CurrentRevisionNumber);
            }
            finally { Directory.Delete(root, true); }
        }

        private static dynamic Load(string path)
        {
            var type = typeof(ReviewSnapshot).Assembly.GetType("ClearPlan.Core.Review.CtCompatibilityConfiguration");
            TestAssert.NotNull(type);
            return type.GetMethod("Load").Invoke(null, new object[] { path });
        }
    }
}
