using System;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Text;

namespace ClearPlan.Core.Tests
{
    public static class ConfigurationWorkspaceTests
    {
        public static void RunAll()
        {
            var rateCatalog = new ClearPlan.Core.PlanAnalysis.DoseRateEstimationProfileCatalog {
                SchemaVersion = 1, Profiles = new System.Collections.Generic.List<ClearPlan.Core.PlanAnalysis.DoseRateEstimationProfile>() };
            for (int i = 0; i < 10; i++) rateCatalog.Profiles.Add(new ClearPlan.Core.PlanAnalysis.DoseRateEstimationProfile {
                Id = "Synthetic-" + i, MachineIds = new System.Collections.Generic.List<string> { "Synthetic-" + i },
                MaxGantrySpeedDegreesPerSecond = 6, Assumption = new string('ä', 4000) });
            string rateJson = Newtonsoft.Json.JsonConvert.SerializeObject(rateCatalog);
            TestAssert.NotNull(ClearPlan.Core.PlanAnalysis.DoseRateEstimationProfileCatalog.Parse(rateJson));
            TestAssert.True(rateJson.Length < 65536 && Encoding.UTF8.GetByteCount(rateJson) > 65536);
            TestAssert.Throws<FormatException>(() => ClearPlan.ConfigurationContentValidator.Validate("dose-rate-profiles", Encoding.UTF8.GetBytes(rateJson), null));
            rateCatalog.Profiles.RemoveRange(1, 9);
            ClearPlan.ConfigurationContentValidator.Validate("dose-rate-profiles", Encoding.UTF8.GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(rateCatalog)), null);
            Type workspaceType = Assembly.GetExecutingAssembly().GetType("ClearPlan.ConfigurationWorkspaceViewModel");
            Check(workspaceType != null, "Configuration workspace must expose a testable native editor model.");
            Type entryType = Assembly.GetExecutingAssembly().GetType("ClearPlan.ConfigurationEntryViewModel");
            string root = Path.Combine(Path.GetTempPath(), "clearplan-config-ui-" + Guid.NewGuid().ToString("N"));
            string source = Path.Combine(root, "external.json");
            Directory.CreateDirectory(root);
            File.WriteAllText(source, "{\"value\":1}", new UTF8Encoding(false));
            try
            {
                dynamic entry = Activator.CreateInstance(entryType);
                entry.Key = "defaults"; entry.Title = "Defaults"; entry.Extension = ".json";
                entry.SourcePath = new Func<string>(() => source);
                entry.Validate = new Action<byte[]>(bytes =>
                {
                    if (Encoding.UTF8.GetString(bytes).Contains("invalid")) throw new FormatException("Rejected test content.");
                });
                Array entries = Array.CreateInstance(entryType, 1); entries.SetValue(entry, 0);
                string candidate = null;
                dynamic vm = Activator.CreateInstance(workspaceType, Path.Combine(root, "managed"), entries,
                    new Action<string, string>((key, path) => candidate = path));
                Check(!Directory.Exists(Path.Combine(root, "managed")), "Opening settings must not create a repository or import external files.");
                Check((string)vm.ConfiguredPath == source, "Configured external source must stay visible.");
                Check(!(bool)vm.IsManaged, "External source must not be silently managed.");
                Check(!vm.ImportFile(source), "Import requires an explicit change reason.");
                vm.ChangeReason = "Original reviewed";
                Check(vm.ImportFile(source), "Validated explicit import must succeed.");
                Check(File.ReadAllText(source) == "{\"value\":1}", "Import must leave the external original unchanged.");
                Check(candidate != source && File.Exists(candidate), "Activation candidate must point to a managed copy.");
                string originallyPublished = candidate;
                Check((int)vm.CurrentRevisionNumber == 1, "Import must record the original revision.");
                vm.EditorText = "invalid"; vm.ChangeReason = "Reject invalid JSON";
                Check(!vm.SaveJson(), "Invalid edits must be rejected before committing.");
                Check((int)vm.CurrentRevisionNumber == 1, "Validation failure must not create a revision.");
                Check(File.ReadAllText(candidate) == "{\"value\":1}", "Validation failure must not overwrite managed content.");
                vm.CancelEdit();
                Check((string)vm.EditorText == "{\"value\":1}", "Cancel must restore persisted JSON.");
                vm.EditorText = "{\"value\":2}"; vm.ChangeReason = "Change reviewed";
                Check(vm.SaveJson(), "Valid JSON edit must commit.");
                Check((int)vm.CurrentRevisionNumber == 2, "Edit must append a revision.");
                Check(candidate != originallyPublished && File.ReadAllText(originallyPublished) == "{\"value\":1}",
                    "Saving a new revision must not mutate the previously activated immutable version.");
                bool versionSelectionNotified = false;
                ((INotifyPropertyChanged)vm).PropertyChanged += (sender, args) =>
                {
                    if (args.PropertyName == "SelectedVersion") versionSelectionNotified = true;
                };
                vm.SelectedVersion = vm.Versions[1]; vm.ChangeReason = "Return to baseline";
                Check(versionSelectionNotified, "Selecting another version must notify the bound SHA-256 details.");
                Check(vm.RestoreSelected(), "Restoring an older version must succeed.");
                Check((int)vm.CurrentRevisionNumber == 3, "Restore must append, never rewind or delete history.");
                Check(File.ReadAllText(candidate) == "{\"value\":1}", "Restore must publish the chosen content.");
                Check(File.ReadAllText(source) == "{\"value\":1}", "Edit and restore must never write external originals.");

                dynamic parallel = Activator.CreateInstance(workspaceType, Path.Combine(root, "managed"), entries,
                    new Action<string, string>((key, path) => { }));
                parallel.EditorText = "{\"value\":4}"; parallel.ChangeReason = "Other reviewer";
                Check(parallel.SaveJson(), "A second reviewer may save a fresh edit.");
                vm.EditorText = "{\"value\":5}"; vm.ChangeReason = "Stale editor";
                Check(!vm.SaveJson(), "A stale JSON editor must not silently overwrite another reviewer's revision.");
                Check((string)vm.EditorText == "{\"value\":5}", "A conflict must preserve the unsaved JSON for review.");

                string excelSource = Path.Combine(root, "external.xlsx");
                File.WriteAllText(excelSource, "version-one", new UTF8Encoding(false));
                dynamic excelEntry = Activator.CreateInstance(entryType);
                excelEntry.Key = "constraints"; excelEntry.Title = "Constraints"; excelEntry.Extension = ".xlsx";
                excelEntry.SourcePath = new Func<string>(() => excelSource);
                excelEntry.Validate = new Action<byte[]>(bytes => { });
                Array excelEntries = Array.CreateInstance(entryType, 1); excelEntries.SetValue(excelEntry, 0);
                dynamic excel = Activator.CreateInstance(workspaceType, Path.Combine(root, "managed"), excelEntries,
                    new Action<string, string>((key, path) => { }));
                excel.ChangeReason = "Reviewed workbook";
                Check(excel.ImportFile(excelSource), "Excel import must create a managed revision.");
                string publishedExcel = excel.ManagedPath;
                string excelDraft = excel.PrepareExcelDraft();
                Check(excelDraft != excelSource && excelDraft != publishedExcel,
                    "Excel must open a separate draft, never the external or published file.");
                File.WriteAllText(excelDraft, "version-two", new UTF8Encoding(false));
                Check(File.ReadAllText(publishedExcel) == "version-one", "Editing Excel must not activate draft content.");
                excel.ChangeReason = "Accept workbook change";
                Check(excel.AcceptExcelDraft(), "Explicit Excel acceptance must append a revision.");
                Check(File.ReadAllText(excelSource) == "version-one", "Excel acceptance must preserve the external original.");
                excel.PrepareExcelDraft();
                dynamic otherExcel = Activator.CreateInstance(workspaceType, Path.Combine(root, "managed"), excelEntries,
                    new Action<string, string>((key, path) => { }));
                string otherDraft = otherExcel.PrepareExcelDraft();
                File.WriteAllText(otherDraft, "version-three", new UTF8Encoding(false));
                otherExcel.ChangeReason = "Other workbook reviewer";
                Check(otherExcel.AcceptExcelDraft(), "Second Excel reviewer must be able to save a current draft.");
                excel.Refresh();
                excel.ChangeReason = "Stale Excel draft after status refresh";
                Check(!excel.AcceptExcelDraft(), "Refreshing status must not rebase an old Excel draft over a newer committed version.");
            }
            finally { Directory.Delete(root, true); }
        }

        private static void Check(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }
    }
}
