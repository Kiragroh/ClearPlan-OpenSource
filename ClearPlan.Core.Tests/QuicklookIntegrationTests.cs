using System;
using System.IO;
using System.Windows.Input;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;
using ClearPlan.Reporting;
using ClearPlan.Reporting.MigraDoc;

namespace ClearPlan.Core.Tests
{
    internal static class QuicklookIntegrationTests
    {
        public static void ActionIsSeparateAndUnsubscribed()
        {
            var property = typeof(ReviewWorkspaceViewModel).GetProperty("HtmlReportCommand");
            var binding = typeof(ReviewWorkspaceActionBindings).GetProperty("HtmlReport");
            TestAssert.True(property != null && binding != null, "HTML Quicklook needs an independent action, not the slow PDF export.");
            int calls = 0;
            var bindings = new ReviewWorkspaceActionBindings();
            EventHandler<ReviewWorkspaceActionEventArgs> callback = (sender, args) => {
                TestAssert.Equal("html-report", args.ActionCode);
                calls++;
            };
            binding.SetValue(bindings, callback);
            var controller = new ReviewWorkspaceHostController(() => SyntheticScenarioFactory.Create("baseline-pass"), bindings);
            Exception failure;
            TestAssert.True(controller.TryRefresh(out failure));
            var old = controller.CurrentViewModel;
            ((ICommand)property.GetValue(old)).Execute(null);
            TestAssert.Equal(1, calls);
            TestAssert.True(controller.TryRefresh(out failure));
            ((ICommand)property.GetValue(old)).Execute(null);
            TestAssert.Equal(1, calls);
            var current = controller.CurrentViewModel;
            ((ICommand)property.GetValue(current)).Execute(null);
            TestAssert.Equal(2, calls);
            controller.Dispose();
            ((ICommand)property.GetValue(current)).Execute(null);
            TestAssert.Equal(2, calls);
        }

        public static void NativeQuicklookUsesCapturedDataAndPrivacy()
        {
            string path = Path.Combine("ClearPlan.Script", "MainView.xaml.cs");
            string source = File.ReadAllText(path);
            int start = source.IndexOf("internal void HandleSharedHtmlReport(", StringComparison.Ordinal);
            TestAssert.True(start >= 0, "Native HTML Quicklook action is missing.");
            int end = source.IndexOf("internal void HandleSharedSyntheticReport(", start, StringComparison.Ordinal);
            string method = source.Substring(start, end - start);
            TestAssert.True(method.Contains("CanExportCurrentSnapshot") && method.Contains("ReviewSnapshotReportMapper"));
            TestAssert.True(method.Contains("Anonymize_CheckBox") && method.Contains("snapshot.Synthetic"));
            TestAssert.True(method.Contains("HtmlReportPreviewWindow") && method.Contains("HtmlReviewReportRenderer"));
            foreach (string forbidden in new[] { "PrepareReportBevsAsync", "EsapiPlanImageBuilder", "TryRefresh", "PrintButtonClicked", "Task.Run" })
                TestAssert.False(method.Contains(forbidden), "Quicklook must not trigger acquisition/recalculation: " + forbidden);
            string xaml = File.ReadAllText(Path.Combine("ClearPlan.Presentation", "Views", "ReviewWorkspaceView.xaml"));
            TestAssert.True(xaml.Contains("HtmlReportButton") && xaml.Contains("HtmlReportCommand"));
            string simulator = File.ReadAllText(Path.Combine("ClearPlan.Simulator", "MainWindow.xaml.cs"));
            TestAssert.True(simulator.Contains("HtmlReportRequested += OnHtmlReportRequested")
                && simulator.Contains("HtmlReportRequested -= OnHtmlReportRequested"));
            TestAssert.True(source.Contains("ShowSavedReportStatus(dialog.FileName, document)") && source.Contains("ExportCompanion(pdfPath, document)"),
                "Full native report export should also save HTML without silently overwriting an earlier HTML file.");
            string renderer = File.ReadAllText(Path.Combine("ClearPlan.Reporting.MigraDoc", "HtmlReviewReportRenderer.cs"));
            TestAssert.True(renderer.Contains("FileMode.CreateNew"));
            TestAssert.True(File.ReadAllText(Path.Combine("ClearPlan.Simulator", "SyntheticReportService.cs")).Contains("ExportCompanion(fullPath, document)"));
            string host = File.ReadAllText(Path.Combine("ClearPlan.Script", "Review", "ClinicalReviewWorkspaceHost.cs"));
            int dispatchStart = host.IndexOf("private void OnHtmlReportRequested(", StringComparison.Ordinal);
            int dispatchEnd = host.IndexOf("private void OnWorkspaceDataContextChanged(", dispatchStart, StringComparison.Ordinal);
            TestAssert.True(dispatchStart >= 0 && dispatchEnd > dispatchStart);
            string dispatch = host.Substring(dispatchStart, dispatchEnd - dispatchStart);
            TestAssert.True(dispatch.Contains("ReferenceEquals(sender, CurrentViewModel)") &&
                dispatch.Contains("HtmlReportTask = PrepareHtmlQuicklookAsync()"),
                "The production host must reject stale actions and expose its asynchronous Quicklook task.");
            int preparationStart = host.IndexOf("private async Task PrepareHtmlQuicklookAsync()", StringComparison.Ordinal);
            int preparationEnd = host.IndexOf("internal string HtmlReportDiagnostic", preparationStart, StringComparison.Ordinal);
            TestAssert.True(preparationStart >= 0 && preparationEnd > preparationStart);
            string preparation = host.Substring(preparationStart, preparationEnd - preparationStart);
            int currentWorkspaceGuard = preparation.IndexOf("ReferenceEquals(workspace, CurrentViewModel)", StringComparison.Ordinal);
            int preview = preparation.IndexOf("owner.HandleSharedHtmlReport(snapshot, workspace)", StringComparison.Ordinal);
            TestAssert.True(preparation.Contains("var snapshot = controller.CurrentSnapshot") &&
                preparation.Contains("var workspace = CurrentViewModel") &&
                currentWorkspaceGuard >= 0 && preview > currentWorkspaceGuard,
                "After optional image preparation, the host must recheck the captured workspace before preview.");
            TestAssert.False(preparation.Contains("ConfigureAwait(false)") || preparation.Contains("Task.Run"),
                "Production Quicklook must resume on the owner UI context, not move the preview to a worker.");

            // Exercise the public, headless renderer directly; no local capture harness is required.
            var document = new ReviewSnapshotReportMapper().Map(SyntheticScenarioFactory.Create("baseline-pass"));
            string directory = Path.Combine(Path.GetTempPath(), "ClearPlan-quicklook-contract-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string pdf = Path.Combine(directory, "synthetic.pdf"), first = null, second = null;
            try
            {
                byte[] pdfSentinel = { 37, 80, 68, 70, 45, 83, 89, 78, 84, 72, 69, 84, 73, 67 };
                File.WriteAllBytes(pdf, pdfSentinel);
                var htmlRenderer = new HtmlReviewReportRenderer();
                first = htmlRenderer.ExportCompanion(pdf, document);
                string originalHtml = File.ReadAllText(first);
                second = htmlRenderer.ExportCompanion(pdf, document);
                TestAssert.Equal(Path.ChangeExtension(pdf, ".html"), first);
                TestAssert.False(string.Equals(first, second, StringComparison.OrdinalIgnoreCase),
                    "A repeated companion export must use a new filename.");
                TestAssert.Equal(originalHtml, File.ReadAllText(first), "Existing HTML must not be overwritten.");
                TestAssert.Equal(originalHtml, File.ReadAllText(second), "Both exports use the same captured document.");
                TestAssert.Equal(Convert.ToBase64String(pdfSentinel), Convert.ToBase64String(File.ReadAllBytes(pdf)),
                    "HTML export must leave the adjacent PDF untouched.");
            }
            finally
            {
                foreach (string file in new[] { pdf, first, second })
                    if (file != null && File.Exists(file)) File.Delete(file);
                Directory.Delete(directory);
            }
        }

        public static void PreviewIsInAppAndBlocksNavigation()
        {
            string path = Path.Combine("ClearPlan.Presentation", "Views", "HtmlReportPreviewWindow.cs");
            TestAssert.True(File.Exists(path), "An in-app HTML preview with explicit save is required.");
            string source = File.ReadAllText(path);
            TestAssert.True(source.Contains("NavigateToString") && source.Contains("Navigating") && source.Contains("e.Cancel = true"));
            TestAssert.True(source.Contains("SaveFileDialog") && source.Contains("UTF8Encoding"));
            TestAssert.False(source.Contains("Process.Start") || source.Contains("ShellExecute") || source.Contains("ObjectForScripting"));
        }
    }
}
