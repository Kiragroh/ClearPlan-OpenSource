using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Simulation;
using ClearPlan.Reporting;
using ClearPlan.Reporting.MigraDoc;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.IO;
using MigraDoc.DocumentObjectModel.Tables;

namespace ClearPlan.Core.Tests
{
    internal static class ReportPresentationTests
    {
        public static void PatientIdentityIsInTheRepeatingFooter()
        {
            var data = new ReviewSnapshotReportMapper().Map(SyntheticScenarioFactory.Create("baseline-pass"));
            data.Synthetic = false;
            data.PatientDisplayLabel = "Example-Family, Demo-Given | ID: DEMO-001";
            var document = new Document(); var section = document.AddSection();
            var method = typeof(ReportPdf).GetMethod("AddReviewHeaderAndFooter", BindingFlags.Instance | BindingFlags.NonPublic);
            method.Invoke(new ReportPdf(), new object[] { section, data, "" });
            string ddl = DdlWriter.WriteToString(document);
            TestAssert.True(ddl.Contains(data.PatientDisplayLabel), "Patient name and ID must appear in the repeating primary footer, not only the report body.");
            TestAssert.True(ddl.Contains("Page") && ddl.Contains("NumPages"), "Keep page numbering alongside patient identification.");
            TestAssert.True(section.Elements.Count == 0, "Identity belongs to the footer, not a one-off body paragraph.");
        }

        public static void PatientFooterKeepsSyntheticAndPseudonymContextsSeparate()
        {
            var data = new ReviewSnapshotReportMapper().Map(SyntheticScenarioFactory.Create("baseline-pass"));
            data.PatientDisplayLabel = "DO-NOT-COPY-NATIVE-IDENTITY";
            var method = typeof(ReportPdf).GetMethod("AddReviewHeaderAndFooter", BindingFlags.Instance | BindingFlags.NonPublic);
            var document = new Document(); var section = document.AddSection();
            method.Invoke(new ReportPdf(), new object[] { section, data, "" });
            string ddl = DdlWriter.WriteToString(document);
            TestAssert.True(ddl.Contains("Synthetic demo patient") && ddl.Contains("ID: DEMO-"), "Synthetic footer needs explicit dummy name and ID.");
            TestAssert.False(ddl.Contains("DO-NOT-COPY-NATIVE-IDENTITY"));
            data.Synthetic = false;
            data.PatientDisplayLabel = "Pseudonym: REVIEW-007";
            document = new Document(); section = document.AddSection();
            method.Invoke(new ReportPdf(), new object[] { section, data, "" });
            TestAssert.True(DdlWriter.WriteToString(document).Contains("Pseudonym: REVIEW-007"), "Do not bypass the existing GUI anonymization choice.");
        }

        public static void LegacyReportAlsoRepeatsPatientIdentity()
        {
            var document = new Document(); var section = document.AddSection();
            var type = typeof(ReportPdf).Assembly.GetType("ClearPlan.Reporting.MigraDoc.Internal.HeaderAndFooter");
            type.GetMethod("AddFooter", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(
                Activator.CreateInstance(type, true), new object[] { section,
                    new ReportPatient { FirstName = "Demo-Given", LastName = "Example-Family", Id = "DEMO-002" } });
            string ddl = DdlWriter.WriteToString(document);
            TestAssert.True(ddl.Contains("Example-Family") && ddl.Contains("Demo-Given") && ddl.Contains("DEMO-002"),
                "The legacy report path also needs patient identity on every page.");
        }

        public static void HeadlessIdentityIsAttachedOnlyForThePrivatePdf()
        {
            // Keep the registered regression name; verify the public mapper/renderer boundary
            // rather than requiring an institution-local headless capture utility.
            string source = File.ReadAllText(Path.Combine("ClearPlan.Script", "MainView.xaml.cs"));
            int start = source.IndexOf("internal async void HandleSharedPlanReport(", StringComparison.Ordinal);
            int end = source.IndexOf("private static void ApplyReportOptions(", start, StringComparison.Ordinal);
            TestAssert.True(start >= 0 && end > start);
            string handler = source.Substring(start, end - start);
            int map = handler.IndexOf("new ReviewSnapshotReportMapper().Map(snapshot)", StringComparison.Ordinal);
            int identity = handler.IndexOf("document.PatientDisplayLabel =", StringComparison.Ordinal);
            int export = handler.IndexOf("new ReportPdf().Export(dialog.FileName, document)", StringComparison.Ordinal);
            TestAssert.True(map >= 0 && identity > map && export > identity,
                "The production PDF route must attach identity only to the detached report copy before export.");
            string identityProjection = handler.Substring(identity, export - identity);
            foreach (string required in new[] { "Anonymize_CheckBox.IsChecked", "Pseudonym: ",
                "_vm.Patient.LastName", "_vm.Patient.FirstName", "_vm.Patient.Id" })
                TestAssert.True(identityProjection.Contains(required), "The PDF identity projection must preserve " + required);
            TestAssert.False(handler.Contains("snapshot.PatientDisplayLabel ="),
                "The private report must not attach patient identity to the reusable snapshot.");

            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            snapshot.Synthetic = false;
            snapshot.PatientDisplayLabel = "Synthetic identity-free review context";
            string before = ClearPlan.Core.Review.ReviewSnapshotJson.Serialize(snapshot);
            var report = new ReviewSnapshotReportMapper().Map(snapshot);
            report.PatientDisplayLabel = "Example-Family, Demo-Given | ID: SYNTHETIC-PRIVATE-001";
            var document = new Document(); var section = document.AddSection();
            typeof(ReportPdf).GetMethod("AddReviewHeaderAndFooter", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { section, report, "" });
            TestAssert.True(DdlWriter.WriteToString(document).Contains(report.PatientDisplayLabel));
            TestAssert.True(new HtmlReviewReportRenderer().Render(report).Contains("SYNTHETIC-PRIVATE-001"),
                "The explicitly private HTML companion uses the same selected identity as the PDF.");
            TestAssert.Equal(before, ClearPlan.Core.Review.ReviewSnapshotJson.Serialize(snapshot),
                "Attaching identity to either report must leave diagnostic snapshot serialization unchanged.");
            TestAssert.False(before.Contains("SYNTHETIC-PRIVATE-001"));
        }

        public static void GuiReportButtonsUseTheCurrentWorkspace()
        {
            string source = System.IO.File.ReadAllText(System.IO.Path.Combine("ClearPlan.Script", "MainView.xaml.cs"));
            int start = source.IndexOf("private void PrintButtonClicked(", StringComparison.Ordinal);
            int end = source.IndexOf("private void PrintLegacyReportClicked(", StringComparison.Ordinal);
            TestAssert.True(start >= 0 && end > start, "Keep the old exporter explicitly separate from the current PDF button.");
            string handler = source.Substring(start, end - start);
            TestAssert.True(handler.Contains("_clinicalReviewHost.RequestCurrentReport()"), "Every GUI PDF button must dispatch through the current snapshot and mode guard.");
            TestAssert.False(handler.Contains("CreateReportData("), "Never silently fall back to the legacy patient report.");
            string host = System.IO.File.ReadAllText(System.IO.Path.Combine("ClearPlan.Script", "Review", "ClinicalReviewWorkspaceHost.cs"));
            TestAssert.True(host.Contains("public void RequestCurrentReport()") && host.Contains("OnReportRequested(CurrentViewModel, null)"));
            string xaml = System.IO.File.ReadAllText(System.IO.Path.Combine("ClearPlan.Script", "MainView.xaml"));
            TestAssert.False(xaml.Contains("Click=\"PrintLegacyReportClicked\""));
        }

        public static void ManuscriptSandboxHidesTheNativeContext()
        {
            string xaml = System.IO.File.ReadAllText(System.IO.Path.Combine("ClearPlan.Script", "MainView.xaml"));
            TestAssert.True(xaml.Contains("Sandbox (Manuskript)") && xaml.Contains("Sandbox aktiv"), "Make the synthetic screenshot mode discoverable in the ESAPI GUI.");
            string source = System.IO.File.ReadAllText(System.IO.Path.Combine("ClearPlan.Script", "MainView.xaml.cs"));
            foreach (string control in new[] { "ClinicalContextText", "ClinicalConstraintPanel", "ClinicalConstraintStatus", "ClinicalExtrasMenuItem" })
                TestAssert.True(source.Contains(control + ".Visibility = enabled\n                ? Visibility.Collapsed".Replace("\n", "\r\n")) ||
                    source.Replace("\r\n", "\n").Contains(control + ".Visibility = enabled\n                ? Visibility.Collapsed"), control + " must be hidden in sandbox mode.");
            TestAssert.True(source.Contains("UpdateWindowTitleForSyntheticDemo(enabled)"));
            TestAssert.True(source.Contains("ClearPlan · Synthetische Demonstration"));
            TestAssert.True(source.Contains("NewID_TextBox.Visibility = pseudonymVisibility") && source.Contains("NewID_TextBlock.Visibility = pseudonymVisibility"), "Hide a previously entered clinical pseudonym in the sandbox header.");
            TestAssert.True(source.Contains("Sandbox bleibt aktiv; interne Fehlerdetails sind ausgeblendet."), "Do not show native exception details after a failed switch back to clinical data.");
            TestAssert.True(source.Contains("Lokale Konfiguration ist im Manuskript-Sandbox-Modus ausgeblendet."), "Keep local configuration and paths outside manuscript screenshots.");
            int syntheticReport = source.IndexOf("internal void HandleSharedSyntheticReport(", StringComparison.Ordinal);
            int clinicalReport = source.IndexOf("internal async void HandleSharedPlanReport", syntheticReport, StringComparison.Ordinal);
            TestAssert.True(source.Substring(syntheticReport, clinicalReport - syntheticReport).Contains("series.Selected = selection.IsSelected"), "Synthetic PDF curves must follow the current GUI selections.");
            string host = System.IO.File.ReadAllText(System.IO.Path.Combine("ClearPlan.Script", "Review", "ClinicalReviewWorkspaceHost.cs"));
            TestAssert.True(host.Contains("SyntheticScenarioFactory.Create(\"mixed-review\")") && host.Contains("if (enabled) planContext.Clear()"));
            int report = host.IndexOf("private void OnReportRequested(", StringComparison.Ordinal);
            int synthetic = host.IndexOf("if (IsSyntheticDemo)", report, StringComparison.Ordinal);
            int native = host.IndexOf("owner.HandleSharedPlanReport", report, StringComparison.Ordinal);
            TestAssert.True(synthetic > report && native > synthetic && host.Substring(synthetic, native - synthetic).Contains("owner.HandleSharedSyntheticReport("));
        }

        public static void SixSourcesFitTheSummaryPage()
        {
            var data = new ReviewSnapshotReportMapper().Map(SyntheticScenarioFactory.Create("baseline-pass"));
            data.Synthetic = false;
            data.Title = "ClearPlan - Native plan review";
            data.Subtitle = "Read-only development PDF preview";
            data.Watermark = "Development preview; not a clinical release or approval.";
            data.ProvenanceText = "Eclipse ESAPI clinical read-only projection.";
            data.Plans = data.Plans.Take(1).ToList();
            data.PlanAnalysis = SyntheticPlanAnalysisFactory.Create(true, true);
            data.PlanAnalysis.Beams = data.PlanAnalysis.Beams.Take(2).ToList();
            foreach (var beam in data.PlanAnalysis.Beams) { beam.BeamId = "ARC" + beam.BeamNumber; beam.MlcModel = "DUAL"; }
            data.PlanAnalysis.Pam = null;
            data.PlanAnalysis.PamReason = "No explicit, unambiguous target structure selected. PAM is structure-specific.";
            data.Sources.Clear();
            foreach (int length in new[] { 48, 230, 180, 198, 690, 300 })
            {
                string sentence = "Synthetic source status and configuration provenance, explicitly separate from clinical approval. ";
                string message = string.Concat(Enumerable.Repeat(sentence, 12)).Substring(0, length);
                data.Sources.Add(new ReviewReportSourceRow { PathDisplayLabel = "Configuration source", Status = "available", Message = message });
            }
            var document = (Document)typeof(ReportPdf).GetMethod("CreateReviewReport", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { data });
            var elements = document.Sections[0].Elements;
            var summary = elements.Cast<DocumentObject>().TakeWhile(item => !(item is PageBreak)).Select(item => item.Clone()).ToList();
            elements.Clear();
            foreach (var item in summary) elements.Add((DocumentObject)item);
            var renderer = new MigraDoc.Rendering.PdfDocumentRenderer(true) { Document = document };
            renderer.RenderDocument();
            if (renderer.PdfDocument.PageCount != 1)
            {
                string diagnostic = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ClearPlan-summary-regression-" + Guid.NewGuid().ToString("N") + ".pdf");
                renderer.PdfDocument.Save(diagnostic);
                Console.WriteLine("Synthetic summary pagination diagnostic: " + diagnostic);
            }
            TestAssert.Equal(1, renderer.PdfDocument.PageCount,
                "The field-naming source must not orphan a final source row onto a nearly empty summary continuation.");
        }

        public static void CtSectionsHaveDedicatedPages()
        {
            var document = new Document(); var section = document.AddSection();
            typeof(ReportPdf).GetMethod("AddPlanImages", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { section, SyntheticPlanImageFactory.Create("synthetic-plan"), false });
            TestAssert.Equal(2, section.Elements.Cast<DocumentObject>().OfType<PageBreak>().Count());
            TestAssert.True(DdlWriter.WriteToString(document).Contains("Isocenter"));
        }

        public static void CtImageIsLargerAndItsIsocenterIsCentered()
        {
            var images = SyntheticPlanImageFactory.Create("synthetic-plan");
            var image = images[0];
            image.IsocenterPixelX = image.WidthPixels * 0.3;
            image.IsocenterPixelY = image.HeightPixels * 0.7;
            string encoded = (string)typeof(ReportPdf).Assembly.GetType("ClearPlan.Reporting.MigraDoc.Internal.ReviewImageRenderer")
                .GetMethod("Ct", BindingFlags.Static | BindingFlags.Public).Invoke(null, new object[] { image });
            using (var stream = new System.IO.MemoryStream(Convert.FromBase64String(encoded.Substring(7))))
            using (var bitmap = new System.Drawing.Bitmap(stream))
            {
                var color = bitmap.GetPixel(720, 720);
                TestAssert.True(color.G > color.R + 25 && color.G > color.B,
                    "The projected isocenter crosshair must lie at the exact image center, not at the old full-CT position.");
            }
            var document = new Document(); var section = document.AddSection();
            typeof(ReportPdf).GetMethod("AddPlanImages", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { section, images, false });
            var table = section.Elements.Cast<DocumentObject>().OfType<Table>().First();
            TestAssert.True(table.Columns[0].Width.Centimeter >= 17,
                "CT image panel still smaller than 17 cm despite available landscape width.");
            TestAssert.True(DdlWriter.WriteToString(document).Contains("Isocenter-centered"));
            var paginated = CreateDocument(); paginated.Sections[0].Elements.Clear();
            typeof(ReportPdf).GetMethod("AddPlanImages", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { paginated.Sections[0], images, false });
            var renderer = new MigraDoc.Rendering.PdfDocumentRenderer(true) { Document = paginated };
            renderer.RenderDocument();
            TestAssert.Equal(3, renderer.PdfDocument.PageCount);
        }

        public static void SyntheticCtBannerDoesNotCoverOrientation()
        {
            var method = typeof(ReportPdf).Assembly.GetType("ClearPlan.Reporting.MigraDoc.Internal.ReviewImageRenderer")
                .GetMethod("Ct", BindingFlags.Static | BindingFlags.Public);
            foreach (var image in SyntheticPlanImageFactory.Create("synthetic-plan"))
            {
                string encoded = (string)method.Invoke(null, new object[] { image });
                using (var stream = new System.IO.MemoryStream(Convert.FromBase64String(encoded.Substring(7))))
                using (var bitmap = new System.Drawing.Bitmap(stream))
                {
                    int visibleOrientationPixels = 0;
                    for (int y = 1390; y < 1435; y++)
                    for (int x = 690; x < 750; x++)
                    {
                        var color = bitmap.GetPixel(x, y);
                        if (color.G > color.R + 40 && color.G > color.B) visibleOrientationPixels++;
                    }
                    TestAssert.True(visibleOrientationPixels > 20, "The simulation banner must not cover the bottom CT orientation: " + image.Kind);
                }
            }
        }

        public static void CtMaximumSupportedLegendRetainsOnePagePerPlane()
        {
            var images = SyntheticPlanImageFactory.Create("synthetic-plan");
            foreach (var image in images)
            {
                image.Caption = "Planning CT; nearest planes to treatment isocenter. W400 / L40 HU; nearest-neighbor overview. Crosshair is the projected treatment isocenter. Plane offsets x/y/z: 0.16/-0.25/0.33 mm.";
                image.OverlaySummary = "Structure display; up to 24 nonempty segmented structures, targets first. Empty paths mean no intersection with this plane; unavailable sources are listed explicitly.";
                for (int i = 0; i < 31; i++) image.Overlays.Add(new ClearPlan.Core.Review.ReviewImageOverlay {
                    Kind = i < 24 ? "structure" : "isodose", Label = i < 24 ? "Structure_long_example_name_" + i : "107% Rx / 53.5 Gy",
                    SourceStatus = "available", ColorHex = "#33BBCC",
                    Paths = new List<ClearPlan.Core.Review.ReviewImagePath> { new ClearPlan.Core.Review.ReviewImagePath {
                        Points = new List<ClearPlan.Core.Review.ReviewImagePoint> {
                            new ClearPlan.Core.Review.ReviewImagePoint { X = 1, Y = 1 }, new ClearPlan.Core.Review.ReviewImagePoint { X = 2, Y = 2 } } } }
                });
            }
            var document = CreateDocument(); document.Sections[0].Elements.Clear();
            typeof(ReportPdf).GetMethod("AddPlanImages", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { document.Sections[0], images, false });
            string ddl = DdlWriter.WriteToString(document);
            TestAssert.False(ddl.Contains("nearest-neighbor overview") || ddl.Contains("Structure display; up to 24"), "CT captions must omit capture algorithm prose.");
            TestAssert.True(ddl.Contains("W400 / L40 HU"), "Keep the actual supplied window setting compactly.");
            var renderer = new MigraDoc.Rendering.PdfDocumentRenderer(true) { Document = document };
            renderer.RenderDocument();
            TestAssert.Equal(3, renderer.PdfDocument.PageCount);
        }

        public static void MissingPqmIsVisibleInsteadOfAnEmptyHeader()
        {
            var document=new Document(); var section=document.AddSection();
            typeof(ReportPdf).GetMethod("AddPqmRows",BindingFlags.Instance|BindingFlags.NonPublic)
                .Invoke(new ReportPdf(),new object[] {section,new List<ReviewReportPqmRow>()});
            TestAssert.True(DdlWriter.WriteToString(document).Contains("PQM not evaluated"));
            TestAssert.Equal(0,section.Elements.Cast<DocumentObject>().OfType<Table>().Count());
        }
        public static void DvhTableLeadsWithRequiredTargetsAndRetainsManualChoices()
        {
            var document = new Document(); var section = document.AddSection();
            var rows = new List<ReviewReportDvhSeries> {
                new ReviewReportDvhSeries { DisplayName="Heart", StructureId="Heart", Selected=true },
                new ReviewReportDvhSeries { DisplayName="Kidney_unselected", StructureId="Kidney_unselected", Role="organ-at-risk", Selected=false },
                new ReviewReportDvhSeries { DisplayName="Lung_selected_unavailable", StructureId="Lung_selected_unavailable", Selected=true, Statistics=null },
                new ReviewReportDvhSeries { DisplayName="GTV_small", StructureId="GTV_small", TargetKind="GTV" },
                new ReviewReportDvhSeries { DisplayName="CTV_largest", StructureId="CTV_largest", TargetKind="CTV", RequiredForTargetReview=true, Selected=false },
                new ReviewReportDvhSeries { DisplayName="PTV", StructureId="PTV", TargetKind="PTV", RequiredForTargetReview=true, Selected=true },
                new ReviewReportDvhSeries { DisplayName="ITV_manual", StructureId="ITV_manual", TargetKind="ITV", Selected=true }
            };
            typeof(ReportPdf).GetMethod("AddDvhRows", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { section, rows });
            string ddl = DdlWriter.WriteToString(document);
            var table = section.Elements.Cast<DocumentObject>().OfType<Table>().Single();
            TestAssert.Equal(5, table.Columns.Count, "DVH statistics must omit the Curve and Availability columns.");
            TestAssert.False(ddl.Contains("Availability") || ddl.Contains("Shown in plot") || ddl.Contains("Not selected"));
            TestAssert.False(ddl.Contains("GTV_small"), "Default report table must not add every smaller inner target.");
            TestAssert.False(ddl.Contains("Kidney_unselected"), "Unselected OARs must not appear in the PDF table even without GUI hidden IDs.");
            TestAssert.True(ddl.Contains("Lung_selected_unavailable"), "Selection must retain unavailable statistics, not silently drop the structure.");
            TestAssert.Equal(6, table.Rows.Count, "Header plus five selected or required structures must remain.");
            TestAssert.True(ddl.IndexOf("PTV", StringComparison.Ordinal) < ddl.IndexOf("CTV_largest", StringComparison.Ordinal));
            TestAssert.True(ddl.IndexOf("CTV_largest", StringComparison.Ordinal) < ddl.IndexOf("Heart", StringComparison.Ordinal));
            TestAssert.True(ddl.Contains("ITV_manual"), "Explicit manual target selections remain available.");
            TestAssert.Equal(7, rows.Count, "Rendering must never remove DVH inventory entries.");
        }
        public static void PlanCheckFindingsUseTheAvailableSpace()
        {
            var document = new Document(); var section = document.AddSection();
            var rows = new List<ReviewReportCheckRow> { new ReviewReportCheckRow {
                CheckCode = "test", Category = "Geometry", Status = "info", Message = "Named finding",
                ObservedValue = "4", ExpectedValue = "5", Unit = "mm" } };
            typeof(ReportPdf).GetMethod("AddPlanCheckRows", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { section, rows });
            TestAssert.Equal(4, section.Elements.Cast<DocumentObject>().OfType<Table>().Single().Columns.Count);
            string ddl = DdlWriter.WriteToString(document);
            TestAssert.True(ddl.Contains("Named finding") && !ddl.Contains("Observed:") && !ddl.Contains("Expected:"));
            var table = section.Elements.Cast<DocumentObject>().OfType<Table>().Single();
            TestAssert.True(DdlWriter.WriteToString(table.Rows[0].Cells[0]).Contains("Message"));
            TestAssert.True(DdlWriter.WriteToString(table.Rows[0].Cells[1]).Contains("Status"));
        }
        public static void CtLegendDescribesContinuousSegmentOverlays()
        {
            var snapshot = SyntheticScenarioFactory.Create("mixed-review");
            snapshot.PlanImages = SyntheticPlanImageFactory.Create(snapshot.ActivePlanKey);
            snapshot.PlanImages[0].Overlays.Add(new ClearPlan.Core.Review.ReviewImageOverlay {
                Kind = "isodose", Label = "50% Rx", SourceStatus = "available", ColorHex = "#33BBCC",
                Paths = new List<ClearPlan.Core.Review.ReviewImagePath> { new ClearPlan.Core.Review.ReviewImagePath {
                    Points = new List<ClearPlan.Core.Review.ReviewImagePoint> { new ClearPlan.Core.Review.ReviewImagePoint { X = 1, Y = 1 }, new ClearPlan.Core.Review.ReviewImagePoint { X = 2, Y = 2 } } } }
            });
            var document = new Document(); var section = document.AddSection();
            typeof(ReportPdf).GetMethod("AddPlanImages", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { section, snapshot.PlanImages, true });
            TestAssert.True(DdlWriter.WriteToString(document).Contains("Contours: thin | Isodoses: thick"));
        }
        public static void UnknownPhysicalLayersAreNotPresentedAsOne()
        {
            var document=new Document();var section=document.AddSection();
            var analysis=SyntheticPlanAnalysisFactory.Create(false,true);
            analysis.Beams[0].GeometryStatus="unavailable";analysis.Beams[0].MlcLayerCount=1;
            typeof(ReportPdf).GetMethod("AddPlanAnalysis",BindingFlags.Instance|BindingFlags.NonPublic)
                .Invoke(new ReportPdf(),new object[] {section,analysis});
            TestAssert.True(DdlWriter.WriteToString(document).Contains("Unverified"),
                "An unavailable native profile must not turn the DTO default layer count into a physical assertion.");
        }
        public static void PamSelectionCountsAreCompactAndExplicit()
        {
            var data = new ReviewSnapshotReportMapper().Map(SyntheticScenarioFactory.Create("mixed-review"));
            data.PlanAnalysis = SyntheticPlanAnalysisFactory.Create(true, true);
            var plan = data.PlanAnalysis;
            plan.TargetSelectionMode = "AutomaticLowestPam";
            plan.TargetStructureId = "PTV_SELECTED_SYNTHETIC";
            plan.PamTargetCandidateCount = 3;
            plan.PamValidTargetCandidateCount = 2;
            plan.TargetSelectionProvenance = "LONG_SELECTION_PROSE_MUST_NOT_PRINT";
            foreach (string text in ParameterTexts(data))
            {
                TestAssert.True(text.Contains("PTV_SELECTED_SYNTHETIC") && text.Contains("2/3 eligible PTVs valid"));
                TestAssert.True(text.Contains("Automatic") && text.Contains("lowest valid PAM") && text.Contains("not clinical superiority"));
                TestAssert.False(text.Contains("LONG_SELECTION_PROSE_MUST_NOT_PRINT"));
            }
            plan.TargetSelectionMode = "Explicit";
            foreach (string text in ParameterTexts(data))
            {
                TestAssert.True(text.Contains("PTV_SELECTED_SYNTHETIC") && text.Contains("Explicit selection"));
                TestAssert.False(text.Contains("2/3 eligible PTVs valid") || text.Contains("Automatic"));
            }
            plan.TargetSelectionMode = "AutomaticLowestPam";
            plan.PamValidTargetCandidateCount = 0;
            plan.TargetStructureId = null;
            plan.Pam = null;
            plan.PamReason = "No eligible PTV has complete PAM. LONG_SELECTION_PROSE_MUST_NOT_PRINT";
            foreach (string text in ParameterTexts(data))
            {
                TestAssert.True(text.Contains("0/3 eligible PTVs valid") && text.Contains("No valid PTV candidate"));
                TestAssert.False(text.Contains("LONG_SELECTION_PROSE_MUST_NOT_PRINT"));
            }
        }

        private static IEnumerable<string> ParameterTexts(ReviewReportDocument data)
        {
            var document = new Document(); var section = document.AddSection();
            typeof(ReportPdf).GetMethod("AddPlanAnalysis", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { section, data.PlanAnalysis });
            yield return DdlWriter.WriteToString(document);
            yield return System.Net.WebUtility.HtmlDecode(new HtmlReviewReportRenderer().Render(data));
        }

        public static void MissingTargetQualityDoesNotReserveAnEmptyPage()
        {
            var data = new ReviewSnapshotReportMapper().Map(SyntheticScenarioFactory.Create("mixed-review"));
            data.IncludeBeamEyeViews = false;
            data.PlanAnalysis = SyntheticPlanAnalysisFactory.Create(true, true);
            data.PlanAnalysis.TargetQuality = new List<ReviewTargetQuality> { new ReviewTargetQuality { StructureId = "PTV_SYNTHETIC", TargetVolumeCm3 = 1 } };
            var create = typeof(ReportPdf).GetMethod("CreateReviewReport", BindingFlags.Instance | BindingFlags.NonPublic);
            var populated = (Document)create.Invoke(new ReportPdf(), new object[] { data });
            int withQuality = populated.Sections[0].Elements.Cast<DocumentObject>().OfType<PageBreak>().Count();
            data.PlanAnalysis.TargetQuality.Clear();
            var unavailable = (Document)create.Invoke(new ReportPdf(), new object[] { data });
            TestAssert.Equal(withQuality - 1, unavailable.Sections[0].Elements.Cast<DocumentObject>().OfType<PageBreak>().Count());
            TestAssert.True(DdlWriter.WriteToString(unavailable).Contains("PTV quality unavailable"));
        }
        public static void PqmSourcesAreGroupedAndMatchesStayVisible()
        {
            var document = new Document(); var section = document.AddSection();
            var rows = new List<ReviewReportPqmRow> {
                new ReviewReportPqmRow { TemplateStructure = "Heart", ResolvedStructureId = "Heart", SourceLabel = "RefDB: ExampleTable; Reference test paper 2024", MappingDescription = "Exact name" },
                new ReviewReportPqmRow { TemplateStructure = "Lung", SourceLabel = "RefDB: ExampleTable; Reference test paper 2024", MappingDescription = "Unresolved", Status = "not-evaluated", Severity = "info" },
                new ReviewReportPqmRow { TemplateStructure = "PTV", SourceLabel = "Default", MappingDescription = "Exact target" },
                new ReviewReportPqmRow { TemplateStructure = "Brain", SourceLabel = "Eclipse Clinical Goals: assigned plan objectives", MappingDescription = "Exact native structure" },
                new ReviewReportPqmRow { TemplateStructure = "Kidney", SourceLabel = "Another reference paper", MappingDescription = "Alias" }
            };
            typeof(ReportPdf).GetMethod("AddPqmRows", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { section, rows });
            string ddl = DdlWriter.WriteToString(document);
            TestAssert.False(ddl.Contains("Reference test paper") || ddl.Contains("Another reference paper") || ddl.Contains("S1:"),
                "Clinical report goals must not print literature references or bibliography footnotes.");
            TestAssert.False(ddl.Contains("Source / structure match") || ddl.Contains("Exact name") || ddl.Contains("Exact native structure"));
            TestAssert.Equal(8, section.Elements.Cast<DocumentObject>().OfType<Table>().Single().Columns.Count);
            TestAssert.True(ddl.Contains("Heart") && ddl.Contains("Lung"));
            TestAssert.True(rows[0].SourceLabel.Contains("Reference test paper"), "Report formatting must retain original configuration provenance in the detached data.");
            TestAssert.True(ddl.Contains("not-evaluated") && !ddl.Contains("(info)"),
                "A redundant info suffix must not double every unconfirmed PQM row and orphan report notes.");
        }
        public static void BeamStartPanelsStayWithinOnePagePerBeam()
        {
            // Isolate BEV pagination from the length of PQM, source and check tables.
            var document = CreateDocument();
            document.Sections[0].Elements.Clear();
            typeof(ReportPdf).GetMethod("AddBeamViews", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { document.Sections[0], SyntheticPlanAnalysisFactory.Create(true, true), true });
            var renderer = new MigraDoc.Rendering.PdfDocumentRenderer(true) { Document = document };
            renderer.RenderDocument();
            // MigraDoc ignores the leading page break: exactly one page per beam.
            if (renderer.PdfDocument.PageCount != 2)
            {
                string output = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bev-pagination-diagnostic.pdf");
                renderer.PdfDocument.Save(output);
            }
            TestAssert.Equal(2, renderer.PdfDocument.PageCount);
        }
        public static void ReportOptionsFilterOnlyTheRenderedEvidence()
        {
            var data = new ReviewSnapshotReportMapper().Map(SyntheticScenarioFactory.Create("mixed-review"));
            data.PlanImages = SyntheticPlanImageFactory.Create(data.ActivePlanKey);
            data.PlanAnalysis = SyntheticPlanAnalysisFactory.Create(true, true);
            data.Sources[0].Message = "SOURCE_PROSE_MUST_NOT_PRINT";
            data.PqmRows[0].TemplateStructure = "UNMATCHED_PQM_MUST_NOT_PRINT";
            data.PqmRows[0].ResolvedStructureId = " ";
            data.DvhSeries[0].StructureId = "HIDDEN_STRUCTURE";
            data.DvhSeries[0].DisplayName = "HIDDEN_STRUCTURE";
            data.DvhSeries[0].RequiredForTargetReview = true;
            data.PlanImages[0].Overlays.Add(new ClearPlan.Core.Review.ReviewImageOverlay { Kind = "structure", Label = "HIDDEN_STRUCTURE", SourceStatus = "available",
                Paths = new List<ClearPlan.Core.Review.ReviewImagePath> { new ClearPlan.Core.Review.ReviewImagePath {
                    Points = new List<ClearPlan.Core.Review.ReviewImagePoint> { new ClearPlan.Core.Review.ReviewImagePoint { X = 5, Y = 5 },
                        new ClearPlan.Core.Review.ReviewImagePoint { X = 25, Y = 25 } } } } });
            foreach (var option in new Dictionary<string, object> { { "IncludeBeamEyeViews", false }, { "HideUnmatched", true },
                { "HiddenStructureIds", new List<string> { "hidden_structure" } }, { "DisabledCheckCount", 2 } })
            {
                var property = typeof(ReviewReportDocument).GetProperty(option.Key);
                TestAssert.NotNull(property, "Missing report option " + option.Key);
                property.SetValue(data, option.Value, null);
            }
            var document = (Document)typeof(ReportPdf).GetMethod("CreateReviewReport", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { data });
            string ddl = DdlWriter.WriteToString(document);
            TestAssert.False(ddl.Contains("SOURCE_PROSE_MUST_NOT_PRINT") || ddl.Contains("UNMATCHED_PQM_MUST_NOT_PRINT") || ddl.Contains("HIDDEN_STRUCTURE"));
            TestAssert.False(ddl.Contains("Beam's-eye view"));
            TestAssert.True(ddl.Contains("checks disabled") && ddl.Contains("unmatched goal"));
            TestAssert.True(data.PlanImages[0].Overlays.Any(item => item.Label == "HIDDEN_STRUCTURE"));
            TestAssert.True(data.DvhSeries.Any(item => item.StructureId == "HIDDEN_STRUCTURE"));
            var empty = new Document(); var section = empty.AddSection();
            typeof(ReportPdf).GetMethod("AddBeamViews", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { section, data.PlanAnalysis, false });
            TestAssert.Equal(0, section.Elements.Cast<DocumentObject>().OfType<PageBreak>().Count(), "Missing native DRRs need compact status, not empty image pages.");
        }
        public static void TablesUseAvailablePageWidth()
        {
            Document report = CreateDocument();
            var section = report.Sections[0];
            var tables = section.Elements.Cast<DocumentObject>().OfType<Table>().ToList();
            TestAssert.True(tables.Count > 5);
            foreach (Table table in tables)
                TestAssert.True(Math.Abs(27.3 - table.Columns.Cast<Column>().Sum(column => column.Width.Centimeter)) < 0.001,
                    "Every review table must span the 27.3 cm content width.");
        }

        public static void ActivePlanHeaderDoesNotImplyApproval()
        {
            var snapshot = SyntheticScenarioFactory.Create("mixed-review");
            var document = new Document();
            var section = document.AddSection();
            var method = typeof(ReportPdf).GetMethod("AddPlanRows", BindingFlags.NonPublic | BindingFlags.Instance);
            TestAssert.NotNull(method);
            method.Invoke(new ReportPdf(), new object[] { section, new ReviewSnapshotReportMapper().Map(snapshot).Plans });
            string ddl = DdlWriter.WriteToString(document);
            TestAssert.False(ddl.Contains("Status"), "A plan metadata row status is not an overall plan review or approval result.");
        }

        public static void RepresentativeBevUsesMiddleCompletePositiveMuSample()
        {
            var beam = SyntheticPlanAnalysisFactory.Create(true, true).Beams[0];
            foreach (var cp in beam.ControlPoints) cp.MetricMetersetWeightMu = 1;
            beam.ControlPoints[0].MetricMetersetWeightMu = 0;
            beam.ControlPoints[1].Aperture = null;
            var candidates = beam.ControlPoints.Where(cp => cp.MetricMetersetWeightMu > 0 && cp.Aperture != null).ToList();
            MethodInfo method = RendererMethod("RepresentativeControlPoint");
            var chosen = (ReviewControlPointSample)method.Invoke(null, new object[] { beam });
            TestAssert.Equal(candidates[candidates.Count / 2].Index, chosen.Index);
            foreach (var cp in beam.ControlPoints) cp.MetricMetersetWeightMu = 0;
            TestAssert.True(method.Invoke(null, new object[] { beam }) == null,
                "No positive-MU complete sample must remain unavailable, not fall back to CP0.");
        }

        public static void TraceGapsAreNotConnected()
        {
            var samples = new List<ReviewControlPointSample>
            {
                new ReviewControlPointSample { Index = 0, PlannedDoseRateMuPerMin = 500 },
                new ReviewControlPointSample { Index = 1, PlannedDoseRateMuPerMin = null },
                new ReviewControlPointSample { Index = 2, PlannedDoseRateMuPerMin = 600 },
                new ReviewControlPointSample { Index = 3, PlannedDoseRateMuPerMin = double.NaN },
                new ReviewControlPointSample { Index = 4, PlannedDoseRateMuPerMin = 700 }
            };
            var runs = (System.Collections.ICollection)RendererMethod("TraceRuns").Invoke(null,
                new object[] { samples, new Func<ReviewControlPointSample, double?>(cp => cp.PlannedDoseRateMuPerMin) });
            TestAssert.Equal(3, runs.Count);
        }

        public static void ReportLabelsControlPointTracesAndSelectedBev()
        {
            string ddl = DdlWriter.WriteToString(CreateDocument());
            TestAssert.True(ddl.Contains("Control-point trajectories"));
            TestAssert.False(ddl.Contains("Nominal: dashed") || ddl.Contains("any dashed line is the nominal"),
                "Nominal MU/min must not be presented as a trajectory in the report.");
            TestAssert.True(ddl.Contains("Nominal MU/min is a field setting, not a trajectory"));
            TestAssert.True(ddl.Contains("Estimated plan trajectory"), "PDF must distinguish the newly supported estimate from supplied planned values.");
            string renderer = File.ReadAllText(Path.Combine("ClearPlan.Reporting.MigraDoc", "Internal", "ReviewImageRenderer.cs"));
            TestAssert.False(renderer.Contains("TraceRuns(samples, cp => cp.NominalDoseRateMuPerMin"));
            TestAssert.True(renderer.Contains("EstimatedDoseRateMuPerMin"));
            TestAssert.True(ddl.Contains("CP 0;"), "BEV reports must show explicitly labeled beam-start geometry.");
            TestAssert.True(ddl.Contains("Field start"), "Start apertures are not representative arc-integrated fluence.");
            TestAssert.True(ddl.Contains("not the source volume for the BEV DRRs"), "Synthetic slices and DRRs use distinct fixtures and must not imply shared anatomy.");
        }

        private static Document CreateDocument()
        {
            var snapshot = SyntheticScenarioFactory.Create("mixed-review");
            snapshot.PlanImages = SyntheticPlanImageFactory.Create(snapshot.ActivePlanKey);
            snapshot.PlanAnalysis = SyntheticPlanAnalysisFactory.Create(true, true);
            return (Document)typeof(ReportPdf).GetMethod("CreateReviewReport", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { new ReviewSnapshotReportMapper().Map(snapshot) });
        }

        private static MethodInfo RendererMethod(string name)
        {
            var type = typeof(ReportPdf).Assembly.GetType("ClearPlan.Reporting.MigraDoc.Internal.ReviewImageRenderer");
            MethodInfo method = type.GetMethod(name, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            TestAssert.NotNull(method, "Required report policy is missing: " + name);
            return method;
        }
    }
}
