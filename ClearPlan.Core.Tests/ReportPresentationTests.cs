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
        public static void DvhLegendWrapsBelowAndRetainsEverySeries()
        {
            var data = DvhLegendFixture(24);
            string encoded = (string)RendererMethod("Dvh").Invoke(null, new object[] { data });
            using (var stream = new MemoryStream(Convert.FromBase64String(encoded.Substring("base64:".Length))))
            using (var bitmap = new System.Drawing.Bitmap(stream))
            {
                TestAssert.True(bitmap.Height > 500, "The DVH needs a separate wrapped legend below the full-width plot.");
                foreach (var row in data.DvhSeries)
                {
                    var color = System.Drawing.ColorTranslator.FromHtml(row.ColorHex);
                    bool swatch = false;
                    for (int y = 490; y < bitmap.Height - 10 && !swatch; y++)
                        for (int x = 20; x < bitmap.Width - 20; x++)
                            if (bitmap.GetPixel(x, y).ToArgb() == color.ToArgb()) { swatch = true; break; }
                    TestAssert.True(swatch, "Every selected curve needs its legend swatch below the plot, including " + row.StructureId);
                }
                for (int y = 490; y < bitmap.Height; y++)
                    TestAssert.Equal(System.Drawing.Color.White.ToArgb(), bitmap.GetPixel(bitmap.Width - 1, y).ToArgb(),
                        "Long structure labels must wrap inside the image, never clip at its right edge.");
            }
            TestAssert.Equal(24, data.DvhSeries.Count, "Rendering must not trim the source inventory.");
            TestAssert.True(data.DvhSeries[0].DisplayName.EndsWith("TAIL_0"), "Long labels must not be shortened in the source.");
        }

        public static void PdfPaginatesEveryDvhCurveWithReadableLegends()
        {
            var data = DvhLegendFixture(27);
            var document = (Document)typeof(ReportPdf).GetMethod("CreateReviewReport", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { data });
            string ddl = DdlWriter.WriteToString(document);
            TestAssert.False(ddl.Contains("DVH panel"), "Report must contain one combined DVH, not a second panel after thirteen curves.");
            string html = new HtmlReviewReportRenderer().Render(data);
            TestAssert.False(html.Contains("DVH panel"), "HTML must use the same combined DVH.");
            TestAssert.Equal(27, data.DvhSeries.Count, "All selected curves must remain available.");
            // Exercise the actual landscape pagination with long IDs, independently
            // of unrelated tables and BEV pages. A chart must not leave its label behind.
            var section = document.Sections[0];
            var elements = section.Elements.Cast<DocumentObject>().ToList();
            int start = elements.FindIndex(item => item is Paragraph && DdlWriter.WriteToString(item).Contains("Dose-volume overview"));
            int end = elements.FindIndex(start + 1, item => item is Paragraph && DdlWriter.WriteToString(item).Contains("DVH statistics"));
            TestAssert.True(start >= 0 && end > start);
            section.Elements.Clear();
            foreach (var item in elements.Skip(start).Take(end - start)) section.Elements.Add((DocumentObject)item.Clone());
            var renderer = new MigraDoc.Rendering.PdfDocumentRenderer(true) { Document = document };
            renderer.RenderDocument();
            TestAssert.Equal(1, section.Elements.Cast<DocumentObject>().OfType<MigraDoc.DocumentObjectModel.Shapes.Image>().Count());
            TestAssert.Equal(1, renderer.PdfDocument.PageCount, "One DVH with a complete wrapped legend must fit one page.");
        }

        private static ReviewReportDocument DvhLegendFixture(int count)
        {
            var data = new ReviewSnapshotReportMapper().Map(SyntheticScenarioFactory.Create("baseline-pass"));
            data.LanguageCode = "en";
            data.DvhSeries = Enumerable.Range(0, count).Select(index => new ReviewReportDvhSeries {
                StructureId = "DEMO_" + index,
                DisplayName = "DEMO_" + new string('L', index % 3 == 0 ? 95 : 12) + "_TAIL_" + index,
                ColorHex = "#" + (32 + index * 7).ToString("X2") + "40B0",
                LineStyle = index % 2 == 0 ? "solid" : "dash", Selected = true,
                Points = new List<ReviewReportDvhPoint> { new ReviewReportDvhPoint { DoseGy=0, VolumePercent=100 },
                    new ReviewReportDvhPoint { DoseGy=50 + index, VolumePercent=0 } }
            }).ToList();
            return data;
        }

        public static void StableTitleAndVersionReplaceDecorativeWarningBanners()
        {
            var empty = new Document(); var emptySection=empty.AddSection();
            var staticAnalysis=new ReviewPlanAnalysis();
            var staticBeam=new ReviewBeamAnalysis { BeamId="STATIC", NominalDoseRateMuPerMin=600 };
            staticBeam.ControlPoints.Add(new ReviewControlPointSample { Index=0,NominalDoseRateMuPerMin=600 });
            staticBeam.ControlPoints.Add(new ReviewControlPointSample { Index=1,NominalDoseRateMuPerMin=600 });
            staticAnalysis.Beams.Add(staticBeam);
            typeof(ReportPdf).GetMethod("AddFieldReviews",BindingFlags.Instance|BindingFlags.NonPublic)
                .Invoke(new ReportPdf(),new object[] { emptySection,staticAnalysis,false,false });
            TestAssert.Equal(0,emptySection.Elements.Count,"Do not reserve full pages for unavailable rate trajectories.");
            var data = new ReviewSnapshotReportMapper().Map(SyntheticScenarioFactory.Create("baseline-pass"));
            data.Synthetic = false;
            data.Title = "Do not use this changing title";
            data.ModeLabel = "Local review; incomplete geometry is not clearance.";
            data.Watermark = data.ModeLabel;
            var document = (Document)typeof(ReportPdf).GetMethod("CreateReviewReport", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { data });
            TestAssert.Equal("Plan Quality Report", document.Info.Title);
            string ddl = DdlWriter.WriteToString(document);
            TestAssert.False(ddl.Contains(data.Title));
            TestAssert.True(ddl.Contains("ClearPlan") && ddl.Contains("3.2.0"), "Repeating footer must identify the software version.");
            var first = document.Sections[0].Elements[0] as Paragraph;
            TestAssert.True(first != null && !first.Format.Font.Color.Equals(Colors.White), "No decorative white-on-warning banner.");
            var html = new HtmlReviewReportRenderer().Render(data);
            TestAssert.True(html.Contains("<h1>Plan Quality Report</h1>"));
            TestAssert.True(html.Contains("<footer>ClearPlan") && html.Contains("3.2.0"));
            TestAssert.True(html.Contains("incomplete geometry"), "Quiet styling must not hide unavailable clinical data.");
        }

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

        public static void PaginatedTablesStayAboveTheThreeLineFooter()
        {
            var document = CreateDocument();
            var section = document.Sections[0];
            section.Elements.Clear();
            // A small preceding block offsets row boundaries, as a DVH image does.
            // Do not accidentally pass only because a row leaves spare bottom space.
            var preceding = section.AddParagraph(" ");
            preceding.Format.LineSpacingRule = LineSpacingRule.Exactly;
            preceding.Format.LineSpacing = Unit.FromPoint(4);
            preceding.Format.SpaceAfter = Unit.FromPoint(0);
            var table = section.AddTable();
            table.AddColumn(Unit.FromCentimeter(18));
            table.AddColumn(Unit.FromCentimeter(9));
            for (int index = 0; index < 90; index++)
            {
                var row = table.AddRow();
                row.Cells[0].AddParagraph("Synthetic statistics row " + index);
                row.Cells[1].AddParagraph("12.34 / 56.78");
            }
            var renderer = new MigraDoc.Rendering.PdfDocumentRenderer(true) { Document = document };
            renderer.RenderDocument();
            TestAssert.True(renderer.PdfDocument.PageCount > 1, "The regression must exercise a full body and table continuation.");
            AssertPageBodyClearsFooter(renderer, section);
        }

        private static void AssertPageBodyClearsFooter(MigraDoc.Rendering.PdfDocumentRenderer renderer, Section section)
        {
            var formatted = renderer.DocumentRenderer.FormattedDocument;
            var footerMethod = formatted.GetType().GetMethod("GetFormattedFooter", BindingFlags.Instance | BindingFlags.NonPublic);
            TestAssert.NotNull(footerMethod);
            for (int page = 1; page <= renderer.PdfDocument.PageCount; page++)
            {
                var footer = footerMethod.Invoke(formatted, new object[] { page });
                var footerHeight = (PdfSharp.Drawing.XUnit)footer.GetType().GetProperty("ContentHeight",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).GetValue(footer, null);
                double footerTop = renderer.PdfDocument.Pages[page - 1].Height.Point -
                    section.PageSetup.FooterDistance.Point - footerHeight.Point;
                double bodyBottom = renderer.DocumentRenderer.GetRenderInfoFromPage(page)
                    .Max(info => info.LayoutInfo.ContentArea.Y.Point + info.LayoutInfo.ContentArea.Height.Point);
                TestAssert.True(bodyBottom + 4 <= footerTop,
                    "Page " + page + " body must leave at least 4 pt before the complete identity/version/status footer; gap=" +
                    (footerTop - bodyBottom).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + " pt.");
            }
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
            AssertPageBodyClearsFooter(renderer, paginated.Sections[0]);
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
            AssertPageBodyClearsFooter(renderer, document.Sections[0]);
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
            using (ClearPlan.Core.Localization.ReviewLanguage.Scope("en"))
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
            using (ClearPlan.Core.Localization.ReviewLanguage.Scope("en"))
                typeof(ReportPdf).GetMethod("AddPqmRows", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(new ReportPdf(), new object[] { section, rows });
            string ddl = DdlWriter.WriteToString(document);
            TestAssert.False(ddl.Contains("Reference test paper") || ddl.Contains("Another reference paper") || ddl.Contains("S1:"),
                "Clinical report goals must not print literature references or bibliography footnotes.");
            TestAssert.False(ddl.Contains("Source / structure match") || ddl.Contains("Exact name") || ddl.Contains("Exact native structure"));
            TestAssert.Equal(8, section.Elements.Cast<DocumentObject>().OfType<Table>().Single().Columns.Count);
            TestAssert.True(ddl.Contains("Heart") && ddl.Contains("Lung"));
            TestAssert.True(rows[0].SourceLabel.Contains("Reference test paper"), "Report formatting must retain original configuration provenance in the detached data.");
            TestAssert.True(ddl.Contains("Not evaluated") && !ddl.Contains("(Note)") && !ddl.Contains("(info)"),
                "A redundant info suffix must not double every unconfirmed PQM row and orphan report notes.");
            TestAssert.Equal("not-evaluated", rows[1].Status);
            TestAssert.Equal("info", rows[1].Severity);
        }
        public static void BeamStartPanelsStayWithinOnePagePerBeam()
        {
            // Isolate BEV pagination from the length of PQM, source and check tables.
            var document = CreateDocument();
            document.Sections[0].Elements.Clear();
            typeof(ReportPdf).GetMethod("AddFieldReviews", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { document.Sections[0], SyntheticPlanAnalysisFactory.Create(true, true), true, true });
            var panels = document.Sections[0].Elements.Cast<DocumentObject>().OfType<Table>().ToList();
            TestAssert.Equal(2, panels.Count, "Each field needs one shared BEV/trajectory panel, not separate trace and BEV pages.");
            foreach (var panel in panels)
            {
                TestAssert.Equal(2, panel.Columns.Count, "BEV belongs on the left, stacked trajectories on the right.");
                TestAssert.Equal(1, panel.Rows[0].Cells[0].Elements.Cast<DocumentObject>().OfType<MigraDoc.DocumentObjectModel.Shapes.Image>().Count());
                TestAssert.Equal(1, panel.Rows[0].Cells[1].Elements.Cast<DocumentObject>().OfType<MigraDoc.DocumentObjectModel.Shapes.Image>().Count());
            }
            var renderer = new MigraDoc.Rendering.PdfDocumentRenderer(true) { Document = document };
            renderer.RenderDocument();
            // MigraDoc ignores the leading page break: exactly one page per beam.
            if (renderer.PdfDocument.PageCount != 2)
            {
                string output = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bev-pagination-diagnostic.pdf");
                renderer.PdfDocument.Save(output);
            }
            TestAssert.Equal(2, renderer.PdfDocument.PageCount);
            AssertPageBodyClearsFooter(renderer, document.Sections[0]);
            // Exercise the field counts of the current report workload for both
            // single- and dual-layer fixtures, using actual MigraDoc pagination.
            foreach (int count in new[] { 13, 3, 2 })
            {
                var analysis = new ReviewPlanAnalysis();
                for (int index = 0; index < count; index++)
                {
                    var beam = SyntheticPlanAnalysisFactory.Create(count == 2, true).Beams[index % 2];
                    beam.BeamId = "SYNTHETIC_FIELD_" + (index + 1);
                    analysis.Beams.Add(beam);
                }
                document = CreateDocument();
                document.Sections[0].Elements.Clear();
                typeof(ReportPdf).GetMethod("AddFieldReviews", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(new ReportPdf(), new object[] { document.Sections[0], analysis, true, true });
                renderer = new MigraDoc.Rendering.PdfDocumentRenderer(true) { Document = document };
                renderer.RenderDocument();
                TestAssert.Equal(count, renderer.PdfDocument.PageCount,
                    "Every BEV/trajectory field panel must fit exactly one page, including the full footer.");
                AssertPageBodyClearsFooter(renderer, document.Sections[0]);
            }
            string encoded = (string)RendererMethod("ControlPointTraces").Invoke(null,
                new object[] { SyntheticPlanAnalysisFactory.Create().Beams[0] });
            using (var stream = new MemoryStream(Convert.FromBase64String(encoded.Substring(7))))
            using (var image = new System.Drawing.Bitmap(stream))
            {
                TestAssert.True(image.Height > image.Width, "The two trajectory plots must be stacked vertically beside the BEV.");
                int aperturePixels = 0;
                for (int y = 0; y < image.Height; y++)
                for (int x = 0; x < image.Width; x++)
                    if (image.GetPixel(x, y).ToArgb() == System.Drawing.Color.FromArgb(198, 130, 34).ToArgb())
                    {
                        TestAssert.True(y > image.Height / 2, "The effective-aperture curve belongs below the dose-rate plot.");
                        aperturePixels++;
                    }
                TestAssert.True(aperturePixels > 50, "The vertical stack must retain the actual aperture curve.");
            }
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
            typeof(ReportPdf).GetMethod("AddFieldReviews", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(new ReportPdf(), new object[] { section, data.PlanAnalysis, false, true });
            TestAssert.Equal(data.PlanAnalysis.Beams.Count, section.Elements.Cast<DocumentObject>().OfType<Table>().Count(),
                "Missing native DRRs share the trajectory page; they must not add a separate empty image page.");
            TestAssert.True(DdlWriter.WriteToString(empty).Contains("DRR unavailable"));
            foreach (var panel in section.Elements.Cast<DocumentObject>().OfType<Table>())
                TestAssert.Equal(0, panel.Rows[0].Cells[0].Elements.Cast<DocumentObject>().OfType<MigraDoc.DocumentObjectModel.Shapes.Image>().Count(),
                    "Missing native DRRs must remain missing; no synthetic fallback is allowed.");
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
