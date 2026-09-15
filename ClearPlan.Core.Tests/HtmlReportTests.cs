using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text.RegularExpressions;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Reporting;
using ClearPlan.Reporting.MigraDoc;

namespace ClearPlan.Core.Tests
{
    internal static class HtmlReportTests
    {
        public static void AllSectionsAndDoseStatisticsArePresent()
        {
            var document = Fixture();
            string html = WebUtility.HtmlDecode(Render(document));
            foreach (string label in new[] { "Clinical goals", "PlanCheck", "Structure mapping", "DVH", "Plan parameters", "Target quality", "CT overview", "Beam's-eye views", "Volume [cm³]", "Dmin [Gy]", "D98 [Gy]", "Dmean [Gy]", "D50 [Gy]", "Dmax [Gy]", "D2 [Gy]", "Paddick CI", "1 / Paddick CI", "GI", "HI", "MU", "PAM", "Current snapshot", "2026-09-08 10:00 UTC", "SYNTHETIC", "NOT FOR CLINICAL USE" })
                TestAssert.True(html.Contains(label), "Missing HTML report content: " + label);
            TestAssert.True(html.Contains("12.345") && html.Contains("51.25") && html.Contains("59.75"));
            TestAssert.False(html.Contains("Unit range") || html.Contains(">Samples<"));
            TestAssert.True(html.Contains("Control-point trajectories"), "HTML quicklook must include the same rate/aperture plots as PDF.");
            TestAssert.True(html.Contains("Estimated plan trajectory"), "The report must qualify estimated rather than measured delivery.");
            document.PlanAnalysis.Beams[0].DoseRateEstimateReason = "Synthetic assumption <not approved>";
            string withAssumptions = Render(document);
            TestAssert.True(withAssumptions.Contains("Model assumptions / availability") && withAssumptions.Contains("Synthetic assumption &lt;not approved&gt;"));
        }

        public static void HostileTextCannotBecomeMarkupOrAttributes()
        {
            var document = Fixture();
            const string hostile = "\"><img src=x onerror=alert(1)><script>alert('ä')</script>& Ω";
            document.PatientDisplayLabel = hostile;
            document.PlanDisplayLabel = hostile;
            document.PqmRows[0].MappingDescription = hostile;
            document.PlanCheckRows[0].Message = hostile;
            document.DvhSeries[0].DisplayName = hostile;
            document.DvhSeries[0].ColorHex = "red\" onload=\"alert(1)";
            document.PlanImages = SyntheticPlanImageFactory.Create(document.ActivePlanKey);
            document.PlanImages[0].Title = hostile;
            string html = Render(document);
            TestAssert.False(html.Contains("<script") || html.Contains("<img src=x") || html.Contains(" onload=\""));
            TestAssert.True(html.Contains("&quot;&gt;&lt;img") && html.Contains("&amp; Ω"));
            foreach (Match tag in Regex.Matches(html, @"<[^>]+>"))
                foreach (Match attribute in Regex.Matches(tag.Value, @"\s([a-z][a-z0-9:-]*)\s*=\s*(?:""[^""]*""|'[^']*'|[^\s>]+)", RegexOptions.IgnoreCase))
                    TestAssert.False(attribute.Groups[1].Value.StartsWith("on", StringComparison.OrdinalIgnoreCase), "Untrusted event-handler attribute.");
        }

        public static void GoalsOmitReferencesButKeepResultAndMapping()
        {
            var document = Fixture();
            document.PqmRows[0].SourceLabel = "Reference paper UNIQUE_PAPER_DO_NOT_PRINT";
            document.PqmRows[0].MappingDescription = "Exact synthetic target match";
            string html = Render(document);
            TestAssert.False(html.Contains("UNIQUE_PAPER_DO_NOT_PRINT") || html.Contains("<th scope=\"col\">Source</th>"));
            TestAssert.False(html.Contains("Exact synthetic target match") || html.Contains("Mapping / explanation"));
            TestAssert.True(html.Contains(document.PqmRows[0].ResolvedStructureId) && html.Contains("59.75"));
            TestAssert.True(document.PqmRows[0].SourceLabel.Contains("UNIQUE_PAPER_DO_NOT_PRINT"));
        }

        public static void CompactOptionsKeepUnavailabilityAndDoNotMutateData()
        {
            var document = Fixture(); document.PlanImages.Clear(); document.DvhSeries.Clear();
            document.PqmRows = new List<ReviewReportPqmRow> {
                new ReviewReportPqmRow { TemplateStructure = "UNMATCHED_GOAL", ResolvedStructureId = " ", Status = "not-evaluated" },
                new ReviewReportPqmRow { TemplateStructure = "MATCHED_MISSING_DVH", ResolvedStructureId = "Heart", Objective = "Dmean", Status = "not-evaluated", Explanation = "DVH unavailable." }
            };
            document.StructureMappings = new List<ReviewReportStructureMappingRow> {
                new ReviewReportStructureMappingRow { TemplateStructure = "UNMATCHED_MAPPING", SelectedStructureId = "" },
                new ReviewReportStructureMappingRow { TemplateStructure = "MATCHED_MAPPING", SelectedStructureId = "Heart" }
            };
            SetOption(document, "HideUnmatched", true); SetOption(document, "DisabledCheckCount", 3);
            SetOption(document, "IncludeBeamEyeViews", false);
            document.PlanCheckRows[0].Message = "CHECK_MESSAGE_FIRST";
            document.PlanCheckRows[0].ObservedValue = "OBSERVED_DETAIL_ONLY";
            document.PlanCheckRows[0].ExpectedValue = "EXPECTED_DETAIL_ONLY";
            string html = Render(document);
            TestAssert.False(html.Contains("UNMATCHED_GOAL") || html.Contains("UNMATCHED_MAPPING"));
            TestAssert.True(html.Contains("MATCHED_MISSING_DVH") && html.Contains("DVH unavailable") && html.Contains("not-evaluated"));
            TestAssert.True(html.Contains("1 unmatched goal hidden") && html.Contains("3 checks disabled"));
            TestAssert.False(html.Contains("Beam's-eye views") || html.Contains("OBSERVED_DETAIL_ONLY") || html.Contains("EXPECTED_DETAIL_ONLY"));
            int checks = html.IndexOf("id=\"checks\"", StringComparison.Ordinal);
            string checkSection = html.Substring(checks, html.IndexOf("</section>", checks, StringComparison.Ordinal) - checks);
            TestAssert.True(checkSection.IndexOf(">Message<", StringComparison.Ordinal) < checkSection.IndexOf(">Status<", StringComparison.Ordinal));
            TestAssert.Equal(2, document.PqmRows.Count); TestAssert.Equal(2, document.StructureMappings.Count);
            TestAssert.Equal("OBSERVED_DETAIL_ONLY", document.PlanCheckRows[0].ObservedValue);
            SetOption(document, "HideUnmatched", false);
            TestAssert.True(Render(document).Contains("UNMATCHED_GOAL"));
        }

        public static void BeamViewsUseExactStartOrStayCompactlyUnavailable()
        {
            var document = Fixture(); document.DvhSeries.Clear(); document.PlanImages.Clear();
            var beam = document.PlanAnalysis.Beams[0];
            var later = beam.ControlPoints[1];
            later.BevImage = new BeamEyeViewImage { WidthPixels = 8, HeightPixels = 8, GrayscalePixels = Enumerable.Repeat((byte)100, 64).ToArray(), ExtentMm = 220,
                Synthetic = true, SourceStatus = "synthetic", ControlPointIndex = later.Index, GantryAngleDegrees = later.GantryAngleDegrees,
                CollimatorAngleDegrees = later.CollimatorAngleDegrees, PatientSupportAngleDegrees = later.PatientSupportAngleDegrees };
            string html = Render(document);
            TestAssert.Equal(0, ImageCount(html, "bev-raster"), "A later cached CP must never replace the field-start view.");
            TestAssert.True(html.Contains("Field start") && html.Contains("CP 0"));
            var first = beam.ControlPoints[0];
            first.GantryAngleDegrees = 123.4;
            first.BevImage = new BeamEyeViewImage { WidthPixels = 8, HeightPixels = 8, GrayscalePixels = Enumerable.Repeat((byte)100, 64).ToArray(), ExtentMm = 220,
                Synthetic = true, SourceStatus = "synthetic", ControlPointIndex = 0, GantryAngleDegrees = first.GantryAngleDegrees,
                CollimatorAngleDegrees = first.CollimatorAngleDegrees, PatientSupportAngleDegrees = first.PatientSupportAngleDegrees };
            beam.ControlPoints.Reverse();
            html = Render(document);
            TestAssert.True(ImageCount(html, "bev-raster") == 1 && html.Contains("123.4"), "Start means CP index 0, not the first list item.");
        }

        private static void SetOption(ReviewReportDocument document, string name, object value)
        {
            var property = typeof(ReviewReportDocument).GetProperty(name);
            TestAssert.NotNull(property, "Missing report option: " + name);
            property.SetValue(document, value, null);
        }

        public static void MissingAndNonfiniteValuesRemainUnavailable()
        {
            var document = new ReviewReportDocument { PatientDisplayLabel = "Synthetic ÄÖÜ ß", PlanDisplayLabel = "Synthetic plan" };
            document.PlanAnalysis = new ReviewPlanAnalysis { TotalMetersetMu = double.NaN, Pam = double.PositiveInfinity };
            document.DvhSeries.Add(new ReviewReportDvhSeries { DisplayName = "Synthetic invalid", Selected = true, VolumeCc = double.NaN,
                Statistics = new ReviewDvhStatistics { MeanDoseGy = double.NegativeInfinity },
                Points = new List<ReviewReportDvhPoint> { new ReviewReportDvhPoint { DoseGy = double.NaN, VolumePercent = 100 }, new ReviewReportDvhPoint { DoseGy = 1, VolumePercent = 0 } } });
            string html = Render(document);
            TestAssert.True(html.Contains("Unavailable") && WebUtility.HtmlDecode(html).Contains("Synthetic ÄÖÜ ß"));
            TestAssert.False(html.Contains("NaN") || html.Contains("Infinity") || html.Contains("<td>0</td>"));
            TestAssert.False(html.Contains("data:image/png;base64,"), "An invalid DVH must not produce a misleading curve.");
            document.Plans = null; document.PqmRows = null; document.PlanCheckRows = null; document.StructureMappings = null;
            document.DvhSeries = null; document.PlanImages = null; document.FieldRows = null; document.Notes = null;
            TestAssert.True(Render(document).Contains("Unavailable"));
        }

        public static void OutputHasNoExternalDependencies()
        {
            string html = Render(Fixture());
            TestAssert.True(html.StartsWith("<!DOCTYPE html>", StringComparison.OrdinalIgnoreCase));
            TestAssert.True(html.Contains("X-UA-Compatible") && html.Contains("IE=edge") && html.Contains("@media print"));
            TestAssert.False(Regex.IsMatch(html, @"<(?:script|link|iframe|object|embed|base)\b", RegexOptions.IgnoreCase));
            TestAssert.False(Regex.IsMatch(html, @"(?:src|href)\s*=\s*[""'](?:https?:|//|file:|javascript:)", RegexOptions.IgnoreCase));
            TestAssert.False(html.Contains("@import") || html.Contains("url(") || html.Contains("@font-face"));
            TestAssert.True(html.Contains("data:image/png;base64,") && html.Contains("overflow-x:auto"));
        }

        public static void CurrentPlanAndRequiredTargetsAreNotLost()
        {
            var document = Fixture();
            document.Plans.Add(new ReviewReportPlanRow { PlanKey = "other-plan", DisplayLabel = "OTHER_PLAN_MUST_NOT_APPEAR" });
            document.PlanImages = SyntheticPlanImageFactory.Create("other-plan");
            document.PlanImages[0].Title = "OTHER_CT_MUST_NOT_APPEAR";
            document.DvhSeries[0].Selected = false;
            document.DvhSeries[0].RequiredForTargetReview = true;
            document.DvhSeries[0].DisplayName = "REQUIRED_SYNTHETIC_PTV";
            string html = Render(document);
            TestAssert.False(html.Contains("OTHER_PLAN_MUST_NOT_APPEAR") || html.Contains("OTHER_CT_MUST_NOT_APPEAR"));
            TestAssert.True(html.Contains("REQUIRED_SYNTHETIC_PTV") && html.Contains("data:image/png;base64,"));
        }

        public static void ImagesAreEmbeddedOnlyFromAvailableDetachedPixels()
        {
            var document = Fixture();
            document.DvhSeries.Clear();
            document.PlanImages = SyntheticPlanImageFactory.Create(document.ActivePlanKey);
            document.PlanImages[0].Caption = "CAPTURE_ALGORITHM_PROSE_MUST_NOT_PRINT";
            document.PlanImages[0].OverlaySummary = "OVERLAY_TECHNICAL_PROSE_MUST_NOT_PRINT";
            document.PlanImages[0].Overlays.Add(new ReviewImageOverlay { Kind = "structure", SourceStatus = "unavailable", UnavailableReason = "Detailed capture technical reason" });
            document.PlanImages[0].Overlays.Add(new ReviewImageOverlay { Kind = "isodose", SourceStatus = "unavailable", UnavailableReason = "Detailed dose technical reason" });
            var legendPath = new ReviewImagePath { Points = new List<ReviewImagePoint> {
                new ReviewImagePoint { X = 10, Y = 10 }, new ReviewImagePoint { X = 60, Y = 60 } } };
            document.PlanImages[0].Overlays[0].Label = "UNAVAILABLE_CONTOUR_MUST_NOT_APPEAR";
            document.PlanImages[0].Overlays[0].Paths.Add(legendPath);
            foreach (var overlay in new[] {
                new ReviewImageOverlay { Kind = "structure", Label = "Cord <critical> & left", ColorHex = "#123456", SourceStatus = ReviewStatusCodes.Available },
                new ReviewImageOverlay { Kind = "isodose", Label = "95% Rx / 57 Gy", ColorHex = "#D75055", SourceStatus = ReviewStatusCodes.Available },
                new ReviewImageOverlay { Kind = "structure", Label = "LEGEND_CAN_HIDE", ColorHex = "#ABCDEF", SourceStatus = ReviewStatusCodes.Available },
                new ReviewImageOverlay { Kind = "structure", Label = "SAFE_COLOR_FALLBACK", ColorHex = "#fff\" onmouseover=\"unsafe", SourceStatus = ReviewStatusCodes.Available } })
            { overlay.Paths.Add(legendPath); document.PlanImages[0].Overlays.Add(overlay); }
            document.PlanImages[0].Overlays.Add(new ReviewImageOverlay { Kind = "structure", Label = "NO_PATH_MUST_NOT_APPEAR", SourceStatus = ReviewStatusCodes.Available });
            document.PlanImages[0].Overlays.Add(new ReviewImageOverlay { Kind = "structure", Label = "EMPTY_PATH_MUST_NOT_APPEAR", SourceStatus = ReviewStatusCodes.Available,
                Paths = new List<ReviewImagePath> { new ReviewImagePath() } });
            var beam = document.PlanAnalysis.Beams[0];
            var cp = beam.ControlPoints[0];
            cp.BevImage = new BeamEyeViewImage { WidthPixels = 8, HeightPixels = 8, GrayscalePixels = Enumerable.Repeat((byte)100, 64).ToArray(), ExtentMm = 220,
                Synthetic = true, SourceStatus = "synthetic", ControlPointIndex = cp.Index, GantryAngleDegrees = cp.GantryAngleDegrees,
                CollimatorAngleDegrees = cp.CollimatorAngleDegrees, PatientSupportAngleDegrees = cp.PatientSupportAngleDegrees };
            string html = Render(document);
            TestAssert.Equal(3, ImageCount(html, "ct-raster"));
            TestAssert.Equal(1, ImageCount(html, "bev-raster"));
            TestAssert.Equal(2, ImageCount(html, "parameter-traces"));
            TestAssert.False(html.Contains("CAPTURE_ALGORITHM_PROSE_MUST_NOT_PRINT") || html.Contains("OVERLAY_TECHNICAL_PROSE_MUST_NOT_PRINT"));
            TestAssert.True(html.Contains("1 contour unavailable") && html.Contains("1 isodose unavailable"));
            TestAssert.True(html.Contains("class=\"ct-legend\"") && html.Contains("Contours") && html.Contains("Isodoses"), "Available CT overlays need a compact typed color legend.");
            TestAssert.True(html.Contains("Cord &lt;critical&gt; &amp; left") && html.Contains("95% Rx / 57 Gy"));
            TestAssert.True(html.Contains("border-top-color:#123456") && html.Contains("border-top-color:#D75055"));
            TestAssert.True(html.Contains("LEGEND_CAN_HIDE") && html.Contains("SAFE_COLOR_FALLBACK") && html.Contains("border-top-color:#0F766E"));
            TestAssert.False(html.Contains("Cord <critical>") || html.Contains("onmouseover=") || html.Contains("NO_PATH_MUST_NOT_APPEAR") || html.Contains("EMPTY_PATH_MUST_NOT_APPEAR") || html.Contains("UNAVAILABLE_CONTOUR_MUST_NOT_APPEAR"));
            document.HiddenStructureIds.Add("LEGEND_CAN_HIDE");
            string hidden = Render(document);
            TestAssert.False(hidden.Contains("LEGEND_CAN_HIDE"), "The CT legend must use the same filtered overlays as its image.");
            TestAssert.True(document.PlanImages[0].Overlays.Any(overlay => overlay.Label == "LEGEND_CAN_HIDE"), "Visibility filtering must not mutate source contours.");
            document.PlanImages[0].SourceStatus = "unavailable";
            document.PlanImages[1].WidthPixels = int.MaxValue;
            document.PlanImages[2].GrayscalePixels = null;
            cp.BevImage.GrayscalePixels = null;
            string invalidImages = Render(document);
            TestAssert.Equal(0, ImageCount(invalidImages, "ct-raster") + ImageCount(invalidImages, "bev-raster"));
            TestAssert.False(invalidImages.Contains("class=\"ct-legend\"") || invalidImages.Contains("95% Rx / 57 Gy"), "Unavailable images must not display a misleading overlay legend.");
        }

        public static void NullInputAndUtf8ExportAreExplicit()
        {
            var type = RendererType();
            var method = type.GetMethod("Render");
            var error = TestAssert.Throws<TargetInvocationException>(() => method.Invoke(Activator.CreateInstance(type), new object[] { null }));
            TestAssert.True(error.InnerException is ArgumentNullException);
            string path = Path.Combine(Path.GetTempPath(), "ClearPlan-Synthetic-Html-" + Guid.NewGuid().ToString("N") + ".html");
            try
            {
                var document = Fixture(); document.PatientDisplayLabel = "Synthetic ÄÖÜ ß Ω";
                type.GetMethod("Export").Invoke(Activator.CreateInstance(type), new object[] { path, document });
                TestAssert.Equal(Render(document), File.ReadAllText(path));
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }

        public static void EverySelectedCurveHasAChartPanel()
        {
            var document = Fixture();
            document.DvhSeries.Clear();
            for (int i = 0; i < 14; i++) document.DvhSeries.Add(new ReviewReportDvhSeries {
                DisplayName = "Synthetic structure " + i, Selected = true, DoseUnit = "Gy", VolumeUnit = "%",
                Points = new List<ReviewReportDvhPoint> { new ReviewReportDvhPoint { DoseGy = 0, VolumePercent = 100 }, new ReviewReportDvhPoint { DoseGy = 60, VolumePercent = 0 } }
            });
            string html = Render(document);
            TestAssert.Equal(1, ImageCount(html, "plot"), "All selected curves share one DVH and complete wrapped legend.");
            TestAssert.False(html.Contains("DVH panel"));
            for (int i=0;i<14;i++) TestAssert.True(html.Contains("Synthetic structure " + i));
            var staircase = document.DvhSeries[0]; document.DvhSeries = new List<ReviewReportDvhSeries> { staircase };
            staircase.Points = new List<ReviewReportDvhPoint> {
                new ReviewReportDvhPoint { DoseGy=0, VolumePercent=100 }, new ReviewReportDvhPoint { DoseGy=10, VolumePercent=100 },
                new ReviewReportDvhPoint { DoseGy=10, VolumePercent=50 }, new ReviewReportDvhPoint { DoseGy=20, VolumePercent=0 } };
            TestAssert.Equal(1, ImageCount(Render(document),"plot"),"Equal-dose staircase points are valid, not a missing curve.");
            TestAssert.Equal(4, staircase.Points.Count,"Do not reorder or deduplicate native DVH samples.");
        }

        public static void EveryCachedBeamHasAnImagePanel()
        {
            var document = Fixture(); document.DvhSeries.Clear();
            var beam = document.PlanAnalysis.Beams[0]; var cp = beam.ControlPoints[0];
            cp.BevImage = new BeamEyeViewImage { WidthPixels = 8, HeightPixels = 8, GrayscalePixels = Enumerable.Repeat((byte)100, 64).ToArray(), ExtentMm = 220,
                Synthetic = true, SourceStatus = "synthetic", ControlPointIndex = cp.Index, GantryAngleDegrees = cp.GantryAngleDegrees,
                CollimatorAngleDegrees = cp.CollimatorAngleDegrees, PatientSupportAngleDegrees = cp.PatientSupportAngleDegrees };
            document.PlanAnalysis.Beams = Enumerable.Repeat(beam, 14).ToList();
            string html = Render(document);
            TestAssert.Equal(14, ImageCount(html, "bev-raster"), "Available fields must not be silently truncated.");
            TestAssert.Equal(14, ImageCount(html, "parameter-traces"));
            TestAssert.Equal(14, Regex.Matches(html, "class=\"field-review\"").Count,
                "Each field must group its BEV and stacked dose-rate/aperture plots in one printable panel.");
            int firstField = html.IndexOf("class=\"field-review\"", StringComparison.Ordinal);
            TestAssert.True(html.IndexOf("class=\"parameter-traces\"", StringComparison.Ordinal) > firstField,
                "Parameter traces must live with the field, not on a separate preceding page.");
        }

        public static void CompanionExportNeverOverwritesExistingFiles()
        {
            var type = RendererType();
            var method = type.GetMethod("ExportCompanion");
            TestAssert.NotNull(method, "Companion export API is missing.");
            string pdfPath = Path.Combine(Path.GetTempPath(), "ClearPlan-Synthetic-Companion-" + Guid.NewGuid().ToString("N") + ".pdf");
            string htmlPath = Path.ChangeExtension(pdfPath, ".html");
            string secondPath = null;
            try
            {
                File.WriteAllText(pdfPath, "synthetic PDF sentinel");
                var instance = Activator.CreateInstance(type);
                TestAssert.Equal(htmlPath, (string)method.Invoke(instance, new object[] { pdfPath, Fixture() }));
                File.WriteAllText(htmlPath, "existing HTML sentinel");
                secondPath = (string)method.Invoke(instance, new object[] { pdfPath, Fixture() });
                TestAssert.False(string.Equals(htmlPath, secondPath, StringComparison.OrdinalIgnoreCase));
                TestAssert.Equal(Path.GetDirectoryName(pdfPath), Path.GetDirectoryName(secondPath));
                TestAssert.Equal("synthetic PDF sentinel", File.ReadAllText(pdfPath));
                TestAssert.Equal("existing HTML sentinel", File.ReadAllText(htmlPath));
                TestAssert.True(File.ReadAllText(secondPath).Contains("Current snapshot"));
            }
            finally
            {
                foreach (string path in new[] { pdfPath, htmlPath, secondPath }) if (path != null && File.Exists(path)) File.Delete(path);
            }
        }

        public static void ExtremeFiniteCurvesAreWithheldAndMetricScopeIsVisible()
        {
            var document = Fixture();
            document.DvhSeries.Clear();
            document.DvhSeries.Add(new ReviewReportDvhSeries { DisplayName = "Synthetic overflow", Selected = true,
                Points = new List<ReviewReportDvhPoint> { new ReviewReportDvhPoint { DoseGy = 0, VolumePercent = 100 }, new ReviewReportDvhPoint { DoseGy = double.MaxValue, VolumePercent = 0 } } });
            string html = Render(document);
            TestAssert.Equal(0, ImageCount(html, "ct-raster") + ImageCount(html, "bev-raster") + ImageCount(html, "plot"), "Finite inputs must also stay inside raster-safe bounds.");
            TestAssert.True(WebUtility.HtmlDecode(html).Contains("Small aperture: < 4 cm²"));
            TestAssert.True(html.Contains(document.PlanAnalysis.PamWeightingMode));
        }

        private static int ImageCount(string html, string css)
        {
            return Regex.Matches(html, "<img\\b[^>]*class=\"" + Regex.Escape(css) + "\"[^>]*src=\"data:image/png;base64,").Count;
        }

        private static Type RendererType()
        {
            var type = typeof(ReportPdf).Assembly.GetType("ClearPlan.Reporting.MigraDoc.HtmlReviewReportRenderer");
            TestAssert.NotNull(type, "HTML quicklook renderer is not implemented.");
            return type;
        }

        private static string Render(ReviewReportDocument document)
        {
            var type = RendererType();
            return (string)type.GetMethod("Render").Invoke(Activator.CreateInstance(type), new object[] { document });
        }

        private static ReviewReportDocument Fixture()
        {
            var snapshot = SyntheticScenarioFactory.Create("mixed-review");
            snapshot.PlanAnalysis = SyntheticPlanAnalysisFactory.Create(true, true);
            var document = new ReviewSnapshotReportMapper().Map(snapshot);
            document.GeneratedUtc = new DateTimeOffset(2026, 9, 8, 10, 0, 0, TimeSpan.Zero);
            document.PqmRows[0].AchievedValue = 59.75;
            document.DvhSeries[0].VolumeCc = 12.345;
            document.DvhSeries[0].Selected = true;
            document.DvhSeries[0].Statistics.MeanDoseGy = 51.25;
            return document;
        }
    }
}
