using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using ClearPlan.Reporting.MigraDoc.Internal;
using ClearPlan.Core.Review;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Reporting.MigraDoc
{
    public class ReportPdf : IReport
    {
        public void Export(string path, ReportData data)
        {
            ExportPdf(path, CreateReport(data));
        }

        public void Export(string path, ReviewReportDocument data)
        {
            if (data == null)
            {
                throw new ArgumentNullException("data");
            }

            ExportPdf(path, CreateReviewReport(data), data.Synthetic
                ? ReviewSnapshotReportMapper.SyntheticWatermark : Safe(data.ModeLabel, "ACTIVE PLAN REVIEW - NOT AN APPROVAL"));
        }

        private void ExportPdf(string path, Document report, string pageMark = null)
        {
            using(var pdf=RenderPdf(report,pageMark)) pdf.Save(path);
        }

        public byte[] RenderToBytes(ReviewReportDocument data)
        {
            if(data==null) throw new ArgumentNullException("data");
            using(var pdf=RenderPdf(CreateReviewReport(data),data.Synthetic ? ReviewSnapshotReportMapper.SyntheticWatermark :
                Safe(data.ModeLabel,"ACTIVE PLAN REVIEW - NOT AN APPROVAL")))
            using(var stream=new System.IO.MemoryStream())
            {pdf.Save(stream,false);return stream.ToArray();}
        }

        private PdfSharp.Pdf.PdfDocument RenderPdf(Document report, string pageMark)
        {
            var pdfRenderer = new PdfDocumentRenderer();
            pdfRenderer.Document = report;
            pdfRenderer.RenderDocument();
            if (!string.IsNullOrWhiteSpace(pageMark))
            {
                // Stamp after pagination: the simulation mark must remain visibly present on every page.
                foreach (PdfSharp.Pdf.PdfPage page in pdfRenderer.PdfDocument.Pages)
                    using (var graphics = PdfSharp.Drawing.XGraphics.FromPdfPage(page, PdfSharp.Drawing.XGraphicsPdfPageOptions.Append))
                        graphics.DrawString(pageMark, new PdfSharp.Drawing.XFont("Segoe UI", 8, PdfSharp.Drawing.XFontStyle.Bold),
                            PdfSharp.Drawing.XBrushes.DarkRed, new PdfSharp.Drawing.XRect(0, 15, page.Width.Point, 15),
                            PdfSharp.Drawing.XStringFormats.Center);
            }
            return pdfRenderer.PdfDocument;
        }

        private Document CreateReport(ReportData data)
        {
            var doc = new Document();
            CustomStyles.Define(doc);
            doc.Add(CreateMainSection(data));
            return doc;
        }

        private Document CreateReviewReport(ReviewReportDocument data)
        {
            var document = new Document();
            CustomStyles.Define(document);
            document.Styles[StyleNames.Normal].Font.Name = "Segoe UI";
            document.Styles[StyleNames.Normal].Font.Size = 9;
            document.Styles[StyleNames.Heading1].Font.Color = Color.FromRgb(18, 48, 70);
            document.Styles[StyleNames.Heading2].Font.Color = Color.FromRgb(18, 92, 112);
            document.Styles[StyleNames.Heading2].ParagraphFormat.Shading.Color = Color.FromRgb(226, 240, 244);
            document.Styles[StyleNames.Heading2].ParagraphFormat.KeepWithNext = true;
            document.Styles[StyleNames.Heading2].ParagraphFormat.SpaceBefore = Unit.FromPoint(4);
            string watermark = data.Synthetic
                ? ReviewSnapshotReportMapper.SyntheticWatermark
                : data.Watermark;

            document.Info.Title = Safe(
                data.Title,
                "ClearPlan review report");
            document.Info.Subject = Safe(watermark, data.ModeLabel);

            var section = new Section();
            SetUpReviewPage(section);
            AddReviewHeaderAndFooter(section, data, watermark);
            AddReviewContents(section, data, watermark);
            document.Add(section);
            return document;
        }

        private Section CreateMainSection(ReportData data)
        {
            var section = new Section();
            SetUpPage(section);
            AddHeaderAndFooter(section, data);
            AddContents(section, data);
            return section;
        }

        private void SetUpPage(Section section)
        {
            section.PageSetup.PageFormat = PageFormat.Letter;
            //section.PageSetup.Orientation = Orientation.Landscape;
            section.PageSetup.LeftMargin = Size.LeftRightPageMargin;
            section.PageSetup.TopMargin = Size.TopBottomPageMargin;
            section.PageSetup.RightMargin = Size.LeftRightPageMargin;
            section.PageSetup.BottomMargin = Size.TopBottomPageMargin;

            section.PageSetup.HeaderDistance = Size.HeaderFooterMargin;
            section.PageSetup.FooterDistance = Size.HeaderFooterMargin;
        }

        private void SetUpReviewPage(Section section)
        {
            section.PageSetup.PageFormat = PageFormat.A4;
            section.PageSetup.Orientation = Orientation.Landscape;
            section.PageSetup.LeftMargin = Unit.FromCentimeter(1.2);
            section.PageSetup.TopMargin = Unit.FromCentimeter(1.8);
            section.PageSetup.RightMargin = Unit.FromCentimeter(1.2);
            section.PageSetup.BottomMargin = Unit.FromCentimeter(1.5);
            section.PageSetup.HeaderDistance = Unit.FromCentimeter(0.6);
            section.PageSetup.FooterDistance = Unit.FromCentimeter(0.6);
        }

        private void AddReviewHeaderAndFooter(
            Section section,
            ReviewReportDocument data,
            string watermark)
        {
            string identity = data.Synthetic
                ? "Synthetic demo patient | ID: DEMO-" + Safe(data.ScenarioId, "example")
                : Safe(data.PatientDisplayLabel, "Name / ID unavailable");
            var patient = section.Footers.Primary.AddParagraph("Patient: " +
                System.Text.RegularExpressions.Regex.Replace(identity, @"\s+", " ").Trim());
            patient.Format.Font.Size = 8;
            patient.Format.Font.Bold = true;
            patient.Format.SpaceAfter = Unit.FromPoint(1);
            var footer = section.Footers.Primary.AddParagraph();
            footer.Format.Font.Size = 8;
            footer.Format.AddTabStop(
                Unit.FromCentimeter(25.0),
                TabAlignment.Right);
            footer.AddText(
                "ClearPlan · " +
                Safe(data.ScenarioId, "detached review"));
            footer.AddTab();
            footer.AddText("Page ");
            footer.AddPageField();
            footer.AddText(" of ");
            footer.AddNumPagesField();
        }

        private void AddReviewContents(
            Section section,
            ReviewReportDocument data,
            string watermark)
        {
            Paragraph banner = section.AddParagraph(
                Safe(watermark, data.ModeLabel));
            banner.Format.Alignment = ParagraphAlignment.Center;
            banner.Format.Font.Size = data.Synthetic ? 15 : 10;
            banner.Format.Font.Bold = true;
            banner.Format.Font.Color = Colors.White;
            banner.Format.Shading.Color = data.Synthetic ? Colors.DarkRed : Color.FromRgb(18, 48, 70);
            banner.Format.SpaceAfter = Unit.FromCentimeter(0.3);

            Paragraph title = section.AddParagraph(
                Safe(data.Title, "ClearPlan | Active plan report"));
            title.Style = StyleNames.Heading1;
            title.Format.Font.Size = 22;
            title.Format.SpaceBefore = Unit.FromPoint(4);
            var subtitle = section.AddParagraph(Safe(data.Subtitle, data.ScenarioTitle));
            subtitle.Format.Font.Size = 11;
            subtitle.Format.SpaceAfter = Unit.FromMillimeter(3);

            Paragraph provenance = section.AddParagraph();
            provenance.AddFormattedText(data.Synthetic ? "Scenario: " : "Context: ", TextFormat.Bold);
            provenance.AddText(Safe(data.ScenarioId, "detached review"));
            provenance.AddText(" · ");
            provenance.AddFormattedText("Generated: ", TextFormat.Bold);
            provenance.AddText(
                data.GeneratedUtc.ToUniversalTime().ToString(
                    "yyyy-MM-dd HH:mm 'UTC'",
                    CultureInfo.InvariantCulture));
            AddPlanRows(section, data.Plans);
            AddPlanAnalysis(section, data.PlanAnalysis);
            if (!string.IsNullOrWhiteSpace(data.VisibilityDisclosure()))
                section.AddParagraph(data.VisibilityDisclosure()).Format.Font.Size = 8;

            section.AddPageBreak();
            AddPqmRows(section, data.VisiblePqmRows());

            if (data.PlanAnalysis != null && data.PlanAnalysis.TargetQuality != null && data.PlanAnalysis.TargetQuality.Any(row => row != null))
                section.AddPageBreak();
            AddTargetQuality(section, data.PlanAnalysis);
            section.AddPageBreak();
            AddPlanImagesWithSource(section, ReviewImageRenderer.VisiblePlanImages(data), data.Synthetic,
                data.Synthetic && Core.Simulation.SyntheticPublicationScenarioFactory.ScenarioIds.Contains(data.ScenarioId, StringComparer.Ordinal));

            section.AddPageBreak();
            AddHeading(section, "Dose-volume overview | Active plan");
            var dvhImage = section.AddImage(ReviewImageRenderer.Dvh(data));
            dvhImage.Width = Unit.FromCentimeter(25.5);
            dvhImage.LockAspectRatio = true;
            AddDvhRows(section, data.VisibleDvhSeries());

            AddControlPointTraces(section, data.PlanAnalysis);
            if (data.IncludeBeamEyeViews) AddBeamViews(section, data.PlanAnalysis, data.Synthetic);
            section.AddPageBreak();
            AddHeading(section, "Review details | Active plan");
            AddPlanCheckRows(section, data.PlanCheckRows);
            AddFieldRows(section, data.FieldRows);
            AddMappingRows(section, data.VisibleStructureMappings());
        }

        private void AddPlanAnalysis(Section section, ReviewPlanAnalysis analysis)
        {
            AddHeading(section, "Delivery and modulation summary");
            if (analysis == null)
            {
                section.AddParagraph("Plan analysis unavailable - no values have been substituted.");
                return;
            }
            Table table = CreateTable(section, 4.2, 4.2, 4.2, 5.2, 5.5);
            AddHeader(table, "Total MU", "MU / Gy per fraction", "PAM [0-1]", "Mean aperture [cm2]", "Small-aperture fraction");
            SetCells(table.AddRow(), Safe(FormatNumber(analysis.TotalMetersetMu), "Unavailable"), Safe(FormatNumber(analysis.MuPerGy), "Unavailable"),
                analysis.Pam.HasValue ? FormatNumber(analysis.Pam) : "Unavailable",
                Safe(FormatNumber(analysis.MeanApertureAreaCm2), "Unavailable"), Safe(FormatNumber(analysis.SmallApertureFraction), "Unavailable"));
            Paragraph note = section.AddParagraph(ReviewImageRenderer.PamTargetSummary(analysis) +
                (analysis.PamWeightingMode == "BeamWeightFactor" ? " | PlanCheck beam weighting. " : " | MU beam weighting. ") +
                "Small aperture: < " + FormatNumber(analysis.SmallApertureThresholdCm2) + " cm2. Descriptive metrics.");
            note.Format.Font.Size = 8;
            note.Format.SpaceAfter = Unit.FromPoint(0);
            if (!analysis.Pam.HasValue)
            {
                var unavailable = section.AddParagraph("PAM unavailable: " + ReviewImageRenderer.PamUnavailableReason(analysis));
                unavailable.Format.Font.Size = 8;
                unavailable.Format.SpaceAfter = Unit.FromPoint(0);
            }
            var normalization = section.AddParagraph("Plan normalization: " +
                (analysis.PlanNormalizationPercent.HasValue ? FormatNumber(analysis.PlanNormalizationPercent) + " %" : "Unavailable") +
                ". TPS setting.");
            normalization.Format.Font.Size = 8;
            normalization.Format.SpaceAfter = Unit.FromPoint(0);
            Table beams = CreateTable(section, 3.0, 3.0, 2.0, 2.0, 3.5, 3.0, 7.0);
            AddHeader(beams, "Beam", "MLC model", "Layers", "MU", "Nominal MU/min", "Geometry", "Availability note");
            foreach (var beam in analysis.Beams ?? new List<ReviewBeamAnalysis>())
                SetCells(beams.AddRow(), beam.BeamId, beam.MlcModel, beam.GeometryStatus=="available" ? beam.MlcLayerCount.ToString(CultureInfo.InvariantCulture) : "Unverified",
                    FormatNumber(beam.MetersetMu), FormatNumber(beam.NominalDoseRateMuPerMin), beam.GeometryStatus,
                    beam.GeometryStatus == "available" ? "Geometry supplied." : "Physical geometry unverified.");
            Paragraph doseRate = section.AddParagraph("Nominal MU/min: plan setting, not measured delivery.");
            doseRate.Format.Font.Size = 8;
            doseRate.Format.SpaceAfter = Unit.FromPoint(0);
        }

        private void AddTargetQuality(Section section, ReviewPlanAnalysis analysis)
        {
            AddHeading(section, "PTV quality | Conformity, gradient and homogeneity");
            var rows = analysis == null ? new List<ReviewTargetQuality>() : analysis.TargetQuality ?? new List<ReviewTargetQuality>();
            if (rows.Count == 0)
            {
                section.AddParagraph("PTV quality unavailable: no evaluated target DVH inputs.");
                return;
            }
            else
            {
                Table table = CreateTable(section, 5.3, 2.4, 3.0, 3.3, 3.0, 4.3);
                AddHeader(table, "PTV", "Plan Rx [Gy]", "Paddick CI", "1 / Paddick CI", "GI V50/V100", "HI (PlanCheck)");
                foreach (var row in rows)
                {
                    SetCells(table.AddRow(), row.StructureId, QualityNumber(row.ReferenceDoseGy), QualityNumber(row.PaddickCi),
                        QualityNumber(row.PlanCheckCi), QualityNumber(row.GradientIndex), QualityNumber(row.HomogeneityIndex));
                    if (!string.IsNullOrWhiteSpace(row.Note)) section.AddParagraph(row.StructureId + ": " + row.Note).Format.Font.Size = 8;
                }
                Table inputs = CreateTable(section, 5.3, 3.0, 3.0, 3.0, 3.0, 2.0, 2.0);
                AddHeader(inputs, "PTV / inputs", "TV [cm3]", "TV at Rx [cm3]", "Body V100 [cm3]", "Body V50 [cm3]", "D2 [Gy]", "D98 [Gy]");
                foreach (var row in rows) SetCells(inputs.AddRow(), row.StructureId, QualityNumber(row.TargetVolumeCm3),
                    QualityNumber(row.CoveredTargetVolumeCm3), QualityNumber(row.BodyV100Cm3), QualityNumber(row.BodyV50Cm3), QualityNumber(row.D2Gy), QualityNumber(row.D98Gy));
            }
            section.AddParagraph("GI: whole EXTERNAL V50 / V100. HI: (D2 - D98) / plan Rx. Rx is the displayed plan prescription.").Format.Font.Size = 8;
        }

        private static string QualityNumber(double? value)
        { return value.HasValue ? value.Value.ToString("0.###", CultureInfo.InvariantCulture) : "Unavailable"; }

        private void AddPlanImages(Section section, IList<ReviewPlanImage> images, bool synthetic)
        { AddPlanImagesWithSource(section, images, synthetic, false); }

        private void AddPlanImagesWithSource(Section section, IList<ReviewPlanImage> images, bool synthetic, bool sharedAnalyticalPhantom)
        {
            string[] kinds = { "transversal", "coronal", "sagittal" };
            string[] labels = { "Transversal", "Coronal (orthogonal)", "Sagittal" };
            for (int index = 0; index < 3; index++)
            {
                if (index > 0) section.AddPageBreak();
                AddHeading(section, (synthetic ? "Synthetic CT | " : "Isocenter CT | ") + labels[index]);
                Table table = CreateTable(section, 17.4, 9.9);
                table.Borders.Visible = false;
                Row panel = table.AddRow();
                var legend = panel.Cells[1];
                var source = (images ?? new List<ReviewPlanImage>()).FirstOrDefault(item => item.Kind == kinds[index]);
                if (source == null || source.SourceStatus != ReviewStatusCodes.Available)
                {
                    panel.Cells[0].AddParagraph("IMAGE UNAVAILABLE");
                    legend.AddParagraph(source == null ? "No planning-image snapshot was supplied." : Safe(source.UnavailableReason, "Image extraction did not provide usable pixels."));
                    continue;
                }
                try
                {
                    var image = panel.Cells[0].AddImage(ReviewImageRenderer.Ct(source));
                    // The synthetic image adds its own 32-DIP banner below the
                    // 720-DIP orientation frame. Keep the total panel height fixed.
                    image.Width = Unit.FromCentimeter(source.Synthetic ? 16.1 * 720.0 / 752.0 : 16.1);
                    image.LockAspectRatio = true;
                    legend.AddParagraph(ReviewImageRenderer.CtSummary(source));
                }
                catch (ArgumentException)
                {
                    panel.Cells[0].AddParagraph("IMAGE UNAVAILABLE - invalid detached pixel geometry.");
                }
                legend.Format.Font.Size = 8.5;
                legend.Format.SpaceAfter = Unit.FromPoint(0);
                string missingOverlays = ReviewImageRenderer.CtUnavailableOverlays(source);
                if (!string.IsNullOrWhiteSpace(missingOverlays)) legend.AddParagraph(missingOverlays).Format.Font.Size = 8;
                var visible = (source.Overlays ?? new List<ReviewImageOverlay>()).Where(item => item != null &&
                    item.SourceStatus == ReviewStatusCodes.Available && item.Paths != null && item.Paths.Count > 0).ToList();
                if (visible.Count > 0)
                {
                    legend.AddParagraph("Contours: thin | Isodoses: thick").Format.Font.Bold = true;
                    foreach (var overlay in visible.Where(item => item.Kind != "isodose"))
                    {
                        var line = legend.AddParagraph(); line.Format.SpaceAfter = Unit.FromPoint(0);
                        var key = line.AddFormattedText("\u2014 ", TextFormat.Bold);
                        try { key.Color = Color.Parse(overlay.ColorHex); } catch (Exception) { key.Color = Colors.Black; }
                        line.AddText(overlay.Label);
                        line.Format.Font.Size = 8.5;
                    }
                    var doseLegend = legend.AddParagraph();
                    doseLegend.Format.SpaceBefore = Unit.FromPoint(3);
                    doseLegend.Format.Font.Size = 8.5;
                    foreach (var overlay in visible.Where(item => item.Kind == "isodose"))
                    {
                        var key = doseLegend.AddFormattedText("\u2014 ", TextFormat.Bold);
                        try { key.Color = Color.Parse(overlay.ColorHex); } catch (Exception) { key.Color = Colors.Black; }
                        doseLegend.AddText(overlay.Label + "   ");
                    }
                }
                legend.AddParagraph(sharedAnalyticalPhantom ?
                "Shared analytical phantom for CT, DVH and BEV; not dose calculated from apertures." : synthetic ?
                "Synthetic CT fixture; not the source volume for the BEV DRRs." :
                "Verify anatomy, dose and contours in the TPS.");
            }
        }

        private void AddBeamViews(Section section, ReviewPlanAnalysis analysis, bool synthetic)
        {
            var beams = analysis == null ? new List<ReviewBeamAnalysis>() : analysis.Beams ?? new List<ReviewBeamAnalysis>();
            if (beams.Count == 0)
            {
                AddHeading(section, "Beam's-eye views");
                section.AddParagraph("BEV unavailable - no detached beam geometry.");
                return;
            }
            var unavailable = new List<string>();
            foreach (var beam in beams.Where(item => item != null))
            {
                var cp = ReviewImageRenderer.StartControlPoint(beam);
                var drr = cp == null ? null : cp.BevImage;
                if (synthetic && cp != null && drr == null)
                    drr = ClearPlan.Core.Simulation.SyntheticDrrFactory.Create(beam, cp);
                var state = ClearPlan.Rendering.BeamEyeViewRenderer.InspectState(beam, cp, drr, synthetic);
                if (!state.ImageAvailable)
                {
                    unavailable.Add(Safe(beam.BeamId, "Beam") + " · " + ReviewImageRenderer.StartAngles(cp) + " DRR unavailable.");
                    continue;
                }
                section.AddPageBreak(); AddHeading(section, "Beam's-eye view | " + Safe(beam.BeamId, "Beam"));
                var introduction = section.AddParagraph("Field start · Exact CP 0 · Not arc-integrated fluence.");
                introduction.Format.Font.Size = 8;
                introduction.Format.KeepWithNext = true;
                introduction.Format.SpaceAfter = Unit.FromMillimeter(2);
                byte[] png = ClearPlan.Rendering.BeamEyeViewRenderer.Render(beam, cp, drr, synthetic);
                var panel = section.AddParagraph();
                panel.Format.Alignment = ParagraphAlignment.Center;
                panel.Format.KeepWithNext = true;
                panel.Format.SpaceAfter = Unit.FromMillimeter(1);
                var image = panel.AddImage("base64:" + Convert.ToBase64String(png));
                // A4 landscape leaves 17.7 cm height: reserve 3.7 cm for heading and provenance.
                image.Width = Unit.FromCentimeter(21); image.LockAspectRatio = true;
                var caption = section.AddParagraph(ReviewImageRenderer.StartAngles(cp));
                caption.Format.Font.Size = 8;
            }
            if (unavailable.Count > 0)
            {
                AddHeading(section, "Beam's-eye views | Unavailable field starts");
                foreach (string message in unavailable) section.AddParagraph(message).Format.Font.Size = 8;
            }
        }

        private void AddControlPointTraces(Section section, ReviewPlanAnalysis analysis)
        {
            var beams = analysis == null ? new List<ReviewBeamAnalysis>() : analysis.Beams ?? new List<ReviewBeamAnalysis>();
            for (int start = 0; start < Math.Max(1, beams.Count); start += 2)
            {
                section.AddPageBreak();
                AddHeading(section, "Control-point trajectories | Active plan");
                section.AddParagraph("Nominal MU/min is a field setting, not a trajectory. Estimated plan trajectory: dashed, PlanCheck-style segment averages; " +
                    "supplied plan values: solid. Not measured delivery. Model excludes acceleration, leaf/jaw motion, ramping and holds; index is not time.");
                if (beams.Count == 0) { section.AddParagraph("Control-point data unavailable."); break; }
                for (int index = start; index < Math.Min(start + 2, beams.Count); index++)
                {
                    var beam = beams[index];
                    Paragraph label = section.AddParagraph(Safe(beam.BeamId, "Beam"));
                    label.Format.Font.Bold = true; label.Format.KeepWithNext = true;
                    label.Format.SpaceBefore = Unit.FromMillimeter(2);
                    var image = section.AddImage(ReviewImageRenderer.ControlPointTraces(beam));
                    image.Width = Unit.FromCentimeter(ContentWidthCentimeter(section)); image.LockAspectRatio = true;
                    var samples = beam.ControlPoints ?? new List<ReviewControlPointSample>();
                    int planned = samples.Count(cp => cp != null && IsFiniteNonnegative(cp.PlannedDoseRateMuPerMin));
                    bool estimated = beam.DoseRateEstimateStatus == "Estimated" && samples.Any(cp => cp != null && IsFiniteNonnegative(cp.EstimatedDoseRateMuPerMin));
                    Paragraph note = section.AddParagraph(estimated
                        ? string.Format(CultureInfo.InvariantCulture, "Estimate profile: {0}; gantry assumption: {1:0.##} deg/s; estimated duration: {2:0.0} s. Configured assumptions require local verification.",
                            Safe(beam.DoseRateEstimateProfile, "unspecified"), beam.DoseRateEstimateMaxGantrySpeedDegreesPerSecond, beam.EstimatedBeamDurationSeconds)
                        : planned > 0 ? "Supplied plan values; no measured delivery timing." : "Rate estimate unavailable; check machine profile and native input completeness.");
                    note.Format.Font.Size = 7.5;
                }
            }
        }

        private static bool IsFiniteNonnegative(double? value)
        {
            return value.HasValue && value.Value >= 0 && !double.IsNaN(value.Value) && !double.IsInfinity(value.Value);
        }

        private void AddPlanRows(
            Section section,
            IList<ReviewReportPlanRow> rows)
        {
            AddHeading(section, "Active plan");
            Table table = CreateTable(
                section,
                4.0,
                3.0,
                2.0,
                2.0,
                1.8,
                3.0);
            AddHeader(
                table,
                "Plan",
                "Created",
                "Dose/Fx [Gy]",
                "Total [Gy]",
                "Fractions",
                "Target");
            foreach (ReviewReportPlanRow row in rows ??
                     new List<ReviewReportPlanRow>())
            {
                Row target = table.AddRow();
                SetCells(
                    target,
                    row.DisplayLabel,
                    FormatUtc(row.CreatedUtc),
                    FormatNumber(row.DosePerFractionGy),
                    FormatNumber(row.TotalDoseGy),
                    FormatNumber(row.FractionCount),
                    row.TargetDisplayLabel);
            }
        }

        private void AddPqmRows(
            Section section,
            IList<ReviewReportPqmRow> rows)
        {
            AddHeading(section, "Plan quality metrics (PQM)");
            if(rows==null || rows.Count==0)
            {
                section.AddParagraph("PQM not evaluated: no visible confirmed objectives. Not a passed review.");
                return;
            }
            Table table = CreateTable(
                section,
                3.0,
                3.0,
                2.2,
                1.4,
                2.0,
                2.0,
                2.0,
                2.0);
            AddHeader(
                table,
                "Template",
                "Resolved structure",
                "Objective",
                "Rule",
                "Goal",
                "Variation",
                "Achieved",
                "Status");
            foreach (ReviewReportPqmRow row in rows ??
                     new List<ReviewReportPqmRow>())
            {
                Row target = table.AddRow();
                SetCells(
                    target,
                    row.TemplateStructure,
                    row.ResolvedStructureId,
                    row.Objective,
                    row.Comparator,
                    FormatValue(row.Goal, row.Unit),
                    FormatValue(row.Variation, row.Unit),
                    FormatValue(row.AchievedValue, row.Unit),
                    StatusWithSeverity(row.Status, row.Severity == "info" ? null : row.Severity) +
                    (string.IsNullOrWhiteSpace(ReviewImageRenderer.PqmUnavailableReason(row)) ? "" : "\n" + ReviewImageRenderer.PqmUnavailableReason(row)));
                ShadeStatus(target.Cells[7], row.Status);
            }
            if (rows.Any(row => row.Status == ReviewStatusCodes.NotEvaluated))
            {
                var note = section.AddParagraph("Not evaluated is not passed. Available measurements use the current plan dose without scaling.");
                note.Format.Font.Size = 8;
            }
        }

        private void AddPlanCheckRows(
            Section section,
            IList<ReviewReportCheckRow> rows)
        {
            AddHeading(section, "PlanCheck");
            Table table = CreateTable(section, 18.5, 3.0, 2.8, 3.0);
            AddHeader(
                table,
                "Message",
                "Status",
                "Check",
                "Category");
            foreach (ReviewReportCheckRow row in rows ??
                     new List<ReviewReportCheckRow>())
            {
                Row target = table.AddRow();
                SetCells(
                    target,
                    row.Message,
                    StatusWithSeverity(row.Status, row.Severity),
                    row.CheckCode,
                    row.Category);
                ShadeStatus(target.Cells[1], row.Status);
            }
        }

        private void AddFieldRows(
            Section section,
            IList<ReviewReportFieldRow> rows)
        {
            AddHeading(section, "Field identifiers and names");
            Table table = CreateTable(
                section,
                1.5,
                1.5,
                2.5,
                2.5,
                4.0,
                4.0,
                2.0,
                2.0);
            AddHeader(
                table,
                "Order",
                "Beam #",
                "Current ID",
                "Expected ID",
                "Current name",
                "Suggested name",
                "ID status",
                "Name status");
            foreach (ReviewReportFieldRow row in rows ??
                     new List<ReviewReportFieldRow>())
            {
                Row target = table.AddRow();
                SetCells(
                    target,
                    row.TreatmentOrder.ToString(
                        CultureInfo.InvariantCulture),
                    row.BeamNumber.ToString(CultureInfo.InvariantCulture),
                    row.CurrentId,
                    row.ExpectedId,
                    row.CurrentName,
                    row.SuggestedName,
                    row.IdStatus,
                    row.NameStatus);
                ShadeStatus(target.Cells[6], row.IdStatus);
                ShadeStatus(target.Cells[7], row.NameStatus);
            }
        }

        private void AddMappingRows(
            Section section,
            IList<ReviewReportStructureMappingRow> rows)
        {
            AddHeading(section, "Structure mapping");
            if(rows==null || rows.Count==0)
            {
                section.AddParagraph("No confirmed objective-to-structure assignments in this report.");
                return;
            }
            Table table = CreateTable(section, 10.0, 10.0, 7.3);
            AddHeader(
                table,
                "Template structure",
                "Resolved structure",
                "Status");
            foreach (ReviewReportStructureMappingRow row in rows ??
                     new List<ReviewReportStructureMappingRow>())
            {
                Row target = table.AddRow();
                SetCells(
                    target,
                    row.TemplateStructure,
                    row.SelectedStructureId,
                    row.Status);
                ShadeStatus(target.Cells[2], row.Status);
            }
        }

        private void AddDvhRows(Section section, IList<ReviewReportDvhSeries> rows)
        {
            AddHeading(section, "DVH statistics | Dose in Gy");
            Table table = CreateTable(section, 8.1, 3.0, 5.4, 5.4, 5.4);
            AddHeader(table, "Structure", "Volume [cm3]", "Dmin / D98", "Dmean / D50", "Dmax / D2");
            var displayed = (rows ?? new List<ReviewReportDvhSeries>()).Where(row => row != null &&
                (row.Selected || row.RequiredForTargetReview))
                .OrderBy(row => row.TargetKind == "PTV" ? 0 : row.RequiredForTargetReview ? 1 : 2)
                .ThenBy(row => row.DisplayName, StringComparer.OrdinalIgnoreCase);
            foreach (var row in displayed)
            {
                var stats = row.Statistics;
                SetCells(table.AddRow(), row.DisplayName, FormatNumber(row.VolumeCc),
                    stats == null ? "-" : DvhPair(stats.MinimumDoseGy, stats.D98Gy, stats.MinimumDoseEstimated),
                    stats == null ? "-" : DvhPair(stats.MeanDoseGy, stats.MedianDoseGy, stats.MeanDoseEstimated),
                    stats == null ? "-" : DvhPair(stats.MaximumDoseGy, stats.D2Gy, stats.MaximumDoseEstimated));
            }
            var note = section.AddParagraph("D98 / D50 / D2: dose at 98% / 50% / 2% volume. Approx.: synthetic estimate; -: unavailable. Selected curves are plotted.");
            note.Format.Font.Size = 8;
        }

        private string DvhPair(double? first, double? second, bool estimated)
        {
            return (estimated ? "approx. " : "") + (first.HasValue ? FormatNumber(first) : "-") + " / " +
                (second.HasValue ? FormatNumber(second) : "-");
        }

        private void AddHeading(Section section, string text)
        {
            section.AddParagraph(text, StyleNames.Heading2);
        }

        private Table CreateTable(
            Section section,
            params double[] columnWidthsCentimeter)
        {
            Table table = section.AddTable();
            table.Borders.Width = 0.25;
            table.Borders.Color = Color.FromRgb(212, 225, 233);
            table.Rows.LeftIndent = 0;
            table.Format.Font.Size = 7.5;
            table.LeftPadding = Unit.FromMillimeter(1.2);
            table.RightPadding = Unit.FromMillimeter(1.2);
            table.TopPadding = Unit.FromMillimeter(0.8);
            table.BottomPadding = Unit.FromMillimeter(0.8);
            // Call-site values are relative width weights; every review table spans the same content box.
            double totalWeight = columnWidthsCentimeter.Sum();
            double availableWidth = ContentWidthCentimeter(section);
            if (totalWeight <= 0 || columnWidthsCentimeter.Any(width => width <= 0))
                throw new ArgumentException("Review table column weights must be positive.");
            foreach (double width in columnWidthsCentimeter)
            {
                table.AddColumn(Unit.FromCentimeter(width / totalWeight * availableWidth));
            }

            return table;
        }

        private static double ContentWidthCentimeter(Section section)
        {
            Unit width, height;
            PageSetup.GetPageSize(section.PageSetup.PageFormat, out width, out height);
            double pageWidth = section.PageSetup.Orientation == Orientation.Landscape ? height.Centimeter : width.Centimeter;
            return pageWidth - section.PageSetup.LeftMargin.Centimeter - section.PageSetup.RightMargin.Centimeter;
        }

        private void AddHeader(Table table, params string[] labels)
        {
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Font.Bold = true;
            row.Shading.Color = Color.FromRgb(224, 237, 243);
            SetCells(row, labels);
        }

        private void SetCells(Row row, params string[] values)
        {
            int columnCount = row.Table == null
                ? row.Cells.Count
                : row.Table.Columns.Count;
            for (int index = 0;
                 index < values.Length && index < columnCount;
                 index++)
            {
                row.Cells[index].AddParagraph(
                    Safe(values[index], string.Empty));
            }
        }

        private void ShadeStatus(Cell cell, string status)
        {
            switch ((status ?? string.Empty).ToLowerInvariant())
            {
                case "pass":
                case "available":
                    cell.Shading.Color = Color.FromRgb(218, 242, 224);
                    break;
                case "variation":
                case "fallback":
                    cell.Shading.Color = Color.FromRgb(255, 239, 184);
                    break;
                case "fail":
                case "unavailable":
                    cell.Shading.Color = Color.FromRgb(255, 214, 214);
                    break;
            }
        }

        private string StatusWithSeverity(string status, string severity)
        {
            if (string.IsNullOrWhiteSpace(severity) ||
                severity.Equals("none", StringComparison.OrdinalIgnoreCase))
            {
                return Safe(status, string.Empty);
            }

            return string.Format(
                CultureInfo.InvariantCulture,
                "{0} ({1})",
                Safe(status, string.Empty),
                severity);
        }

        private string FormatValue(double? value, string unit)
        {
            string formatted = FormatNumber(value);
            if (formatted.Length == 0)
            {
                return formatted;
            }

            return string.IsNullOrWhiteSpace(unit)
                ? formatted
                : formatted + " " + unit;
        }

        private string FormatNumber(double? value)
        {
            return value.HasValue
                ? value.Value.ToString("0.###", CultureInfo.InvariantCulture)
                : string.Empty;
        }

        private string FormatNumber(int? value)
        {
            return value.HasValue
                ? value.Value.ToString(CultureInfo.InvariantCulture)
                : string.Empty;
        }

        private string FormatUtc(DateTimeOffset? value)
        {
            return value.HasValue
                ? value.Value.ToUniversalTime().ToString(
                    "yyyy-MM-dd HH:mm",
                    CultureInfo.InvariantCulture)
                : string.Empty;
        }

        private string Safe(string value, string fallback)
        {
            return string.IsNullOrWhiteSpace(value)
                ? (fallback ?? string.Empty)
                : value;
        }
        

        private void AddHeaderAndFooter(Section section, ReportData data)
        {
            new HeaderAndFooter().Add(section, data);
        }

        private void AddContents(Section section, ReportData data)
        {

            AddPatientInfo(section, data.ReportPatient, data.ReportPlanningItem);
            AddRefs(section, data.ReportStructureSet, data.ReportPlanningItem, data.ReportPQMs, data.ReportRefs, data.ReportDVH);
            AddPCs(section, data.ReportPCs, data.ReportPlanningItem);
            //section.AddPageBreak();
            AddPQMs(section, data.ReportStructureSet, data.ReportPlanningItem, data.ReportPQMs, data.ReportPatient);
            AddDVHstats(section, data.ReportDVHstats);
        }

        private void AddPatientInfo(Section section, ReportPatient patient, ReportPlanningItem reportPlanningItem)
        {
            new PatientInfo().Add(section, patient, reportPlanningItem);
        }

        private void AddPQMs(Section section, ReportStructureSet structureSet, ReportPlanningItem reportPlanningItem, ReportPQMs reportPQMs, ReportPatient reportPatient)
        {
            new PQMsContent().Add(section, structureSet, reportPlanningItem, reportPQMs, reportPatient);
        }

        private void AddRefs(Section section, ReportStructureSet structureSet, ReportPlanningItem reportPlanningItem, ReportPQMs reportPQMs, ReportRefs reportRefs, ReportDVH reportDVH)
        {
            new RefsContent().Add(section, structureSet, reportPlanningItem, reportPQMs, reportRefs, reportDVH);
        }

        private void AddPCs(Section section, ReportPCs reportPCs,ReportPlanningItem reportPlanningItem)
        {
            if (reportPlanningItem.PrintPC_CheckboxState == true)
            { new PCsContent().Add(section, reportPCs); }
        }
        private void AddDVHstats(Section section, ReportDVHstats reportDVHstats)
        {
            new DVHstatsContent().Add(section, reportDVHstats);
        }

    }
}
