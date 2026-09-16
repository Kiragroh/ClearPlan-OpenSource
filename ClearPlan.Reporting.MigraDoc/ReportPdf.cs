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
using ClearPlan.Core.Localization;

namespace ClearPlan.Reporting.MigraDoc
{
    // Fixed presentation labels only. Do not call this on identity cells or free clinical text.
    internal static class ReviewReportLabels
    {
        private static readonly IDictionary<string, string> German = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            { "Active plan", "Aktiver Plan" }, { "Current plan", "Aktueller Plan" },
            { "Clinical goals", "Klinische Ziele" }, { "Plan quality metrics (PQM)", "Planqualitätsmetriken (PQM)" },
            { "Delivery and modulation summary", "Applikations- und Modulationsübersicht" },
            { "PTV quality | Conformity, gradient and homogeneity", "PTV-Qualität | Konformität, Gradient und Homogenität" },
            { "Target quality", "Zielvolumenqualität" }, { "Plan parameters", "Planparameter" },
            { "CT overview", "CT-Übersicht" }, { "Coronal (orthogonal)", "Koronal (orthogonal)" },
            { "Dose-volume overview | Active plan", "Dosis-Volumen-Übersicht | Aktiver Plan" },
            { "Review details | Active plan", "Prüfdetails | Aktiver Plan" },
            { "DVH statistics | Dose in Gy", "DVH-Statistik | Dosis in Gy" },
            { "Field identifiers and names", "Feldkennungen und -namen" }, { "Structure mapping", "Strukturzuordnung" },
            { "Field reviews | Unavailable inputs", "Feldübersichten | Nicht verfügbare Daten" },
            { "Collision / 3D | Sampled geometry review", "Kollision / 3D | Stichprobenartige Geometrieprüfung" },
            { "Collision / 3D | Minimum model distance", "Kollision / 3D | Kleinster Modellabstand" },
            { "Beam's-eye views and control-point trajectories", "Beam's-Eye-Views und Kontrollpunktverläufe" },
            { "Control-point trajectories", "Kontrollpunktverläufe" },
            { "Created", "Erstellt" }, { "Dose/Fx [Gy]", "Dosis/Fx [Gy]" },
            { "Total [Gy]", "Gesamt [Gy]" }, { "Total dose [Gy]", "Gesamtdosis [Gy]" },
            { "Dose / fraction [Gy]", "Dosis / Fraktion [Gy]" }, { "Fractions", "Fraktionen" }, { "Target", "Zielvolumen" },
            { "Total MU", "Gesamt-MU" }, { "MU / Gy per fraction", "MU / Gy je Fraktion" },
            { "Mean aperture [cm2]", "Mittlere Apertur [cm2]" }, { "Mean aperture [cm²]", "Mittlere Apertur [cm²]" },
            { "Small-aperture fraction", "Anteil kleiner Aperturen" }, { "Small aperture fraction", "Anteil kleiner Aperturen" },
            { "Beam", "Feld" }, { "Field", "Feld" }, { "MLC model", "MLC-Modell" }, { "Layers", "Lagen" },
            { "Geometry", "Geometrie" }, { "Availability note", "Verfügbarkeitshinweis" },
            { "Geometry / availability", "Geometrie / Verfügbarkeit" }, { "Availability / scope", "Verfügbarkeit / Geltungsbereich" },
            { "Technique / energy", "Technik / Energie" }, { "Normalization [%]", "Normierung [%]" },
            { "Template", "Vorlage" }, { "Template structure", "Vorlagenstruktur" },
            { "Resolved structure", "Zugeordnete Struktur" }, { "Matched structure", "Zugeordnete Struktur" },
            { "Template structure → matched structure", "Vorlagenstruktur → zugeordnete Struktur" },
            { "Objective", "Kriterium" }, { "Rule", "Regel" }, { "Goal", "Ziel" }, { "Achieved", "Erreicht" },
            { "Result", "Ergebnis" }, { "Message", "Meldung" }, { "Check", "Prüfung" }, { "Category", "Kategorie" },
            { "Order", "Reihenfolge" }, { "Beam #", "Feld #" }, { "Current ID", "Aktuelle ID" }, { "Expected ID", "Erwartete ID" },
            { "Current name", "Aktueller Name" }, { "Suggested name", "Namensvorschlag" },
            { "ID status", "ID-Status" }, { "Name status", "Namensstatus" }, { "Structure", "Struktur" },
            { "Volume [cm3]", "Volumen [cm3]" }, { "Volume [cm³]", "Volumen [cm³]" },
            { "Target / body", "Zielvolumen / Körper" }, { "Target [cm³]", "Zielvolumen [cm³]" },
            { "PTV / inputs", "PTV / Eingangswerte" }, { "Body V100 [cm3]", "Körper V100 [cm3]" },
            { "Body V50 [cm3]", "Körper V50 [cm3]" }, { "Couch", "Tisch" },
            { "Status (whole field)", "Status (gesamtes Feld)" }, { "Body [mm]", "Körper [mm]" }, { "Table [mm]", "Tisch [mm]" }
        };
        public static string Text(string label)
        {
            string german;
            return label != null && German.TryGetValue(label, out german) ? ReviewLanguage.Label(german, label) : label;
        }
        public static string Status(string code)
        {
            switch ((code ?? string.Empty).ToLowerInvariant())
            {
                case "pass": case "passed": return ReviewLanguage.Label("Erfüllt", "Met");
                case "fail": case "failed": return ReviewLanguage.Label("Nicht erfüllt", "Not met");
                case "variation": return ReviewLanguage.Label("Abweichung", "Variation");
                case "warning": return ReviewLanguage.Label("Warnung", "Warning");
                case "error": return ReviewLanguage.Label("Fehler", "Error");
                case "info": return ReviewLanguage.Label("Hinweis", "Note");
                case "available": return ReviewLanguage.Label("Verfügbar", "Available");
                case "unavailable": return ReviewLanguage.Label("Nicht verfügbar", "Unavailable");
                case "not-evaluated": return ReviewLanguage.Label("Nicht bewertet", "Not evaluated");
                case "unmatched": return ReviewLanguage.Label("Nicht zugeordnet", "Unmatched");
                case "uncertain": return ReviewLanguage.Label("Ungewiss", "Uncertain");
                case "hit": return ReviewLanguage.Label("Treffer", "Intersection");
                case "model-hit": return ReviewLanguage.Label("Modelltreffer", "Model intersection");
                default: return ReviewLanguage.Text(code);
            }
        }
    }

    public partial class ReportPdf : IReport
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
            // Source/simulation status repeats in the neutral footer, not as a red page stamp.
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
            using (ClearPlan.Core.Localization.ReviewLanguage.Scope(data.LanguageCode ?? "de"))
                return CreateLocalizedReviewReport(data);
        }

        private Document CreateLocalizedReviewReport(ReviewReportDocument data)
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

            document.Info.Title = ReviewReportDocument.ReportTitle;
            document.Info.Subject = ReviewLanguage.Text(Safe(watermark, data.ModeLabel));

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
            // Reserve the complete identity/version/status footer plus separation
            // from a table that fills the last available body line.
            section.PageSetup.BottomMargin = Unit.FromCentimeter(2.1);
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
                Safe(data.SoftwareVersion, "version unavailable"));
            footer.AddTab();
            footer.AddText(ReviewLanguage.Label("Seite ", "Page "));
            footer.AddPageField();
            footer.AddText(ReviewLanguage.Label(" von ", " of "));
            footer.AddNumPagesField();
            var status = section.Footers.Primary.AddParagraph(ReviewLanguage.Text(Safe(watermark, data.ModeLabel)));
            status.Format.Font.Size = 7;
            status.Format.Font.Color = Color.FromRgb(86, 103, 121);
        }

        private void AddReviewContents(
            Section section,
            ReviewReportDocument data,
            string watermark)
        {
            Paragraph banner = section.AddParagraph(
                ReviewLanguage.Text(Safe(watermark, data.ModeLabel)));
            banner.Format.Font.Size = 8;
            banner.Format.Font.Color = Color.FromRgb(86, 103, 121);
            banner.Format.SpaceAfter = Unit.FromCentimeter(0.3);

            Paragraph title = section.AddParagraph(
                ReviewReportDocument.ReportTitle);
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
            var dvhCurves = data.VisibleDvhSeries().Where(row => (row.Selected || row.RequiredForTargetReview) &&
                row.Points != null && row.Points.Count > 1).OrderByDescending(row => row.RequiredForTargetReview).ToList();
            {
                string dvhPng = ReviewImageRenderer.Dvh(new ReviewReportDocument {
                    DvhSeries = dvhCurves });
                var dvhImage = section.AddImage(dvhPng);
                using (var stream = new System.IO.MemoryStream(Convert.FromBase64String(dvhPng.Substring(7))))
                using (var image = System.Drawing.Image.FromStream(stream))
                    dvhImage.Width = Unit.FromCentimeter(Math.Min(25.5, 15.2 * image.Width / image.Height));
                dvhImage.LockAspectRatio = true;
            }
            AddDvhRows(section, data.VisibleDvhSeries());

            AddFieldReviews(section, data.PlanAnalysis, data.Synthetic, data.IncludeBeamEyeViews);
            if (data.CollisionBeams != null && data.CollisionBeams.Count > 0) AddCollisionSweeps(section, data.CollisionBeams);
            else if (data.CollisionPreviewPng != null && data.CollisionPreviewPng.Length > 0)
            {
                section.AddPageBreak();
                AddHeading(section, "Collision / 3D | Sampled geometry review");
                var collisionImage = section.AddImage("base64:" + Convert.ToBase64String(data.CollisionPreviewPng));
                collisionImage.LockAspectRatio = true;
                using (var stream = new System.IO.MemoryStream(data.CollisionPreviewPng))
                using (var image = System.Drawing.Image.FromStream(stream))
                    collisionImage.Width = Unit.FromCentimeter(Math.Min(22, 12.0 * image.Width / image.Height));
                var caption = section.AddParagraph(ReviewLanguage.Text(data.CollisionPreviewCaption) ?? "Read-only geometry illustration; not clinical clearance.");
                caption.Format.Font.Size = 9;
                caption.Format.SpaceBefore = Unit.FromMillimeter(3);
            }
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
                var unavailable = section.AddParagraph("PAM unavailable: " + ReviewLanguage.Text(ReviewImageRenderer.PamUnavailableReason(analysis)));
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
            if (!(analysis.Beams ?? new List<ReviewBeamAnalysis>()).Any(HasRateTrajectory))
                section.AddParagraph("No valid rate trajectory supplied or estimated; empty trajectory plots omitted.").Format.Font.Size = 8;
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
                    if (!string.IsNullOrWhiteSpace(row.Note)) section.AddParagraph(row.StructureId + ": " + ReviewLanguage.Text(row.Note)).Format.Font.Size = 8;
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
                AddHeading(section, (synthetic ? ReviewLanguage.Label("Synthetisches CT | ", "Synthetic CT | ") : ReviewLanguage.Label("Isozentrums-CT | ", "Isocenter CT | ")) + ReviewReportLabels.Text(labels[index]));
                Table table = CreateTable(section, 17.4, 9.9);
                table.Borders.Visible = false;
                Row panel = table.AddRow();
                var legend = panel.Cells[1];
                var source = (images ?? new List<ReviewPlanImage>()).FirstOrDefault(item => item.Kind == kinds[index]);
                if (source == null || source.SourceStatus != ReviewStatusCodes.Available)
                {
                    panel.Cells[0].AddParagraph("IMAGE UNAVAILABLE");
                    legend.AddParagraph(source == null ? "No planning-image snapshot was supplied." : Safe(ReviewLanguage.Text(source.UnavailableReason), "Image extraction did not provide usable pixels."));
                    continue;
                }
                try
                {
                    var image = panel.Cells[0].AddImage(ReviewImageRenderer.Ct(source));
                    // The shared scale sits below the CT, never over anatomy. Bound total print
                    // height including scale and synthetic marker above the patient footer.
                    image.Width = Unit.FromCentimeter(15.5 * 720.0 / ClearPlan.Rendering.PlanImageRenderer.CanvasHeight(source, true));
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
                    if (visible.Any(item => item.Kind == "isodose"))
                        legend.AddParagraph("Isodose scale below the image: % of plan Rx and Gy. Only lines present in this plane are shown.").Format.SpaceBefore = Unit.FromPoint(5);
                }
                legend.AddParagraph(sharedAnalyticalPhantom ?
                "Shared analytical phantom for CT, DVH and BEV; not dose calculated from apertures." : synthetic ?
                "Synthetic CT fixture; not the source volume for the BEV DRRs." :
                "Verify anatomy, dose and contours in the TPS.");
            }
        }

        private void AddFieldReviews(Section section, ReviewPlanAnalysis analysis, bool synthetic, bool includeBeamEyeViews)
        {
            var beams = analysis == null ? new List<ReviewBeamAnalysis>() : analysis.Beams ?? new List<ReviewBeamAnalysis>();
            if (beams.Count == 0)
            {
                if (includeBeamEyeViews) section.AddParagraph("BEV unavailable - no detached beam geometry.").Format.Font.Size = 8;
                return;
            }
            var unavailable = new List<string>();
            foreach (var beam in beams.Where(item => item != null))
            {
                var cp = ReviewImageRenderer.StartControlPoint(beam);
                var drr = cp == null ? null : cp.BevImage;
                if (includeBeamEyeViews && synthetic && cp != null && drr == null)
                    drr = ClearPlan.Core.Simulation.SyntheticDrrFactory.Create(beam, cp);
                var state = ClearPlan.Rendering.BeamEyeViewRenderer.InspectState(beam, cp, drr, synthetic);
                bool hasTraces = HasRateTrajectory(beam) || (beam.ControlPoints ?? new List<ReviewControlPointSample>())
                    .Any(sample => sample != null && IsFiniteNonnegative(sample.ApertureAreaCm2));
                if (!hasTraces && (!includeBeamEyeViews || !state.ImageAvailable))
                {
                    if (includeBeamEyeViews) unavailable.Add(Safe(beam.BeamId, "Beam") + " · " + ReviewImageRenderer.StartAngles(cp) + " DRR and trajectories unavailable.");
                    continue;
                }
                section.AddPageBreak(); AddHeading(section, ReviewLanguage.Label("Feldübersicht | ", "Field review | ") + Safe(beam.BeamId, "Beam"));
                var introduction = section.AddParagraph(includeBeamEyeViews
                    ? "Beam's-eye view: field start · Exact CP 0 · Not arc-integrated fluence. Control-point trajectories: whole field."
                    : "Control-point trajectories | Whole field");
                introduction.Format.Font.Size = 8;
                introduction.Format.KeepWithNext = true;
                introduction.Format.SpaceAfter = Unit.FromMillimeter(2);
                Table panel = CreateTable(section, 17.0, 10.3);
                panel.Borders.Visible = false;
                panel.TopPadding = panel.BottomPadding = Unit.FromPoint(0);
                Row row = panel.AddRow();
                row.VerticalAlignment = VerticalAlignment.Top;
                var left = row.Cells[0]; var right = row.Cells[1];
                left.Format.SpaceAfter = right.Format.SpaceAfter = Unit.FromPoint(2);
                if (includeBeamEyeViews && state.ImageAvailable)
                {
                    byte[] png = ClearPlan.Rendering.BeamEyeViewRenderer.Render(beam, cp, drr, synthetic);
                    var image = left.AddImage("base64:" + Convert.ToBase64String(png));
                    image.Width = Unit.FromCentimeter(panel.Columns[0].Width.Centimeter - 0.3);
                    image.LockAspectRatio = true;
                }
                else left.AddParagraph(includeBeamEyeViews ? "DRR unavailable - no valid field-start image supplied." : "Beam imagery omitted by report option.");
                if (includeBeamEyeViews) left.AddParagraph(ReviewImageRenderer.StartAngles(cp)).Format.Font.Size = 8;
                var traces = right.AddImage(ReviewImageRenderer.ControlPointTraces(beam));
                traces.Width = Unit.FromCentimeter(panel.Columns[1].Width.Centimeter - 0.3);
                traces.LockAspectRatio = true;
                right.AddParagraph(RateEstimateNote(beam)).Format.Font.Size = 7.5;
                var note = section.AddParagraph("Nominal MU/min is a field setting, not a trajectory. Estimated plan trajectory: dashed; supplied plan values: solid. Not measured delivery. " +
                    "Estimate uses PlanCheck-style segment averages; excludes acceleration, leaf/jaw motion, ramping and holds. CP index is not time.");
                note.Format.Font.Size = 7.5;
                note.Format.SpaceBefore = Unit.FromMillimeter(1);
                note.Format.SpaceAfter = Unit.FromPoint(0);
            }
            if (unavailable.Count > 0)
            {
                AddHeading(section, "Field reviews | Unavailable inputs");
                foreach (string message in unavailable) section.AddParagraph(message).Format.Font.Size = 8;
            }
        }

        private string RateEstimateNote(ReviewBeamAnalysis beam)
        {
            var samples = beam.ControlPoints ?? new List<ReviewControlPointSample>();
            bool estimated = beam.DoseRateEstimateStatus == "Estimated" && samples.Any(cp => cp != null && IsFiniteNonnegative(cp.EstimatedDoseRateMuPerMin));
            return estimated
                ? string.Format(CultureInfo.InvariantCulture, "Estimate profile: {0}; gantry assumption: {1:0.##} deg/s; estimated duration: {2:0.0} s. Configured assumptions require local verification.",
                    Safe(beam.DoseRateEstimateProfile, "unspecified"), beam.DoseRateEstimateMaxGantrySpeedDegreesPerSecond, beam.EstimatedBeamDurationSeconds)
                : samples.Any(cp => cp != null && IsFiniteNonnegative(cp.PlannedDoseRateMuPerMin))
                    ? "Supplied plan values; no measured delivery timing." : "Rate estimate unavailable; check machine profile and native input completeness.";
        }

        private static bool IsFiniteNonnegative(double? value)
        {
            return value.HasValue && value.Value >= 0 && !double.IsNaN(value.Value) && !double.IsInfinity(value.Value);
        }

        private static bool HasRateTrajectory(ReviewBeamAnalysis beam)
        {
            return beam != null && (beam.ControlPoints ?? new List<ReviewControlPointSample>()).Any(cp => cp != null &&
                (IsFiniteNonnegative(cp.PlannedDoseRateMuPerMin) ||
                 beam.DoseRateEstimateStatus == "Estimated" && IsFiniteNonnegative(cp.EstimatedDoseRateMuPerMin)));
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
                    (string.IsNullOrWhiteSpace(ReviewImageRenderer.PqmUnavailableReason(row)) ? "" : "\n" + ReviewLanguage.Text(ReviewImageRenderer.PqmUnavailableReason(row))));
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
                    ReviewLanguage.Text(row.Message),
                    StatusWithSeverity(row.Status, row.Severity),
                    row.CheckCode,
                    ReviewLanguage.Text(row.Category));
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
                    ReviewReportLabels.Status(row.IdStatus),
                    ReviewReportLabels.Status(row.NameStatus));
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
                    ReviewReportLabels.Status(row.Status));
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
            section.AddParagraph(ReviewReportLabels.Text(text), StyleNames.Heading2);
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
            SetCells(row, labels.Select(ReviewReportLabels.Text).ToArray());
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
                return Safe(ReviewReportLabels.Status(status), string.Empty);
            }

            return string.Format(
                CultureInfo.InvariantCulture,
                "{0} ({1})",
                Safe(ReviewReportLabels.Status(status), string.Empty),
                ReviewReportLabels.Status(severity));
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
