using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;
using ClearPlan.Core.Localization;
using ClearPlan.Rendering;
using ClearPlan.Reporting.MigraDoc.Internal;

namespace ClearPlan.Reporting.MigraDoc
{
    /// <summary>
    /// Offline, script-free quicklook over the caller's already privacy-projected document.
    /// Does not acquire TPS data, calculate metrics, project CT volumes or alter the document.
    /// Only bounded, already detached raster/geometry inputs are rendered.
    /// </summary>
    public sealed class HtmlReviewReportRenderer
    {
        private const string Missing = "Unavailable";

        public string Render(ReviewReportDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            using (ClearPlan.Core.Localization.ReviewLanguage.Scope(document.LanguageCode ?? "de"))
                return RenderLocalized(document);
        }

        private string RenderLocalized(ReviewReportDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            var html = new StringBuilder(32768);
            html.Append("<!DOCTYPE html>\n<html lang=\"").Append(ReviewLanguage.Code).Append("\"><head><meta charset=\"utf-8\"><meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">");
            html.Append("<title>Plan Quality Report — ").Append(E(document.PlanDisplayLabel)).Append("</title><style>").Append(Styles).Append("</style></head><body>");
            html.Append("<!-- THESIS: inspect the current detached review without reacquiring plan data. OWN-WORLD: Clinical Blueprint navy headings, teal wayfinding, white evidence tables. STORY: identify the plan, inspect goals and checks, then dose and image evidence. FIRST VIEWPORT: report identity and timestamp above a compact plan table and clinical goals. FORM: precisely scoped report extension of DESIGN.md, code-led. FINISH: regression tests and bounded desktop/mobile review; incumbent DESIGN.md retained. -->");
            html.Append("<div class=\"report\"><header><h1>Plan Quality Report</h1><p class=\"identity\">").Append(E(document.PatientDisplayLabel)).Append(" <span class=\"separator\">/</span> ").Append(E(document.PlanDisplayLabel)).Append("</p>");
            html.Append("<p class=\"snapshot\">").Append(E(ReviewLanguage.Label("Aktueller Datenstand · ", "Current snapshot · "))).Append(document.GeneratedUtc == default(DateTimeOffset) ? Missing : E(document.GeneratedUtc.UtcDateTime.ToString("yyyy-MM-dd HH:mm 'UTC'", CultureInfo.InvariantCulture))).Append("</p>");
            if (document.Synthetic) html.Append("<p class=\"simulation\">SYNTHETIC DEMONSTRATION — NOT FOR CLINICAL USE</p>");
            else if (!string.IsNullOrWhiteSpace(document.ModeLabel)) Paragraph(html, ReviewLanguage.Text(document.ModeLabel), "note");
            html.Append("<p class=\"note\">").Append(E(ReviewLanguage.Label("Erfasste Prüfdaten · Nur lesend · Keine Behandlungsfreigabe.", "Captured review data · Read-only · Not treatment approval."))).Append("</p></header>");
            html.Append("<nav aria-label=\"").Append(E(ReviewLanguage.Label("Reportabschnitte", "Report sections"))).Append("\"><a href=\"#goals\">").Append(E(ReviewReportLabels.Text("Clinical goals"))).Append("</a><a href=\"#checks\">PlanCheck</a><a href=\"#dvh\">DVH</a><a href=\"#parameters\">").Append(E(ReviewReportLabels.Text("Plan parameters"))).Append("</a><a href=\"#ct\">CT</a>");
            if (document.IncludeBeamEyeViews) html.Append("<a href=\"#bev\">BEV / DRR</a>");
            html.Append("</nav>");
            if (!string.IsNullOrWhiteSpace(document.VisibilityDisclosure())) Paragraph(html, document.VisibilityDisclosure(), "note");
            Plans(html, document);
            Goals(html, document);
            Checks(html, document);
            Mappings(html, document);
            Dvh(html, document);
            Parameters(html, document.PlanAnalysis);
            Images(html, document);
            FieldReviews(html, document);
            if (document.CollisionBeams != null && document.CollisionBeams.Count > 0) CollisionSweeps(html, document.CollisionBeams);
            else if (document.CollisionPreviewPng != null && document.CollisionPreviewPng.Length > 0)
            {
                Section(html, "collision", "Collision / 3D | Sampled geometry review");
                html.Append("<img style=\"max-width:100%;max-height:70vh;object-fit:contain\" alt=\"Read-only 3D geometry illustration, not clinical clearance\" src=\"data:image/png;base64,")
                    .Append(Convert.ToBase64String(document.CollisionPreviewPng)).Append("\"/>");
                Paragraph(html, ReviewLanguage.Text(document.CollisionPreviewCaption), "note");
                html.Append("</section>");
            }
            html.Append("<footer>ClearPlan · ").Append(E(document.SoftwareVersion)).Append(" · ").Append(E(ReviewLanguage.Label("Nicht verfügbar bedeutet nicht bestanden.", "Unavailable is not passed."))).Append("</footer></div></body></html>");
            return html.ToString();
        }

        public void Export(string path, ReviewReportDocument document)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An HTML output path is required.", "path");
            File.WriteAllText(path, Render(document), new UTF8Encoding(false));
        }

        /// <summary>Creates a new adjacent HTML report; never opens or overwrites the PDF or an existing HTML file.</summary>
        public string ExportCompanion(string pdfPath, ReviewReportDocument document)
        {
            if (string.IsNullOrWhiteSpace(pdfPath)) throw new ArgumentException("A PDF path is required.", "pdfPath");
            string original = Path.GetFullPath(pdfPath);
            if (!string.Equals(Path.GetExtension(original), ".pdf", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Companion reports require a .pdf path.", "pdfPath");
            string html = Render(document);
            string preferred = Path.ChangeExtension(original, ".html");
            for (int attempt = 0; attempt < 5; attempt++)
            {
                string candidate = attempt == 0 ? preferred : Path.Combine(Path.GetDirectoryName(preferred),
                    Path.GetFileNameWithoutExtension(preferred) + "-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N") + ".html");
                FileStream stream;
                try { stream = new FileStream(candidate, FileMode.CreateNew, FileAccess.Write, FileShare.None); }
                catch (IOException) { if (File.Exists(candidate)) continue; throw; }
                using (stream)
                using (var writer = new StreamWriter(stream, new UTF8Encoding(false))) writer.Write(html);
                return candidate;
            }
            throw new IOException("No unused HTML companion filename could be created.");
        }

        private static void CollisionSweeps(StringBuilder html, IList<CollisionReportBeam> beams)
        {
            Section(html, "collision", "Collision / 3D | Minimum model distance");
            var minimum = CollisionReportBeam.MinimumBeam(beams);
            Paragraph(html, minimum == null ? "Minimum model distance unavailable." : minimum.MinimumCaption, "note");
            if (minimum != null && minimum.OverviewPng != null && minimum.OverviewPng.Length > 0)
                html.Append("<img style=\"width:100%;max-width:1050px;object-fit:contain\" alt=\"Minimum model distance position\" src=\"data:image/png;base64,")
                    .Append(Convert.ToBase64String(minimum.OverviewPng)).Append("\"/>");
            if (minimum != null && minimum.OrientationPng != null && minimum.OrientationPng.Length > 0)
                html.Append("<div style=\"margin-top:8px\"><img style=\"width:100%;max-width:600px;object-fit:contain\" alt=\"Gantry and couch orientation at minimum model distance\" src=\"data:image/png;base64,")
                    .Append(Convert.ToBase64String(minimum.OrientationPng)).Append("\"/></div>");
            else Paragraph(html, "Orientation strip unavailable; no substitute pose shown.", "note");
            if (minimum != null) Paragraph(html, minimum.Summary, "note");
            Table(html, "Field", "Gantry", "Couch", "Status (whole field)", "Body [mm]", "Table [mm]", "Min. [mm]");
            foreach (var beam in beams)
            {
                var point = beam.MinimumRow;
                html.Append("<tr>");
                Cell(html, beam.BeamId); Cell(html, point == null ? "—" : point.GantryDegrees.ToString("0.#", CultureInfo.InvariantCulture) + "°");
                Cell(html, point == null ? "—" : point.CouchDegrees.ToString("0.#", CultureInfo.InvariantCulture) + "°");
                CollisionCell(html, CollisionLabel(beam.Status), beam.Status);
                Cell(html, point == null ? "—" : point.BodyDistance);
                Cell(html, point == null ? "—" : point.TableDistance);
                Cell(html, point == null ? "—" : point.MinimumDistanceMm.Value.ToString("0.0", CultureInfo.InvariantCulture));
                html.Append("</tr>");
            }
            EndTable(html, beams.Count, 7);
            Paragraph(html, "10° states plus endpoints. One minimum position per field; status covers all its samples. Body/table: conservative lower bound … sampled upper bound. Min.: smallest finite lower bound (open bore: radial gap). Green / Full clear*: captured body and table clear. Yellow / Still clear*: only part of the geometry assessed clear; see body/table results. Uncertain / Hit remain distinct. Not continuous-path or clinical clearance. Outside-CT anatomy and actual setup remain unverified. Orientation strip: nominal Gantry/Couch angles at the minimum pose. Gray C-arm reference, when present: schematic, not tested geometry.", "note");
            html.Append("</section>");
        }
        private static string CollisionLabel(string status) { return status == "body-clear" || status == "partial-clear" ? "Still clear*" : status == "model-clear" || status == "clear" || status == "pass" ? "Full clear*" : status == "model-hit" ? "Model hit" : status == "hit" ? "Hit" : "Uncertain"; }
        private static void CollisionCell(StringBuilder html, string value, string status)
        {
            html.Append("<td style=\"background:").Append(status == "pass" || status == "clear" || status == "model-clear" ? "#daf2e0" : status == "hit" || status == "model-hit" ? "#fde0dc" : "#fff1cc")
                .Append(";color:#202b38\">").Append(E(value)).Append("</td>");
        }

        private static void Plans(StringBuilder html, ReviewReportDocument document)
        {
            Section(html, "plan", "Current plan");
            var plans = Rows(document.Plans).Where(p => string.Equals(p.PlanKey, document.ActivePlanKey, StringComparison.Ordinal)).ToList();
            Table(html, "Plan", "Target", "Fractions", "Dose / fraction [Gy]", "Total dose [Gy]");
            foreach (var plan in plans) Row(html, plan.DisplayLabel, plan.TargetDisplayLabel, plan.FractionCount.HasValue && plan.FractionCount > 0 ? plan.FractionCount.Value.ToString(CultureInfo.InvariantCulture) : Missing, N(plan.DosePerFractionGy), N(plan.TotalDoseGy));
            EndTable(html, plans.Count, 5); html.Append("</section>");
        }

        private static void Goals(StringBuilder html, ReviewReportDocument document)
        {
            Section(html, "goals", "Clinical goals");
            var rows = document.VisiblePqmRows();
            Table(html, "Template structure → matched structure", "Objective", "Goal", "Variation", "Achieved", "Result");
            foreach (var row in rows)
            {
                html.Append("<tr>");
                Cell(html, T(row.TemplateStructure) + " → " + T(row.ResolvedStructureId));
                Cell(html, row.Objective);
                Cell(html, T(row.Comparator) + " " + WithUnit(row.Goal, row.Unit));
                Cell(html, WithUnit(row.Variation, row.Unit)); Cell(html, WithUnit(row.AchievedValue, row.Unit));
                StatusCell(html, row.Status, ReviewImageRenderer.PqmUnavailableReason(row)); html.Append("</tr>");
            }
            EndTable(html, rows.Count, 6); html.Append("</section>");
        }

        private static void Checks(StringBuilder html, ReviewReportDocument document)
        {
            Section(html, "checks", "PlanCheck");
            var rows = Rows(document.PlanCheckRows).ToList();
            Table(html, "Message", "Status", "Check", "Category");
            foreach (var row in rows)
            {
                html.Append("<tr>"); Cell(html, ReviewLanguage.Text(row.Message)); StatusCell(html, row.Status);
                Cell(html, row.CheckCode); Cell(html, ReviewLanguage.Text(row.Category)); html.Append("</tr>");
            }
            EndTable(html, rows.Count, 4); html.Append("</section>");
        }

        private static void Mappings(StringBuilder html, ReviewReportDocument document)
        {
            Section(html, "mapping", "Structure mapping");
            var rows = document.VisibleStructureMappings();
            Table(html, "Template structure", "Matched structure", "Result");
            foreach (var row in rows)
            {
                html.Append("<tr>"); Cell(html, row.TemplateStructure); Cell(html, row.SelectedStructureId);
                StatusCell(html, row.Status); html.Append("</tr>");
            }
            EndTable(html, rows.Count, 3); html.Append("</section>");
        }

        private static void Dvh(StringBuilder html, ReviewReportDocument document)
        {
            Section(html, "dvh", "DVH");
            var rows = document.VisibleDvhSeries().Where(row => row.Selected || row.RequiredForTargetReview).ToList();
            var curves = rows.Where(ValidCurve).OrderByDescending(row => row.RequiredForTargetReview).Select(row => new ReviewReportDvhSeries {
                DisplayName = T(row.DisplayName), ColorHex = SafeColor(row.ColorHex), LineStyle = row.LineStyle,
                Selected = true, Points = row.Points
            }).ToList();
            if (curves.Count > 0)
                Image(html, ReviewImageRenderer.Dvh(new ReviewReportDocument { DvhSeries = curves }), "DVH · Dose [Gy] / volume [%]", "plot");
            if (curves.Count == 0) Paragraph(html, "Unavailable — no valid selected or required-target DVH curve in this snapshot.", "empty");
            Paragraph(html, "D98 / D50 / D2: dose at 98% / 50% / 2% volume. ~: stored estimate.", "note");
            if (rows.Count > curves.Count) Paragraph(html, "Some curves are unavailable; available statistics remain below.", "note");
            Table(html, "Structure", "Volume [cm³]", "Dmin [Gy]", "D98 [Gy]", "Dmean [Gy]", "D50 [Gy]", "Dmax [Gy]", "D2 [Gy]");
            foreach (var row in rows)
            {
                var stats = row.Statistics ?? new ReviewDvhStatistics();
                Row(html, T(row.DisplayName ?? row.StructureId), N(row.VolumeCc), Estimate(stats.MinimumDoseGy, stats.MinimumDoseEstimated), N(stats.D98Gy), Estimate(stats.MeanDoseGy, stats.MeanDoseEstimated), N(stats.MedianDoseGy), Estimate(stats.MaximumDoseGy, stats.MaximumDoseEstimated), N(stats.D2Gy));
            }
            EndTable(html, rows.Count, 8); html.Append("</section>");
        }

        private static void Parameters(StringBuilder html, ReviewPlanAnalysis analysis)
        {
            Section(html, "parameters", "Plan parameters");
            var plan = analysis ?? new ReviewPlanAnalysis();
            Table(html, "Total MU", "MU / Gy", "PAM", "Mean aperture [cm²]", "Small aperture fraction", "Normalization [%]");
            Row(html, N(plan.TotalMetersetMu), N(plan.MuPerGy), N(plan.Pam), N(plan.MeanApertureAreaCm2), N(plan.SmallApertureFraction), N(plan.PlanNormalizationPercent));
            EndTable(html, 1, 6);
            Paragraph(html, ReviewImageRenderer.PamTargetSummary(plan) + " · " + T(plan.PamWeightingMode) +
                (N(plan.Pam) == Missing ? " · " + ReviewLanguage.Text(ReviewImageRenderer.PamUnavailableReason(plan)) : ""), "note");
            Paragraph(html, "Small aperture: < " + (analysis == null ? Missing : N(plan.SmallApertureThresholdCm2)) + " cm².", "note");
            var beams = Rows(plan.Beams).ToList();
            Table(html, "Beam", "Technique / energy", "MU", "PAM", "Mean aperture [cm²]", "Nominal MU/min", "Geometry / availability");
            foreach (var beam in beams) Row(html, beam.BeamId, Join(beam.Technique, beam.EnergyDisplay), N(beam.MetersetMu), N(beam.Pam), N(beam.MeanApertureAreaCm2), N(beam.NominalDoseRateMuPerMin), Join(ReviewLanguage.Text(beam.GeometryStatus), ReviewLanguage.Text(beam.GeometryReason)));
            EndTable(html, beams.Count, 7);
            Paragraph(html, "Nominal MU/min: plan setting, not measured delivery. PAM / aperture metrics: descriptive.", "note");
            html.Append("<h3>").Append(E(ReviewReportLabels.Text("Target quality"))).Append("</h3>");
            var targets = Rows(plan.TargetQuality).ToList();
            Table(html, "Target / body", "Plan Rx [Gy]", "Target [cm³]", "Paddick CI", "1 / Paddick CI", "GI", "HI", "Availability / scope");
            foreach (var target in targets) Row(html, Join(target.StructureId, target.BodyStructureId), N(target.ReferenceDoseGy), N(target.TargetVolumeCm3), N(target.PaddickCi), N(target.PlanCheckCi), N(target.GradientIndex), N(target.HomogeneityIndex), ReviewLanguage.Text(target.AvailabilityScope));
            EndTable(html, targets.Count, 8);
            Paragraph(html, "GI: whole EXTERNAL V50 / V100. HI: (D2 − D98) / plan Rx. Rx is the displayed plan prescription.", "note");
            foreach (var warning in Rows(plan.Warnings)) Paragraph(html, ReviewLanguage.Text(warning), "note");
            html.Append("</section>");
        }

        private static void Images(StringBuilder html, ReviewReportDocument document)
        {
            Section(html, "ct", "CT overview");
            Paragraph(html, "Cached orthogonal views. Contours: thin; isodoses: thick.", "note");
            var images = ReviewImageRenderer.VisiblePlanImages(document);
            html.Append("<div class=\"ct-images\">");
            foreach (string kind in new[] { "transversal", "coronal", "sagittal" })
            {
                var img = images.FirstOrDefault(item => string.Equals(item.Kind, kind, StringComparison.OrdinalIgnoreCase));
                html.Append("<div class=\"ct-image\"><h3>").Append(E(img == null ? CultureInfo.InvariantCulture.TextInfo.ToTitleCase(kind) : ReviewLanguage.Text(img.Title))).Append("</h3>");
                if (ValidCt(img, document.Synthetic))
                {
                    Image(html, ReviewImageRenderer.Ct(img), T(ReviewLanguage.Text(img.Title)), "ct-raster");
                    CtLegend(html, img);
                    Paragraph(html, ReviewImageRenderer.CtSummary(img), "note");
                    string missingOverlays = ReviewImageRenderer.CtUnavailableOverlays(img);
                    if (!string.IsNullOrWhiteSpace(missingOverlays)) Paragraph(html, missingOverlays, "note");
                }
                else Paragraph(html, "Unavailable — " + (img == null ? "no cached view in this snapshot." : T(ReviewLanguage.Text(img.UnavailableReason), "invalid or unavailable cached image.")), "empty");
                html.Append("</div>");
            }
            html.Append("</div></section>");
        }

        private static void FieldReviews(StringBuilder html, ReviewReportDocument document)
        {
            Section(html, "bev", document.IncludeBeamEyeViews ? "Beam's-eye views and control-point trajectories" : "Control-point trajectories");
            if (document.Synthetic && document.IncludeBeamEyeViews) Paragraph(html,
                Core.Simulation.SyntheticPublicationScenarioFactory.ScenarioIds.Contains(document.ScenarioId, StringComparer.Ordinal) ?
                "Shared analytical phantom for CT, DVH and BEV; not dose calculated from apertures." :
                "Synthetic orthogonal slices are separate fixtures, not the source volume for the BEV DRRs.", "note");
            var beams = document.PlanAnalysis == null ? new List<ReviewBeamAnalysis>() : Rows(document.PlanAnalysis.Beams).ToList();
            foreach (var beam in beams)
            {
                var cp = ReviewImageRenderer.StartControlPoint(beam);
                html.Append("<article class=\"field-review\"><h3>").Append(E(ReviewLanguage.Label("Feldübersicht | ", "Field review | "))).Append(E(beam.BeamId)).Append("</h3>");
                if (document.IncludeBeamEyeViews) Paragraph(html, "Field start · Exact CP 0 · Not arc-integrated fluence. Control-point trajectories: whole field.", "note");
                html.Append("<div class=\"field-layout\"><div class=\"field-bev\">");
                if (document.IncludeBeamEyeViews && cp != null && ValidBev(beam, cp, document.Synthetic))
                {
                    Image(html, "base64:" + Convert.ToBase64String(BeamEyeViewRenderer.Render(beam, cp, cp.BevImage, document.Synthetic, 1500, 1000)), T(beam.BeamId) + " · CP " + cp.Index, "bev-raster");
                    Paragraph(html, ReviewImageRenderer.StartAngles(cp), "note");
                }
                else Paragraph(html, document.IncludeBeamEyeViews ? T(beam.BeamId) + " · Field start · CP 0 — DRR unavailable." : "Beam imagery omitted by report option.", "note");
                html.Append("</div><div class=\"field-traces\">");
                Image(html, ReviewImageRenderer.ControlPointTraces(beam), "Dose rate above effective aperture area, by control-point index", "parameter-traces");
                Paragraph(html, beam.DoseRateEstimateStatus == "Estimated"
                    ? "Estimate profile: " + T(beam.DoseRateEstimateProfile) + "; gantry assumption: " + N(beam.DoseRateEstimateMaxGantrySpeedDegreesPerSecond) + " deg/s; estimated duration: " + N(beam.EstimatedBeamDurationSeconds) + " s. Configured assumptions require local verification."
                    : "Rate estimate unavailable; check machine profile and native input completeness.", "note");
                if (!string.IsNullOrWhiteSpace(beam.DoseRateEstimateReason))
                    html.Append("<details><summary>Model assumptions / availability</summary><p class=\"note rate-assumptions\" tabindex=\"0\">").Append(E(ReviewLanguage.Text(beam.DoseRateEstimateReason))).Append("</p></details>");
                html.Append("</div></div>");
                Paragraph(html, "Estimated plan trajectory: dashed, PlanCheck-style segment averages. Supplied plan values: solid. Not measured delivery. Model excludes acceleration, leaf/jaw motion, ramping and holds; index is not time.", "note");
                html.Append("</article>");
            }
            if (beams.Count == 0) Paragraph(html, "Unavailable — no detached beam analysis in this snapshot.", "empty");
            html.Append("</section>");
        }

        private static void CtLegend(StringBuilder html, ReviewPlanImage image)
        {
            // The caller supplies the same visibility-filtered image used for the raster.
            var visible = Rows(image.Overlays).Where(overlay => overlay.SourceStatus == ReviewStatusCodes.Available &&
                (overlay.Kind == "structure" || overlay.Kind == "isodose") && Rows(overlay.Paths).Any(path =>
                    path.Points != null && path.Points.Count >= 2 && path.Points.All(point => point != null && Finite(point.X) && Finite(point.Y)))).ToList();
            if (visible.Count == 0) return;
            html.Append("<div class=\"ct-legend\" role=\"group\" aria-label=\"Visible CT overlays\">");
            foreach (string kind in new[] { "structure", "isodose" })
            {
                var group = visible.Where(overlay => overlay.Kind == kind).ToList();
                if (group.Count == 0) continue;
                html.Append("<div class=\"ct-legend-group\"><span class=\"ct-legend-heading\">").Append(kind == "structure" ? "Contours" : "Isodoses").Append("</span><ul>");
                foreach (var overlay in group)
                    html.Append("<li><span aria-hidden=\"true\" class=\"ct-legend-key ").Append(kind)
                        .Append("\" style=\"border-top-color:").Append(SafeColor(overlay.ColorHex)).Append("\"></span>")
                        .Append(E(overlay.Label)).Append("</li>");
                html.Append("</ul></div>");
            }
            html.Append("</div>");
        }

        private static bool ValidCt(ReviewPlanImage img, bool synthetic)
        {
            if (img == null || img.Synthetic != synthetic || !string.Equals(img.SourceStatus, "available", StringComparison.OrdinalIgnoreCase) ||
                img.WidthPixels < 2 || img.HeightPixels < 2 || img.WidthPixels > 2048 || img.HeightPixels > 2048 ||
                img.GrayscalePixels == null || img.GrayscalePixels.LongLength != (long)img.WidthPixels * img.HeightPixels ||
                !Positive(img.PixelSpacingXMillimeters) || !Positive(img.PixelSpacingYMillimeters)) return false;
            long points = 0;
            if (img.Overlays != null && img.Overlays.Count > 1000) return false;
            foreach (var overlay in Rows(img.Overlays))
                foreach (var path in Rows(overlay.Paths))
                {
                    points += path.Points == null ? 0 : path.Points.Count;
                    if (points > 200000) return false;
                    if (Rows(path.Points).Any(point => !Finite(point.X) || !Finite(point.Y) || Math.Abs(point.X) > 100000 || Math.Abs(point.Y) > 100000)) return false;
                }
            return true;
        }

        private static bool ValidBev(ReviewBeamAnalysis beam, ReviewControlPointSample cp, bool synthetic)
        {
            if (cp.BevImage == null || cp.BevImage.WidthPixels > 2048 || cp.BevImage.HeightPixels > 2048 ||
                (cp.Aperture != null && cp.Aperture.EffectiveOpenings != null && cp.Aperture.EffectiveOpenings.Count > 10000)) return false;
            return BeamEyeViewRenderer.InspectState(beam, cp, cp.BevImage, synthetic).ImageAvailable;
        }

        private static bool ValidCurve(ReviewReportDvhSeries row)
        {
            if (row.Points == null || row.Points.Count < 2 || row.Points.Count > 10000 ||
                (!string.IsNullOrWhiteSpace(row.DoseUnit) && row.DoseUnit != "Gy") ||
                (!string.IsNullOrWhiteSpace(row.VolumeUnit) && row.VolumeUnit != "%")) return false;
            ReviewReportDvhPoint previous = null;
            foreach (var point in row.Points)
            {
                if (point == null || !Finite(point.DoseGy) || point.DoseGy < 0 || point.DoseGy > 1000000 || !Finite(point.VolumePercent) || point.VolumePercent < 0 || point.VolumePercent > 100 ||
                    (previous != null && (point.DoseGy < previous.DoseGy || point.VolumePercent > previous.VolumePercent))) return false;
                previous = point;
            }
            return true;
        }

        private static IEnumerable<TItem> Rows<TItem>(IEnumerable<TItem> rows) where TItem : class { return (rows ?? Enumerable.Empty<TItem>()).Where(row => row != null); }
        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static bool Positive(double value) { return Finite(value) && value > 0; }
        private static string N(double? value) { return value.HasValue && Finite(value.Value) && value >= 0 ? value.Value.ToString("0.###", CultureInfo.InvariantCulture) : Missing; }
        private static string Estimate(double? value, bool estimated) { return N(value) == Missing ? Missing : (estimated ? "~" : "") + N(value); }
        private static string WithUnit(double? value, string unit) { return N(value) == Missing ? Missing : N(value) + (string.IsNullOrWhiteSpace(unit) ? "" : " " + unit); }
        private static string T(string value, string fallback = Missing) { return string.IsNullOrWhiteSpace(value) ? fallback : value; }
        private static string E(string value) { return WebUtility.HtmlEncode(T(value)); }
        private static string Join(params string[] values) { return T(string.Join(" · ", values.Where(value => !string.IsNullOrWhiteSpace(value)))); }
        private static string SafeColor(string value) { return value != null && value.Length == 7 && value[0] == '#' && value.Skip(1).All(Uri.IsHexDigit) ? value : "#0F766E"; }
        private static void Section(StringBuilder html, string id, string title) { html.Append("<section id=\"").Append(id).Append("\"><h2>").Append(E(ReviewReportLabels.Text(title))).Append("</h2>"); }
        private static void Paragraph(StringBuilder html, string value, string css) { html.Append("<p class=\"").Append(css).Append("\">").Append(E(value)).Append("</p>"); }
        private static void Cell(StringBuilder html, string value) { html.Append("<td>").Append(E(value)).Append("</td>"); }
        private static void Row(StringBuilder html, params string[] values) { html.Append("<tr>"); foreach (string value in values) Cell(html, value); html.Append("</tr>"); }
        private static void StatusCell(StringBuilder html, string value, string reason = null)
        {
            string status = (value ?? "").ToLowerInvariant();
            string css = status == "pass" || status == "passed" ? "pass" : status == "fail" || status == "failed" ? "fail" : status == "variation" || status == "warning" ? "variation" : "info";
            html.Append("<td><span class=\"status ").Append(css).Append("\">").Append(E(ReviewReportLabels.Status(value))).Append("</span>");
            if (!string.IsNullOrWhiteSpace(reason)) html.Append("<br>").Append(E(ReviewLanguage.Text(reason)));
            html.Append("</td>");
        }
        private static void Table(StringBuilder html, params string[] headings) { html.Append("<div class=\"table-scroll\" tabindex=\"0\"><table><thead><tr>"); foreach (string heading in headings) html.Append("<th scope=\"col\">").Append(E(ReviewReportLabels.Text(heading))).Append("</th>"); html.Append("</tr></thead><tbody>"); }
        private static void EndTable(StringBuilder html, int count, int columns) { if (count == 0) html.Append("<tr><td colspan=\"").Append(columns).Append("\">").Append(E(ReviewLanguage.Label("Nicht verfügbar — keine Einträge in diesem Datenstand.", "Unavailable — no entries in this snapshot."))).Append("</td></tr>"); html.Append("</tbody></table></div>"); }
        private static void Image(StringBuilder html, string encoded, string label, string css)
        {
            // Only the internal PNG encoders supply image bytes; no caller-controlled URL or MIME type is accepted.
            if (!encoded.StartsWith("base64:", StringComparison.Ordinal)) throw new InvalidOperationException("Expected an embedded PNG.");
            html.Append("<img loading=\"lazy\" class=\"").Append(css).Append("\" src=\"data:image/png;base64,").Append(encoded.Substring(7)).Append("\" alt=\"").Append(E(label)).Append("\">");
        }

        private const string Styles = @"
.field-review{margin-top:24px}.field-layout{display:table;table-layout:fixed;width:100%}.field-bev,.field-traces{display:table-cell;vertical-align:top}.field-bev{width:62.3%;padding-right:12px}.field-traces{width:37.7%}.parameter-traces{width:100%}.field-review h3{margin:0 0 8px}.field-review .note{font-size:12px;margin:6px 0}@media(max-width:800px){.field-layout,.field-bev,.field-traces{display:block;width:100%;padding:0}.field-traces{max-width:620px;margin-top:16px}}@media print{.field-review{page-break-before:always;page-break-inside:avoid;margin-top:0}.field-layout{display:table}.field-bev,.field-traces{display:table-cell}.field-bev{width:62.3%;padding-right:3mm}.field-traces{width:37.7%;margin-top:0}.field-review .note{font-size:8pt}.field-review details{display:none}.field-review h3{font-size:14pt}}
.ct-legend{margin:10px 0;font-size:12px;line-height:1.4;color:#202B38}.ct-legend-group{margin:0 0 7px}.ct-legend-heading{display:block;font-size:11px;font-weight:600;color:#526474;margin-bottom:3px}.ct-legend ul{list-style:none;padding:0;margin:0}.ct-legend li{display:inline-block;max-width:100%;margin:0 12px 4px 0;overflow-wrap:break-word;word-wrap:break-word}.ct-legend-key{display:inline-block;vertical-align:middle;width:18px;margin-right:5px;border-top:1px solid}.ct-legend-key.isodose{border-top-width:3px}
.rate-assumptions{max-height:6em;overflow-y:auto;padding-right:8px}.rate-assumptions:focus{outline:2px solid #0F766E;outline-offset:2px}@media print{.rate-assumptions{max-height:none;overflow:visible}}
*{box-sizing:border-box}html{background:#F3F5F7;color:#202B38}body{margin:0;font-family:'Segoe UI',Arial,sans-serif;font-size:14px;line-height:1.5}::selection{background:#D9ECE9;color:#16324A}.report{max-width:1440px;margin:0 auto;padding:32px;background:#FFFFFF}header{padding-bottom:20px;border-bottom:1px solid #D6DDE4}h1{margin:0 0 12px;font-size:30px;line-height:1.15;color:#16324A;font-weight:600}h2{margin:0 0 12px;font-size:22px;line-height:1.25;color:#16324A;font-weight:600}h3{margin:22px 0 10px;font-size:17px;font-weight:600;color:#16324A}.identity{font-size:20px;margin:0 0 8px;overflow-wrap:break-word;word-wrap:break-word}.separator{color:#0F766E;padding:0 6px}.snapshot{margin:0;color:#526474}.simulation{display:inline-block;padding:6px 10px;background:#FFF5DD;color:#704B12;border:1px solid #D99B36;font-weight:600}.note{max-width:100ch;margin:10px 0;color:#526474}.empty{padding:12px;background:#EEF1F4;color:#526474}.note,.empty{overflow-wrap:break-word;word-wrap:break-word}nav{padding:14px 0;border-bottom:1px solid #D6DDE4}nav a{display:inline-block;margin-right:22px;padding:4px 0;color:#0F766E;text-decoration:underline}a:hover{color:#16324A}a:focus,.table-scroll:focus{outline:2px solid #0F766E;outline-offset:2px}section{margin-top:32px;page-break-inside:auto}.table-scroll{width:100%;overflow-x:auto;margin:12px 0 16px;border:1px solid #D6DDE4}table{width:100%;min-width:760px;border-collapse:collapse;font-size:13px;font-variant-numeric:tabular-nums}th{text-align:left;vertical-align:top;font-weight:600;background:#EEF1F4;color:#16324A}th,td{padding:9px 11px;border-bottom:1px solid #D6DDE4;word-wrap:break-word;overflow-wrap:break-word;max-width:360px}td{vertical-align:top}tbody tr:nth-child(even){background:#F7F9FA}tbody tr:last-child td{border-bottom:0}.status{display:inline-block;padding:2px 7px;border-radius:4px;font-size:12px}.pass{background:#EAF6EE;color:#116632}.fail{background:#FDEDEC;color:#B42318}.variation{background:#FFF5DD;color:#704B12}.info{background:#EAF2F7;color:#38546B}img{max-width:100%;height:auto;display:block}.plot{width:100%;min-width:0}.ct-images{font-size:0}.ct-image{display:inline-block;vertical-align:top;width:33.333%;padding:0 10px 0 0;font-size:14px}.ct-raster{width:100%}.bev-image{max-width:1200px;page-break-inside:avoid}.bev-raster{width:100%}footer{margin-top:36px;padding-top:16px;border-top:1px solid #D6DDE4;color:#526474;font-size:12px}@media(max-width:800px){.report{padding:20px 16px}h1{font-size:26px}.identity{font-size:18px}.ct-image{width:100%;max-width:640px;padding:0}nav a{margin-right:14px}.plot{min-width:680px}#dvh{overflow-x:auto}}@media print{@page{size:A4 landscape;margin:12mm}html,body{background:#FFFFFF}.report{max-width:none;padding:0;font-size:11px}nav{display:none}section{margin-top:22px}h2,h3{page-break-after:avoid}.table-scroll{overflow:visible}table{min-width:0;font-size:10px}th,td{padding:5px 6px}thead{display:table-header-group}tr{page-break-inside:avoid}.plot{width:100%;min-width:0;max-height:135mm;object-fit:contain}.ct-image{width:33.333%;font-size:11px}.bev-image{max-width:225mm}.note{font-size:11px}a{color:#16324A;text-decoration:none}}
";
    }
}
