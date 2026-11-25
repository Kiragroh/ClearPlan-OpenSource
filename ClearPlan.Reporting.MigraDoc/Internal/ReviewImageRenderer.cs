using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using ClearPlan.Core.Review;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Reporting.MigraDoc.Internal
{
    internal static class ReviewImageRenderer
    {
        private static readonly Color Back = Color.FromArgb(13, 24, 40);
        private static readonly Color Teal = Color.FromArgb(75, 220, 195);

        public static ReviewControlPointSample StartControlPoint(ReviewBeamAnalysis beam)
        {
            return beam == null ? null : (beam.ControlPoints ?? new List<ReviewControlPointSample>())
                .FirstOrDefault(cp => cp != null && cp.Index == 0);
        }

        public static string StartAngles(ReviewControlPointSample cp)
        {
            return cp == null ? "Field start · CP 0 unavailable." : string.Format(CultureInfo.InvariantCulture,
                "Field start · CP {0}; gantry {1:0.#} deg; collimator {2:0.#} deg; couch {3:0.#} deg.",
                cp.Index, cp.GantryAngleDegrees, cp.CollimatorAngleDegrees, cp.PatientSupportAngleDegrees);
        }

        public static string PqmUnavailableReason(ReviewReportPqmRow row)
        {
            if (row == null || row.Status == ReviewStatusCodes.Pass || row.Status == ReviewStatusCodes.Fail || row.Status == ReviewStatusCodes.Variation) return "";
            if (string.IsNullOrWhiteSpace(row.ResolvedStructureId)) return "Structure unmatched.";
            string detail = (row.Explanation ?? "").ToLowerInvariant();
            if (!row.AchievedValue.HasValue && detail.Contains("dvh")) return "DVH unavailable.";
            if (!row.AchievedValue.HasValue) return "Measurement unavailable.";
            if (detail.Contains("confirm") || detail.Contains("scope")) return "Goal scope unconfirmed.";
            return "Goal not assessed.";
        }

        public static string PamTargetSummary(ReviewPlanAnalysis plan)
        {
            string target = string.IsNullOrWhiteSpace(plan.TargetStructureId) ? "Unavailable" : plan.TargetStructureId;
            string summary = "PAM target: " + target;
            if (string.Equals(plan.TargetSelectionMode, "AutomaticLowestPam", StringComparison.Ordinal))
                return summary + string.Format(CultureInfo.InvariantCulture,
                    " · Automatic lowest valid PAM · {0}/{1} eligible PTVs valid. Numerical selection, not clinical superiority.",
                    plan.PamValidTargetCandidateCount, plan.PamTargetCandidateCount);
            if (string.Equals(plan.TargetSelectionMode, "Explicit", StringComparison.Ordinal)) summary += " · Explicit selection";
            return summary;
        }

        public static string PamUnavailableReason(ReviewPlanAnalysis plan)
        {
            if (plan.TargetSelectionMode == "AutomaticLowestPam" && plan.PamValidTargetCandidateCount == 0)
                return "No valid PTV candidate.";
            return string.IsNullOrWhiteSpace(plan.PamReason) ? "Required geometry unavailable." : plan.PamReason;
        }

        public static ReviewControlPointSample RepresentativeControlPoint(ReviewBeamAnalysis beam)
        {
            if (beam == null) return null;
            var candidates = (beam.ControlPoints ?? new List<ReviewControlPointSample>())
                .Where(cp => cp != null && cp.MetricMetersetWeightMu.HasValue && Positive(cp.MetricMetersetWeightMu.Value) &&
                    CompleteAperture(cp.Aperture, beam.MlcLayerCount)).OrderBy(cp => cp.Index).ToList();
            return candidates.Count == 0 ? null : candidates[candidates.Count / 2];
        }

        private static bool CompleteAperture(ApertureGeometry aperture, int expectedLayers)
        {
            if (aperture == null || aperture.Layers == null || aperture.EffectiveOpenings == null ||
                (expectedLayers > 0 && aperture.Layers.Count != expectedLayers) ||
                !aperture.EffectiveOpenings.Any(r => r != null && Finite(r.X1) && Finite(r.X2) && Finite(r.Y1) && Finite(r.Y2) && r.X2 > r.X1 && r.Y2 > r.Y1)) return false;
            return aperture.Layers.All(layer => layer != null && layer.LeafBoundariesMm != null &&
                layer.Bank1PositionsMm != null && layer.Bank2PositionsMm != null && layer.Bank1PositionsMm.Length > 0 &&
                layer.Bank1PositionsMm.Length == layer.Bank2PositionsMm.Length &&
                layer.LeafBoundariesMm.Length == layer.Bank1PositionsMm.Length + 1 &&
                layer.LeafBoundariesMm.Concat(layer.Bank1PositionsMm).Concat(layer.Bank2PositionsMm).All(Finite) &&
                Enumerable.Range(0, layer.Bank1PositionsMm.Length).All(i => layer.LeafBoundariesMm[i + 1] > layer.LeafBoundariesMm[i] && layer.Bank2PositionsMm[i] >= layer.Bank1PositionsMm[i]));
        }

        // Separate runs preserve missing-data gaps: the report never invents intermediate control-point values.
        public static List<List<Tuple<double, double>>> TraceRuns(IList<ReviewControlPointSample> samples, Func<ReviewControlPointSample, double?> value)
        {
            var runs = new List<List<Tuple<double, double>>>();
            List<Tuple<double, double>> current = null;
            int? previousIndex = null;
            foreach (var sample in (samples ?? new List<ReviewControlPointSample>()).Where(cp => cp != null).OrderBy(cp => cp.Index))
            {
                double? measurement = value(sample);
                if (!measurement.HasValue || !Finite(measurement.Value) || measurement.Value < 0)
                { current = null; previousIndex = null; continue; }
                if (current == null || (previousIndex.HasValue && sample.Index != previousIndex.Value + 1))
                { current = new List<Tuple<double, double>>(); runs.Add(current); }
                current.Add(Tuple.Create((double)sample.Index, measurement.Value));
                previousIndex = sample.Index;
            }
            return runs;
        }

        public static string ControlPointTraces(ReviewBeamAnalysis beam)
        {
            using (var output = new Bitmap(900, 1000))
            using (var g = Graphics.FromImage(output))
            {
                g.Clear(Color.White); g.SmoothingMode = SmoothingMode.AntiAlias;
                var samples = beam.ControlPoints ?? new List<ReviewControlPointSample>();
                var planned = TraceRuns(samples, cp => cp.PlannedDoseRateMuPerMin);
                var estimated = TraceRuns(samples, cp => beam.DoseRateEstimateStatus == "Estimated" ? cp.EstimatedDoseRateMuPerMin : null);
                var area = TraceRuns(samples, cp => cp.ApertureAreaCm2);
                var indices = samples.Where(cp => cp != null).Select(cp => cp.Index).ToList();
                double first = indices.Count == 0 ? 0 : indices.Min(), last = indices.Count == 0 ? 1 : indices.Max();
                if (last <= first) last = first + 1;
                PlotTrace(g, new RectangleF(100, 65, 770, 300), first, last,
                    new[] { planned, estimated }, new[] { Color.FromArgb(22, 50, 74), Color.FromArgb(5, 148, 135) },
                    new[] { DashStyle.Solid, DashStyle.Dash },
                    estimated.Count > 0 ? "Estimated dose rate [MU/min]" : "Planned dose rate [MU/min]",
                    estimated.Count > 0 ? (planned.Count > 0 ? "Solid: supplied; dashed: estimated. Not measured." : "Dashed: estimated segment average, not measured") : "Supplied plan values only; not measured delivery",
                    "No valid rate estimate or supplied rate data");
                PlotTrace(g, new RectangleF(100, 565, 770, 300), first, last,
                    new[] { area }, new[] { Color.FromArgb(198, 130, 34) }, new[] { DashStyle.Solid },
                    "Effective aperture area [cm2]", "MLC layer intersection; jaws when present");
                return Encode(output);
            }
        }

        private static void PlotTrace(Graphics g, RectangleF box, double first, double last,
            IList<List<List<Tuple<double, double>>>> series, Color[] colors, DashStyle[] styles, string title, string legend,
            string unavailableText = "Aperture geometry unavailable")
        {
            double maximum = Math.Max(1, series.SelectMany(runs => runs).SelectMany(run => run).Select(point => point.Item2).DefaultIfEmpty(0).Max() * 1.1);
            Func<double, float> x = value => box.Left + (float)((value - first) / (last - first) * box.Width);
            Func<double, float> y = value => box.Bottom - (float)(value / maximum * box.Height);
            using (var font = new Font("Segoe UI", 19))
            using (var brush = new SolidBrush(Color.FromArgb(38, 55, 71)))
            using (var grid = new Pen(Color.FromArgb(221, 230, 238)))
            {
                g.DrawString(title, font, brush, box.Left, box.Top - 52);
                if (!series.Any(runs => runs.Count > 0))
                {
                    Text(g, unavailableText, font, brush, box.Left + box.Width / 2, box.Top + box.Height / 2);
                    using (var small = new Font("Segoe UI", 15)) Text(g, legend, small, brush, box.Left + box.Width / 2, box.Bottom + 88);
                    return;
                }
                for (int tick = 0; tick <= 4; tick++)
                {
                    double ordinate = maximum * tick / 4;
                    g.DrawLine(grid, box.Left, y(ordinate), box.Right, y(ordinate));
                    Text(g, ordinate.ToString("0.#", CultureInfo.InvariantCulture), font, brush, box.Left - 37, y(ordinate));
                    double abscissa = Math.Round(first + (last - first) * tick / 4);
                    g.DrawLine(grid, x(abscissa), box.Top, x(abscissa), box.Bottom);
                    Text(g, abscissa.ToString("0", CultureInfo.InvariantCulture), font, brush, x(abscissa), box.Bottom + 20);
                }
                Text(g, "Control-point index (not time)", font, brush, box.Left + box.Width / 2, box.Bottom + 50);
                using (var small = new Font("Segoe UI", 15)) Text(g, legend, small, brush, box.Left + box.Width / 2, box.Bottom + 88);
                for (int i = 0; i < series.Count; i++)
                    using (var pen = new Pen(colors[i], 3) { DashStyle = styles[i] })
                    using (var fill = new SolidBrush(colors[i]))
                        foreach (var run in series[i])
                        {
                            var points = run.Select(p => new PointF(x(p.Item1), y(p.Item2))).ToArray();
                            if (points.Length >= 2) g.DrawLines(pen, points);
                            else if (points.Length == 1) g.FillEllipse(fill, points[0].X - 3, points[0].Y - 3, 6, 6);
                        }
            }
        }

        public static string Ct(ReviewPlanImage image)
        {
            // PDF, HTML and WPF share the exact physical raster/contour/isodose transform.
            return "base64:" + Convert.ToBase64String(ClearPlan.Rendering.PlanImageRenderer.Render(image, true, true, true));
        }

        public static string CtSummary(ReviewPlanImage image)
        {
            var viewport = PlanImageViewport.Calculate(image);
            string summary = (viewport.HasIsocenter ? "Isocenter-centered" : "Isocenter unavailable; image-centered") +
                string.Format(CultureInfo.InvariantCulture, " · FOV {0:0} mm", viewport.HalfExtentMillimeters * 2);
            var window = System.Text.RegularExpressions.Regex.Match(image.Caption ?? "", @"\bW\s*\d+(?:\.\d+)?\s*/\s*L\s*-?\d+(?:\.\d+)?\s*HU\b");
            if (window.Success) summary += " · " + window.Value;
            if (viewport.UsesDoseRegion && image.DoseFocusRegion.CoverageLimited) summary += " · Dose extent coverage-limited";
            return summary + ".";
        }

        public static string CtUnavailableOverlays(ReviewPlanImage image)
        {
            var overlays = (image.Overlays ?? new List<ReviewImageOverlay>()).Where(item => item != null).ToList();
            var parts = new List<string>();
            foreach (string kind in new[] { "structure", "isodose" })
            {
                var items = overlays.Where(item => string.Equals(item.Kind, kind, StringComparison.OrdinalIgnoreCase)).ToList();
                string label = kind == "structure" ? "contour" : "isodose";
                int missing = items.Count(item => item.SourceStatus != ReviewStatusCodes.Available);
                if (items.Count == 0) parts.Add(label + " overlays not shown");
                else if (missing > 0) parts.Add(missing + " " + label + (missing == 1 ? " unavailable" : "s unavailable"));
            }
            return string.Join("; ", parts);
        }

        public static List<ReviewPlanImage> VisiblePlanImages(ReviewReportDocument report)
        {
            return (report.PlanImages ?? new List<ReviewPlanImage>()).Where(source => source != null &&
                string.Equals(source.PlanKey, report.ActivePlanKey, StringComparison.Ordinal)).Select(source => new ReviewPlanImage {
                PlanKey = source.PlanKey, Kind = source.Kind, Title = source.Title, Caption = source.Caption,
                SourceStatus = source.SourceStatus, UnavailableReason = source.UnavailableReason, Synthetic = source.Synthetic,
                WidthPixels = source.WidthPixels, HeightPixels = source.HeightPixels, GrayscalePixels = source.GrayscalePixels,
                PixelSpacingXMillimeters = source.PixelSpacingXMillimeters, PixelSpacingYMillimeters = source.PixelSpacingYMillimeters,
                LeftOrientation = source.LeftOrientation, RightOrientation = source.RightOrientation,
                TopOrientation = source.TopOrientation, BottomOrientation = source.BottomOrientation,
                IsocenterPixelX = source.IsocenterPixelX, IsocenterPixelY = source.IsocenterPixelY,
                OverlaySummary = source.OverlaySummary, DoseFocusRegion = source.DoseFocusRegion,
                // Only the display list is copied. Renderers do not modify detached pixels or paths.
                Overlays = (source.Overlays ?? new List<ReviewImageOverlay>()).Where(overlay => overlay != null &&
                    (!string.Equals(overlay.Kind, "structure", StringComparison.OrdinalIgnoreCase) || !report.IsStructureHidden(overlay.Label))).ToList()
            }).ToList();
        }

        public static string Bev(ReviewControlPointSample sample, bool synthetic)
        {
            using (var output = new Bitmap(720, 720))
            using (var g = Graphics.FromImage(output))
            {
                g.Clear(Back); g.SmoothingMode = SmoothingMode.AntiAlias;
                var aperture = sample == null ? null : sample.Aperture;
                double extent = 140;
                if (aperture != null)
                    foreach (var layer in aperture.Layers)
                        foreach (double v in (layer.LeafBoundariesMm ?? new double[0]).Concat(layer.Bank1PositionsMm ?? new double[0]).Concat(layer.Bank2PositionsMm ?? new double[0]))
                            if (!double.IsNaN(v) && !double.IsInfinity(v)) extent = Math.Max(extent, Math.Abs(v) * 1.1);
                if (sample != null)
                {
                    foreach (var point in (sample.TargetOutlines ?? new List<List<BeamPoint>>()).Where(outline => outline != null).SelectMany(outline => outline))
                        if (Finite(point.X) && Finite(point.Y)) extent = Math.Max(extent, Math.Max(Math.Abs(point.X), Math.Abs(point.Y)) * 1.1);
                    foreach (var strip in sample.TargetProjectionStrips ?? new List<ApertureRectangle>())
                        if (strip != null && Finite(strip.X1) && Finite(strip.X2) && Finite(strip.Y1) && Finite(strip.Y2))
                            extent = Math.Max(extent, new[] { Math.Abs(strip.X1), Math.Abs(strip.X2), Math.Abs(strip.Y1), Math.Abs(strip.Y2) }.Max() * 1.1);
                }
                Func<double, float> x = value => (float)(360 + value / extent * 290);
                Func<double, float> y = value => (float)(340 - value / extent * 290);
                if (sample != null && (sample.TargetOutlines == null || sample.TargetOutlines.Count == 0))
                    using (var fill = new SolidBrush(Color.FromArgb(75, 255, 113, 150)))
                        foreach (var strip in sample.TargetProjectionStrips ?? new List<ApertureRectangle>())
                            if (strip != null) DrawRect(g, x, y, strip, null, fill);
                using (var grid = new Pen(Color.FromArgb(38, 58, 76)))
                using (var font = new Font("Segoe UI", 12))
                using (var brush = new SolidBrush(Color.LightGray))
                {
                    for (int value = -100; value <= 100; value += 50)
                    {
                        g.DrawLine(grid, x(value), 50, x(value), 630); g.DrawLine(grid, 70, y(value), 650, y(value));
                        Text(g, value.ToString(CultureInfo.InvariantCulture), font, brush, x(value), 645);
                    }
                    Text(g, "IEC beam-limiting-device X [mm]", font, brush, 360, 668);
                }
                if (aperture != null)
                {
                    Color[] colors = { Color.FromArgb(92, 164, 255), Color.FromArgb(255, 186, 83) };
                    for (int layerIndex = 0; layerIndex < aperture.Layers.Count; layerIndex++)
                    {
                        var layer = aperture.Layers[layerIndex];
                        if (layer.Bank1PositionsMm == null || layer.Bank2PositionsMm == null || layer.LeafBoundariesMm == null ||
                            layer.Bank1PositionsMm.Length != layer.Bank2PositionsMm.Length || layer.LeafBoundariesMm.Length != layer.Bank1PositionsMm.Length + 1) continue;
                        using (var pen = new Pen(colors[layerIndex % 2], 1.2f))
                        using (var fill = new SolidBrush(Color.FromArgb(38, colors[layerIndex % 2])))
                        for (int i = 0; i < layer.Bank1PositionsMm.Length; i++)
                        {
                            bool travelY = string.Equals(layer.LeafTravelAxis, "Y", StringComparison.OrdinalIgnoreCase);
                            DrawRect(g, x, y, travelY ? new ApertureRectangle(layer.LeafBoundariesMm[i], -extent, layer.LeafBoundariesMm[i + 1], layer.Bank1PositionsMm[i]) : new ApertureRectangle(-extent, layer.LeafBoundariesMm[i], layer.Bank1PositionsMm[i], layer.LeafBoundariesMm[i + 1]), pen, fill);
                            DrawRect(g, x, y, travelY ? new ApertureRectangle(layer.LeafBoundariesMm[i], layer.Bank2PositionsMm[i], layer.LeafBoundariesMm[i + 1], extent) : new ApertureRectangle(layer.Bank2PositionsMm[i], layer.LeafBoundariesMm[i], extent, layer.LeafBoundariesMm[i + 1]), pen, fill);
                        }
                    }
                    using (var pen = new Pen(Teal, 1.5f))
                    using (var fill = new SolidBrush(Color.FromArgb(60, Teal)))
                        foreach (var opening in aperture.EffectiveOpenings) DrawRect(g, x, y, opening, pen, fill);
                    if (aperture.Jaws != null)
                        using (var pen = new Pen(Color.White, 3)) DrawRect(g, x, y, aperture.Jaws, pen, null);
                }
                if (sample != null && sample.TargetOutlines != null)
                    using (var target = new Pen(Color.FromArgb(255, 113, 150), 3))
                        foreach (var outline in sample.TargetOutlines)
                            if (outline != null && outline.Count >= 3)
                                g.DrawPolygon(target, outline.Select(point => new PointF(x(point.X), y(point.Y))).ToArray());
                if (aperture == null)
                    using (var font = new Font("Segoe UI", 16, FontStyle.Bold))
                    using (var brush = new SolidBrush(Color.White))
                        Text(g, "MLC geometry unavailable", font, brush, 360, 335);
                if (synthetic) SimulationLabel(g, 720, 720);
                return Encode(output);
            }
        }

        public static string Dvh(ReviewReportDocument report)
        {
            var series = report.VisibleDvhSeries().Where(row => (row.Selected || row.RequiredForTargetReview) && row.Points != null && row.Points.Count > 1).ToList();
            using (var legendFont = new Font("Segoe UI", 14))
            using (var measurement = new Bitmap(1, 1))
            using (var measure = Graphics.FromImage(measurement))
            {
                var legend = DvhLegendLayout(measure, legendFont, series);
                int height = (int)Math.Ceiling(legend.Select(entry => entry.LabelBounds.Bottom).DefaultIfEmpty(480).Max() + 24);
                using (var output = new Bitmap(1400, height))
                using (var g = Graphics.FromImage(output))
                using (var font = new Font("Segoe UI", 17))
                using (var brush = new SolidBrush(Color.FromArgb(38, 55, 71)))
                {
                    g.Clear(Color.White); g.SmoothingMode = SmoothingMode.AntiAlias;
                    double max = Math.Max(1, series.SelectMany(row => row.Points).Select(p => p.DoseGy).DefaultIfEmpty(1).Max());
                    Func<double, float> x = value => (float)(100 + 1240 * value / max);
                    Func<double, float> y = value => (float)(400 - 320 * value / 100);
                    using (var grid = new Pen(Color.FromArgb(221, 230, 238)))
                    {
                        for (int v = 0; v <= 100; v += 20)
                        { g.DrawLine(grid, 100, y(v), 1340, y(v)); Text(g, v.ToString(), font, brush, 65, y(v)); }
                        for (int d = 0; d <= 5; d++)
                        { g.DrawLine(grid, x(max * d / 5), 80, x(max * d / 5), 400); Text(g, (max * d / 5).ToString("0.#", CultureInfo.InvariantCulture), font, brush, x(max * d / 5), 425); }
                    }
                    Text(g, "Dose [Gy]", font, brush, 720, 465);
                    g.DrawString("Cumulative volume [%]", font, brush, 90, 20);
                    foreach (var entry in legend)
                    {
                        var row = entry.Series;
                        Color color;
                        try { color = ColorTranslator.FromHtml(row.ColorHex); } catch { color = Teal; }
                        using (var pen = new Pen(color, 3))
                        {
                            if (row.LineStyle == "dash") pen.DashStyle = DashStyle.Dash;
                            if (row.LineStyle == "dot") pen.DashStyle = DashStyle.Dot;
                            g.DrawLines(pen, row.Points.Select(p => new PointF(x(p.DoseGy), y(p.VolumePercent))).ToArray());
                            float swatchY = entry.LabelBounds.Top + legendFont.GetHeight(g) / 2;
                            g.DrawLine(pen, entry.LabelBounds.Left - 50, swatchY, entry.LabelBounds.Left - 12, swatchY);
                            g.DrawString(entry.Label, legendFont, brush, entry.LabelBounds);
                        }
                    }
                    if (series.Count == 0) Text(g, "No selected DVH series", font, brush, 720, 290);
                    return Encode(output);
                }
            }
        }

        private sealed class DvhLegendEntry
        {
            public ReviewReportDvhSeries Series;
            public string Label;
            public RectangleF LabelBounds;
        }

        private static List<DvhLegendEntry> DvhLegendLayout(Graphics graphics, Font font, IList<ReviewReportDvhSeries> series)
        {
            // Flow entries into compact rows; long IDs wrap inside their own entry.
            // Measure and draw with the same GDI text layout, without shortening IDs.
            var result = new List<DvhLegendEntry>();
            float left = 40, top = 500, rowHeight = 0;
            foreach (var row in series)
            {
                string label = row.DisplayName ?? row.StructureId ?? "";
                float labelWidth = Math.Min(350, graphics.MeasureString(label, font).Width + 4);
                var measured = graphics.MeasureString(label, font, new SizeF(Math.Max(1, labelWidth), 10000));
                float entryWidth = labelWidth + 74;
                if (left > 40 && left + entryWidth > 1360)
                {
                    left = 40;
                    top += rowHeight + 10;
                    rowHeight = 0;
                }
                var bounds = new RectangleF(left + 50, top, Math.Max(1, labelWidth), measured.Height + 2);
                result.Add(new DvhLegendEntry { Series = row, Label = label, LabelBounds = bounds });
                left += entryWidth;
                rowHeight = Math.Max(rowHeight, bounds.Height);
            }
            return result;
        }

        private static void DrawRect(Graphics g, Func<double, float> x, Func<double, float> y, ApertureRectangle r, Pen pen, Brush fill)
        {
            float left = x(r.X1), top = y(r.Y2), width = x(r.X2) - left, height = y(r.Y1) - top;
            if (width <= 0 || height <= 0 || float.IsInfinity(width) || float.IsNaN(width) || float.IsInfinity(height) || float.IsNaN(height)) return;
            if (fill != null) g.FillRectangle(fill, left, top, width, height);
            if (pen != null) g.DrawRectangle(pen, left, top, width, height);
        }
        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static bool Positive(double value) { return value > 0 && !double.IsNaN(value) && !double.IsInfinity(value); }
        private static void Text(Graphics g, string value, Font font, Brush brush, float x, float y)
        {
            using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                g.DrawString(value ?? "", font, brush, x, y, format);
        }
        private static void SimulationLabel(Graphics g, int width, int height)
        {
            using (var fill = new SolidBrush(Color.FromArgb(190, 123, 20, 31))) g.FillRectangle(fill, 0, height - 29, width, 29);
            using (var font = new Font("Segoe UI", 14, FontStyle.Bold)) Text(g, "SIMULATION - NOT PATIENT DATA", font, Brushes.White, width / 2, height - 15);
        }
        private static string Encode(Bitmap image)
        {
            using (var stream = new MemoryStream()) { image.Save(stream, ImageFormat.Png); return "base64:" + Convert.ToBase64String(stream.ToArray()); }
        }
    }
}
