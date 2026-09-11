using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Rendering
{
    /// <summary>Explicitly separates an unavailable DRR from an unavailable aperture.</summary>
    public sealed class BeamEyeViewRenderState
    {
        public bool ImageAvailable { get; internal set; }
        public string ImageMessage { get; internal set; }
        public bool GeometryAvailable { get; internal set; }
        public string GeometryMessage { get; internal set; }
        public bool PhysicalJawsVisible { get; internal set; }
        public bool FixedBoundsVisible { get; internal set; }
        public bool Synthetic { get; internal set; }
        public string LayoutDescription { get; internal set; }
        public string[] JawLabels { get; internal set; }
    }

    /// <summary>
    /// Original, detached-data-only renderer shared by the viewer and PDF report.
    /// THESIS: let the CT projection remain the dominant evidence, with precise outlines over it.
    /// OWN-WORLD: IEC beam-limiting-device millimetres, real jaw semantics, exact control points.
    /// STORY: inspect the field, verify its source and settings, then orient the schematic.
    /// FIRST VIEWPORT: a square radiograph and a quiet white Clinical Blueprint information rail.
    /// FORM: bounded extension of the existing Clinical Blueprint; no new visual identity.
    /// FINISH: decoded PNG tests, geometry/source rejection tests, compact-size visual inspection.
    /// No vendor GUI assets, patient images or third-party viewer drawing code are embedded.
    /// </summary>
    public static class BeamEyeViewRenderer
    {
        private static readonly Color Ink = Color.FromArgb(32, 43, 56);
        private static readonly Color Muted = Color.FromArgb(86, 103, 121);
        private static readonly Color Border = Color.FromArgb(214, 221, 228);
        private static readonly Color Canvas = Color.FromArgb(23, 33, 43);
        private static readonly Color Teal = Color.FromArgb(15, 118, 110);
        private static readonly Color Amber = Color.FromArgb(255, 196, 93);
        private static readonly Color Cyan = Color.FromArgb(71, 215, 247);
        private static readonly Color JawColor = Color.FromArgb(255, 230, 96);
        private static readonly Color SimulationRed = Color.FromArgb(180, 35, 24);
        private const float LogicalWidth = 1500;
        private const float LogicalHeight = 1000;
        private static readonly RectangleF Viewport = new RectangleF(50, 54, 900, 900);

        public static byte[] Render(ReviewBeamAnalysis beam, ReviewControlPointSample cp, BeamEyeViewImage image, bool synthetic, int width = 1500, int height = 1000, bool compactStatus = false)
        {
            if (width < 640 || height < 400 || width > 6000 || height > 4000 || (long)width * height > 18000000)
                throw new ArgumentOutOfRangeException("width", "BEV output must be 640..6000 by 400..4000 pixels, at most 18 megapixels.");
            var state = InspectState(beam, cp, image, synthetic);
            using (var bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                float scale = Math.Min(width / LogicalWidth, height / LogicalHeight);
                graphics.TranslateTransform((width - LogicalWidth * scale) / 2, (height - LogicalHeight * scale) / 2);
                graphics.ScaleTransform(scale, scale);
                Fill(graphics, Canvas, new RectangleF(0, 0, 1000, LogicalHeight));
                Fill(graphics, Color.White, new RectangleF(1000, 0, 500, LogicalHeight));
                Text(graphics, "Beam's-eye view", new RectangleF(50, 12, 320, 34), 25, Color.White, true);
                Text(graphics, "Isocentre plane · BLD coordinates · cm", new RectangleF(367, 17, 583, 30), 21, Color.FromArgb(192, 205, 217), false, StringAlignment.Far);

                double extent = state.ImageAvailable ? image.ExtentMm : GeometryExtent(cp, state.GeometryAvailable);
                if (state.ImageAvailable)
                {
                    using (var pixels = GrayscaleBitmap(image))
                    {
                        graphics.InterpolationMode = InterpolationMode.HighQualityBilinear;
                        graphics.DrawImage(pixels, Viewport);
                    }
                }
                else
                    Fill(graphics, Color.FromArgb(25, 32, 40), Viewport);

                var clipping = graphics.Save();
                graphics.SetClip(Viewport);
                if (state.GeometryAvailable) DrawAperture(graphics, cp.Aperture, extent);
                DrawRuler(graphics, extent);
                graphics.Restore(clipping);
                using (var pen = new Pen(Color.FromArgb(104, 121, 136), 1)) graphics.DrawRectangle(pen, Viewport.X, Viewport.Y, Viewport.Width, Viewport.Height);

                // The native viewer exposes diagnostics in its accessible status control.
                // Reports keep the full in-image notice by default.
                if (!state.ImageAvailable && !compactStatus) DrawUnavailable(graphics, state.ImageMessage);
                if (state.GeometryAvailable) DrawApertureLegend(graphics, cp.Aperture, state);
                if (state.PhysicalJawsVisible) DrawJawLabels(graphics, cp.Aperture.Jaws, extent);
                if (!state.GeometryAvailable)
                    OverlayLabel(graphics, "MLC / jaw geometry unavailable", new RectangleF(68, 73, 420, 40), Amber, 23);

                DrawInformation(graphics, beam, cp, image, state, compactStatus);
                using (var pen = new Pen(Border, 1)) graphics.DrawLine(pen, 1000, 0, 1000, 967);
                if (state.Synthetic)
                {
                    Fill(graphics, SimulationRed, new RectangleF(0, 967, LogicalWidth, 33));
                    Text(graphics, "SIMULATION — NOT PATIENT DATA — NOT FOR CLINICAL USE", new RectangleF(18, 969, 1464, 28), 22, Color.White, true, StringAlignment.Center);
                }
                using (var stream = new MemoryStream())
                {
                    bitmap.Save(stream, ImageFormat.Png);
                    return stream.ToArray();
                }
            }
        }

        public static BeamEyeViewRenderState InspectState(ReviewBeamAnalysis beam, ReviewControlPointSample cp, BeamEyeViewImage image, bool synthetic)
        {
            var state = new BeamEyeViewRenderState
            {
                Synthetic = synthetic || (image != null && image.Synthetic),
                JawLabels = new string[0],
                LayoutDescription = "Geometry unavailable"
            };
            state.ImageMessage = ImageProblem(cp, image, synthetic);
            state.ImageAvailable = state.ImageMessage == null;
            if (state.ImageAvailable) state.ImageMessage = image.Synthetic ? "Synthetic CT projection" : "CT-derived projection";
            state.GeometryMessage = GeometryProblem(beam, cp);
            state.GeometryAvailable = state.GeometryMessage == null;
            if (state.GeometryAvailable)
            {
                state.PhysicalJawsVisible = cp.Aperture.Jaws != null;
                state.FixedBoundsVisible = cp.Aperture.FixedBoundingBox != null;
                state.JawLabels = state.PhysicalJawsVisible ? new[] { "X1", "X2", "Y1", "Y2" } : new string[0];
                state.LayoutDescription = (state.PhysicalJawsVisible ? "Physical jaws" : "Jawless") + " · " + cp.Aperture.Layers.Count + " MLC " + (cp.Aperture.Layers.Count == 1 ? "layer" : "layers");
                state.GeometryMessage = state.LayoutDescription;
            }
            return state;
        }

        /// <summary>
        /// Normalized source position in the explicitly labelled fixed front-room schematic.
        /// This is not a patient coordinate transform: 0 degrees is above isocentre, 90 right.
        /// Couch and collimator angles must never rotate this room-camera gantry position.
        /// </summary>
        public static PointF LinacOrientationPoint(double gantryDegrees)
        {
            if (!Finite(gantryDegrees)) throw new ArgumentException("A finite gantry angle is required.", "gantryDegrees");
            double radians = (gantryDegrees % 360) * Math.PI / 180;
            return new PointF((float)Math.Sin(radians), (float)-Math.Cos(radians));
        }

        private static string ImageProblem(ReviewControlPointSample cp, BeamEyeViewImage image, bool synthetic)
        {
            if (cp == null) return "No selected control point. No CT projection displayed.";
            if (cp.Index < 0) return "Invalid control-point index. No CT projection displayed.";
            if (image == null) return "No CT projection supplied for this control point.";
            if (image.Synthetic != synthetic) return "CT projection source mode does not match this review. Pixels withheld.";
            if (image.ControlPointIndex != cp.Index) return "CT projection belongs to another control point. Pixels withheld.";
            if (!MatchingAngle(image.GantryAngleDegrees, cp.GantryAngleDegrees) ||
                !MatchingAngle(image.CollimatorAngleDegrees, cp.CollimatorAngleDegrees) ||
                !MatchingAngle(image.PatientSupportAngleDegrees, cp.PatientSupportAngleDegrees))
                return "CT projection angles do not match this control point. Pixels withheld.";
            bool availableSource = string.Equals(image.SourceStatus, "available", StringComparison.OrdinalIgnoreCase) ||
                (synthetic && image.Synthetic && string.Equals(image.SourceStatus, "synthetic", StringComparison.OrdinalIgnoreCase));
            if (!availableSource)
                return Clean(image.UnavailableReason, "The CT projection source reports unavailable.", 260);
            if (image.WidthPixels < 2 || image.HeightPixels < 2 || image.WidthPixels > 4096 || image.HeightPixels > 4096 ||
                image.GrayscalePixels == null || (long)image.WidthPixels * image.HeightPixels != image.GrayscalePixels.LongLength ||
                !Finite(image.ExtentMm) || image.ExtentMm <= 0 || image.ExtentMm > 2000)
                return "Invalid CT projection dimensions, pixels or physical extent. Pixels withheld.";
            return null;
        }

        private static string GeometryProblem(ReviewBeamAnalysis beam, ReviewControlPointSample cp)
        {
            if (beam == null || cp == null || cp.Aperture == null)
                return Clean(cp == null ? null : cp.GeometryReason, "No complete detached aperture geometry supplied.", 260);
            var aperture = cp.Aperture;
            if (!beam.HasJaws.HasValue) return "Physical-jaw presence has not been verified.";
            if (beam.HasJaws.Value != (aperture.Jaws != null)) return "Physical-jaw metadata and aperture disagree. Geometry withheld.";
            if (aperture.Jaws != null && !ValidRectangle(aperture.Jaws, false)) return "Physical-jaw coordinates are invalid.";
            if (aperture.FixedBoundingBox != null && (beam.HasJaws.Value || !ValidRectangle(aperture.FixedBoundingBox, true)))
                return "Fixed virtual field boundary is invalid or conflicts with physical jaws.";
            if (aperture.Layers == null || aperture.Layers.Count > 8 || beam.MlcLayerCount != aperture.Layers.Count ||
                (aperture.Layers.Count == 0 && aperture.Jaws == null))
                return "Not all declared MLC layers are available. Geometry withheld.";
            foreach (var layer in aperture.Layers)
            {
                if (layer == null || layer.LeafBoundariesMm == null || layer.Bank1PositionsMm == null || layer.Bank2PositionsMm == null ||
                    layer.Bank1PositionsMm.Length < 1 || layer.Bank1PositionsMm.Length > 400 ||
                    layer.LeafBoundariesMm.Length != layer.Bank1PositionsMm.Length + 1 || layer.Bank2PositionsMm.Length != layer.Bank1PositionsMm.Length ||
                    (layer.LeafTravelAxis != "X" && layer.LeafTravelAxis != "Y") ||
                    !layer.LeafBoundariesMm.All(Physical) || !layer.Bank1PositionsMm.All(Physical) || !layer.Bank2PositionsMm.All(Physical))
                    return "A leaf layer has incomplete or invalid physical boundaries / bank positions.";
                for (int i = 1; i < layer.LeafBoundariesMm.Length; i++)
                    if (layer.LeafBoundariesMm[i] <= layer.LeafBoundariesMm[i - 1]) return "Leaf boundaries are not strictly ordered.";
                // Crossed or touching opposing leaf tips represent a closed pair, not fabricated openings.
            }
            if (aperture.EffectiveOpenings == null || aperture.EffectiveOpenings.Any(r => !ValidRectangle(r, false)))
                return "Effective aperture rectangles are incomplete or invalid.";
            return null;
        }

        private static Bitmap GrayscaleBitmap(BeamEyeViewImage image)
        {
            var bitmap = new Bitmap(image.WidthPixels, image.HeightPixels, PixelFormat.Format24bppRgb);
            var data = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);
            try
            {
                var row = new byte[data.Stride];
                for (int y = 0; y < image.HeightPixels; y++)
                {
                    for (int x = 0; x < image.WidthPixels; x++)
                    {
                        byte value = image.GrayscalePixels[y * image.WidthPixels + x];
                        row[x * 3] = row[x * 3 + 1] = row[x * 3 + 2] = value;
                    }
                    Marshal.Copy(row, 0, IntPtr.Add(data.Scan0, y * data.Stride), row.Length);
                }
            }
            finally { bitmap.UnlockBits(data); }
            return bitmap;
        }

        private static void DrawAperture(Graphics g, ApertureGeometry aperture, double extent)
        {
            for (int layerIndex = 0; layerIndex < aperture.Layers.Count; layerIndex++)
            {
                var layer = aperture.Layers[layerIndex];
                Color color = LayerColor(layerIndex);
                using (var outline = new Pen(Color.FromArgb(210, color), 1.65f))
                using (var tint = new SolidBrush(Color.FromArgb(12, color)))
                {
                    for (int pair = 0; pair < layer.Bank1PositionsMm.Length; pair++)
                    {
                        double low = layer.LeafBoundariesMm[pair], high = layer.LeafBoundariesMm[pair + 1];
                        double tip1 = layer.Bank1PositionsMm[pair], tip2 = layer.Bank2PositionsMm[pair];
                        var bank1 = layer.LeafTravelAxis == "X" ? new ApertureRectangle(-extent, low, tip1, high) : new ApertureRectangle(low, -extent, high, tip1);
                        var bank2 = layer.LeafTravelAxis == "X" ? new ApertureRectangle(tip2, low, extent, high) : new ApertureRectangle(low, tip2, high, extent);
                        DrawLeaf(g, bank1, extent, tint, outline);
                        DrawLeaf(g, bank2, extent, tint, outline);
                    }
                }
            }
            // Only outlines: an opaque effective-opening fill would obscure the radiograph.
            if (aperture.FixedBoundingBox != null)
                using (var pen = new Pen(Color.FromArgb(224, 230, 235), 2))
                {
                    pen.DashStyle = DashStyle.Dash;
                    DrawRectangle(g, pen, Map(aperture.FixedBoundingBox, extent));
                }
            if (aperture.Jaws != null)
                using (var pen = new Pen(JawColor, 2.6f)) DrawRectangle(g, pen, Map(aperture.Jaws, extent));
        }

        private static void DrawLeaf(Graphics g, ApertureRectangle leaf, double extent, Brush tint, Pen outline)
        {
            if (leaf.X2 <= leaf.X1 || leaf.Y2 <= leaf.Y1) return;
            var rectangle = Map(leaf, extent);
            g.FillRectangle(tint, rectangle);
            DrawRectangle(g, outline, rectangle);
        }

        private static void DrawRuler(Graphics g, double extent)
        {
            PointF centre = Map(0, 0, extent);
            using (var pen = new Pen(Color.FromArgb(172, 235, 230, 185), 1.2f))
            {
                pen.DashPattern = new float[] { 7, 5 };
                g.DrawLine(pen, Viewport.Left, centre.Y, Viewport.Right, centre.Y);
                g.DrawLine(pen, centre.X, Viewport.Top, centre.X, Viewport.Bottom);
                pen.DashStyle = DashStyle.Solid;
                int tickMax = Math.Min(200, (int)Math.Floor(extent / 10));
                for (int tick = -tickMax; tick <= tickMax; tick++)
                {
                    var horizontal = Map(tick * 10, 0, extent);
                    var vertical = Map(0, tick * 10, extent);
                    float length = tick % 5 == 0 ? 10 : 5;
                    g.DrawLine(pen, horizontal.X, centre.Y - length, horizontal.X, centre.Y + length);
                    g.DrawLine(pen, centre.X - length, vertical.Y, centre.X + length, vertical.Y);
                    if (tick != 0 && tick % 5 == 0 && Math.Abs(tick * 10) < extent - 15)
                    {
                        OverlayLabel(g, tick.ToString(CultureInfo.InvariantCulture), new RectangleF(horizontal.X - 30, centre.Y + 14, 60, 28), Color.FromArgb(242, 239, 211), 19);
                        OverlayLabel(g, tick.ToString(CultureInfo.InvariantCulture), new RectangleF(centre.X + 14, vertical.Y - 14, 60, 28), Color.FromArgb(242, 239, 211), 19);
                    }
                }
                g.DrawEllipse(pen, centre.X - 7, centre.Y - 7, 14, 14);
            }
            OverlayLabel(g, "+Y", new RectangleF(centre.X + 12, Viewport.Top + 8, 60, 30), Color.FromArgb(244, 241, 214), 21);
            OverlayLabel(g, "+X", new RectangleF(Viewport.Right - 67, centre.Y - 49, 60, 30), Color.FromArgb(244, 241, 214), 21);
            OverlayLabel(g, "ISO · 0,0 cm", new RectangleF(centre.X+18,centre.Y+49,166,32),Color.White,22);
        }

        private static void DrawJawLabels(Graphics g, ApertureRectangle jaws, double extent)
        {
            RectangleF rectangle = Map(jaws, extent);
            float midY = Clamp((rectangle.Top + rectangle.Bottom) / 2, Viewport.Top + 140, Viewport.Bottom - 80);
            float midX = Clamp((rectangle.Left + rectangle.Right) / 2, Viewport.Left + 100, Viewport.Right - 100);
            OverlayLabel(g, "X1 " + Cm(jaws.X1), new RectangleF(Clamp(rectangle.Left - 145, 58, 782), midY - 44, 160, 34), JawColor, 22);
            OverlayLabel(g, "X2 " + Cm(jaws.X2), new RectangleF(Clamp(rectangle.Right + 8, 58, 782), midY - 44, 160, 34), JawColor, 22);
            float yLabelLeft = YJawLabelLeft(midX);
            OverlayLabel(g, "Y2 " + Cm(jaws.Y2), new RectangleF(yLabelLeft, Clamp(rectangle.Top - 43, 127, 867), 160, 34), JawColor, 22);
            OverlayLabel(g, "Y1 " + Cm(jaws.Y1), new RectangleF(yLabelLeft, Clamp(rectangle.Bottom + 8, 127, 867), 160, 34), JawColor, 22);
        }

        private static float YJawLabelLeft(float apertureMidpoint)
        {
            float axisX = Viewport.Left + Viewport.Width / 2;
            float left = Clamp(apertureMidpoint - 80, Viewport.Left + 8, Viewport.Right - 168);
            // Reserve the ruler stroke and its right-hand tick-label band, with a clear gap.
            if (left < axisX + 86 && left + 160 > axisX - 12) left = axisX - 190;
            return left;
        }

        private static void DrawApertureLegend(Graphics g, ApertureGeometry aperture, BeamEyeViewRenderState state)
        {
            float y = 72;
            for (int i = 0; i < aperture.Layers.Count; i++)
            {
                string label = Clean(aperture.Layers[i].Label, "MLC layer " + (i + 1), 40) + " · " + aperture.Layers[i].Bank1PositionsMm.Length + " pairs";
                OverlayLabel(g, label, new RectangleF(68, y, 438, 34), LayerColor(i), 22);
                OverlayLabel(g, LeafWidthLabel(aperture.Layers[i]),new RectangleF(68,y+33,438,30),LayerColor(i),20);
                y += 68;
            }
            string boundary = state.PhysicalJawsVisible ? "Solid yellow: physical jaws" : (state.FixedBoundsVisible ? "Dashed white: fixed virtual field boundary" : "Jawless · no physical jaws");
            OverlayLabel(g, boundary, new RectangleF(68, 874, 680, 34), state.PhysicalJawsVisible ? JawColor : Color.White, 22);
            OverlayLabel(g, "CT proxy, not diagnostic. Leaf bodies extend to viewport edge.", new RectangleF(68, 913, 864, 29), Color.FromArgb(213, 222, 230), 20);
        }

        private static void DrawUnavailable(Graphics g, string reason)
        {
            // Keep the isocenter and central aperture inspectable even when there is no CT projection.
            Fill(g, Color.FromArgb(228, 23, 33, 43), new RectangleF(160, 705, 680, 153));
            Text(g, "CT projection unavailable", new RectangleF(182, 718, 636, 38), 28, Color.White, true, StringAlignment.Center);
            Text(g, Clean(reason, "No CT projection supplied.", 240), new RectangleF(182, 764, 636, 83), 21, Color.FromArgb(217, 227, 236), false, StringAlignment.Center);
        }

        private static string LeafWidthLabel(ApertureLayer layer)
        {
            var widths=layer.LeafBoundariesMm.Skip(1).Select((upper,i)=>upper-layer.LeafBoundariesMm[i])
                .Distinct().OrderBy(value=>value).Select(value=>value.ToString("0.###",CultureInfo.InvariantCulture));
            return "Leaf widths at ISO: " + string.Join(" / ",widths) + " mm";
        }

        private static void DrawInformation(Graphics g, ReviewBeamAnalysis beam, ReviewControlPointSample cp, BeamEyeViewImage image, BeamEyeViewRenderState state, bool compactStatus)
        {
            const float x = 1030, w = 440;
            string beamId = beam == null ? "No beam selected" : Clean(beam.BeamId, "Unnamed beam", 60);
            Text(g, beamId, new RectangleF(x, 26, w, 46), 34, Ink, true);
            // Missing CP0 is not permission to relabel a later aperture as the field start.
            var first = beam == null || beam.ControlPoints == null ? null : beam.ControlPoints.FirstOrDefault(p => p != null && p.Index == 0);
            string selected = cp == null ? "Control point unavailable" : (cp.Index == 0 ? "Field start" : "Selected control point") + " · CP " + cp.Index.ToString(CultureInfo.InvariantCulture);
            if (beam != null && beam.BeamNumber.HasValue) selected += " · Field " + beam.BeamNumber.Value.ToString(CultureInfo.InvariantCulture);
            Text(g, selected, new RectangleF(x, 81, w, 32), 23, Teal, true);
            Rule(g, 123);
            Text(g, beam == null ? "Machine unavailable" : Clean(beam.MachineId, "Machine unavailable", 60), new RectangleF(x, 138, w, 30), 25, Ink, true);
            Text(g, beam == null ? "MLC model unavailable" : Clean(beam.MlcModel, "MLC model unavailable", 100), new RectangleF(x, 175, w, 43), 22, Muted);
            Row(g, "Energy", beam == null ? null : beam.EnergyDisplay, 228);
            Row(g, "Technique", beam == null ? null : beam.Technique, 260);
            Row(g, "Beam meterset", Value(beam == null ? null : beam.MetersetMu, "0.0", " MU"), 292);
            Row(g, "Aperture", state.LayoutDescription, 324, 20);
            Text(g, "Angles", new RectangleF(x, 367, 150, 30), 24, Ink, true);
            Text(g, "Current / field start", new RectangleF(x + 155, 371, 285, 28), 20, Muted, false, StringAlignment.Far);
            AngleRow(g, "Gantry", cp == null ? (double?)null : cp.GantryAngleDegrees, first == null ? (double?)null : first.GantryAngleDegrees, 408);
            AngleRow(g, "Collimator", cp == null ? (double?)null : cp.CollimatorAngleDegrees, first == null ? (double?)null : first.CollimatorAngleDegrees, 440);
            AngleRow(g, "Couch", cp == null ? (double?)null : cp.PatientSupportAngleDegrees, first == null ? (double?)null : first.PatientSupportAngleDegrees, 472);
            Row(g, "Rate, planned", Value(cp == null ? null : cp.PlannedDoseRateMuPerMin, "0", " MU/min"), 519);
            double? nominal = cp == null ? null : cp.NominalDoseRateMuPerMin;
            if (!nominal.HasValue && beam != null) nominal = beam.NominalDoseRateMuPerMin;
            Row(g, "Rate, nominal", Value(nominal, "0", " MU/min"), 551);
            Row(g, "Open area", Value(cp == null || !state.GeometryAvailable ? null : cp.ApertureAreaCm2, "0.00", " cm²"), 583);
            Rule(g, 626);
            if (!compactStatus)
            {
                Text(g, state.ImageAvailable ? state.ImageMessage + " available" : "DRR unavailable — pixels withheld", new RectangleF(x, 640, w, 30), 22, state.ImageAvailable ? Teal : Color.FromArgb(151, 88, 19), true);
                string detail = !state.ImageAvailable ? state.ImageMessage : (!state.GeometryAvailable ? state.GeometryMessage : Clean(image == null ? null : image.ProjectionDescription, "Detached CT-derived line-integral projection.", 320));
                Text(g, detail, new RectangleF(x, 710, w, 55), 20, Muted);
            }
            Text(g, state.GeometryAvailable ? "Complete MLC / boundary geometry" : "MLC / boundary geometry unavailable", new RectangleF(x, compactStatus ? 640 : 675, w, 28), 21, state.GeometryAvailable ? Ink : Color.FromArgb(151, 88, 19));
            DrawOrientation(g, cp);
        }

        private static void DrawOrientation(Graphics g, ReviewControlPointSample cp)
        {
            const float leftX = 1130, rightX = 1378, centreY = 858;
            Text(g, "Orientation schematic", new RectangleF(1030, 780, 440, 29), 23, Ink, true);
            if (cp == null || !Finite(cp.GantryAngleDegrees) || !Finite(cp.PatientSupportAngleDegrees))
            {
                Text(g, "Finite angles unavailable", new RectangleF(1030, 825, 440, 35), 21, Muted);
                return;
            }
            PointF source = LinacOrientationPoint(cp.GantryAngleDegrees);
            using (var pen = new Pen(Border, 2)) g.DrawEllipse(pen, leftX - 34, centreY - 34, 68, 68);
            using (var pen = new Pen(Color.FromArgb(119, 143, 156), 1))
            {
                g.DrawLine(pen, leftX - 46, centreY, leftX + 46, centreY);
                g.DrawLine(pen, leftX, centreY - 39, leftX, centreY + 39);
            }
            using (var pen = new Pen(Teal, 2.5f)) g.DrawLine(pen, leftX, centreY, leftX + source.X * 32, centreY + source.Y * 32);
            var transform = g.Save();
            g.TranslateTransform(leftX + source.X * 32, centreY + source.Y * 32);
            g.RotateTransform((float)(cp.GantryAngleDegrees % 360));
            Fill(g, Color.FromArgb(22, 50, 74), new RectangleF(-12, -8, 24, 17));
            Fill(g, Teal, new RectangleF(-7, 6, 14, 5));
            g.Restore(transform);
            using (var brush = new SolidBrush(Teal)) g.FillEllipse(brush, leftX - 4, centreY - 4, 8, 8);
            Text(g, "90°", new RectangleF(leftX + 57, centreY - 12, 51, 25), 18, Muted);

            // Separate top view: couch rotates the table and patient symbol, never the gantry.
            using (var pen = new Pen(Border, 1.5f)) g.DrawEllipse(pen, rightX - 38, centreY - 38, 76, 76);
            transform = g.Save();
            g.TranslateTransform(rightX, centreY);
            g.RotateTransform((float)(cp.PatientSupportAngleDegrees % 360));
            Fill(g, Color.FromArgb(224, 232, 237), new RectangleF(-15, -37, 30, 74));
            using (var pen = new Pen(Color.FromArgb(117, 134, 148), 1.5f)) g.DrawRectangle(pen, -15, -37, 30, 74);
            using (var brush = new SolidBrush(Teal))
            {
                g.FillEllipse(brush, -5, -25, 10, 10);
                g.FillRectangle(brush, -5, -11, 10, 30);
            }
            g.Restore(transform);
            using (var pen = new Pen(Ink, 1))
            {
                g.DrawLine(pen, rightX - 6, centreY, rightX + 6, centreY);
                g.DrawLine(pen, rightX, centreY - 6, rightX, centreY + 6);
            }
            Text(g, "Front room view", new RectangleF(1025, 901, 222, 28), 20, Muted, false, StringAlignment.Center);
            Text(g, "Couch · top view", new RectangleF(1270, 901, 220, 28), 20, Muted, false, StringAlignment.Center);
            Text(g, "0° above; +90° right. Not a collision check.", new RectangleF(1024, 935, 452, 26), 19, Muted, false, StringAlignment.Center);
        }

        private static void Row(Graphics g, string label, string value, float y, float valueSize = 22)
        {
            Text(g, label, new RectangleF(1030, y, 170, 29), 21, Muted);
            Text(g, Clean(value, "Unavailable", 100), new RectangleF(1200, y, 270, 29), valueSize, Ink, false, StringAlignment.Far);
        }

        private static void AngleRow(Graphics g, string label, double? current, double? start, float y)
        {
            Row(g, label, Value(current, "0.0", "°") + " / " + Value(start, "0.0", "°"), y);
        }

        private static void Rule(Graphics g, float y)
        {
            using (var pen = new Pen(Border, 1)) g.DrawLine(pen, 1030, y, 1470, y);
        }

        private static void OverlayLabel(Graphics g, string text, RectangleF rectangle, Color color, float size)
        {
            Fill(g, Color.FromArgb(205, Canvas), rectangle);
            Text(g, text, new RectangleF(rectangle.X + 6, rectangle.Y + 1, rectangle.Width - 12, rectangle.Height - 2), size, color);
        }

        private static void Text(Graphics g, string text, RectangleF rectangle, float size, Color color, bool bold = false, StringAlignment alignment = StringAlignment.Near)
        {
            using (var font = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel))
            using (var brush = new SolidBrush(color))
            using (var format = new StringFormat(StringFormat.GenericDefault))
            {
                format.Alignment = alignment;
                format.LineAlignment = StringAlignment.Near;
                format.Trimming = StringTrimming.EllipsisWord;
                format.FormatFlags = StringFormatFlags.LineLimit;
                g.DrawString(Clean(text, "", 500), font, brush, rectangle, format);
            }
        }

        private static void Fill(Graphics g, Color color, RectangleF rectangle)
        {
            using (var brush = new SolidBrush(color)) g.FillRectangle(brush, rectangle);
        }

        private static PointF Map(double x, double y, double extent)
        {
            return new PointF(Viewport.X + (float)((x + extent) / (2 * extent)) * Viewport.Width,
                Viewport.Y + (float)((extent - y) / (2 * extent)) * Viewport.Height);
        }

        private static RectangleF Map(ApertureRectangle rectangle, double extent)
        {
            PointF topLeft = Map(rectangle.X1, rectangle.Y2, extent), bottomRight = Map(rectangle.X2, rectangle.Y1, extent);
            return RectangleF.FromLTRB(topLeft.X, topLeft.Y, bottomRight.X, bottomRight.Y);
        }

        private static void DrawRectangle(Graphics g, Pen pen, RectangleF rectangle)
        {
            g.DrawRectangle(pen, rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
        }

        private static double GeometryExtent(ReviewControlPointSample cp, bool available)
        {
            if (!available) return 200;
            var distances = new List<double> { 160 };
            foreach (var layer in cp.Aperture.Layers)
            {
                distances.AddRange(layer.LeafBoundariesMm.Select(Math.Abs));
                distances.AddRange(layer.Bank1PositionsMm.Select(Math.Abs));
                distances.AddRange(layer.Bank2PositionsMm.Select(Math.Abs));
            }
            foreach (var rectangle in new[] { cp.Aperture.Jaws, cp.Aperture.FixedBoundingBox }.Where(r => r != null))
                distances.AddRange(new[] { Math.Abs(rectangle.X1), Math.Abs(rectangle.X2), Math.Abs(rectangle.Y1), Math.Abs(rectangle.Y2) });
            return Math.Min(2000, Math.Ceiling(distances.Max() * 1.12 / 10) * 10);
        }

        private static Color LayerColor(int index)
        {
            return index == 0 ? Cyan : index == 1 ? Amber : Color.FromArgb(179, 158, 235);
        }

        private static bool MatchingAngle(double first, double second)
        {
            if (!Finite(first) || !Finite(second)) return false;
            double difference = Math.Abs((first % 360 - second % 360) % 360);
            return Math.Min(difference, 360 - difference) <= 0.05;
        }

        private static bool ValidRectangle(ApertureRectangle rectangle, bool positiveArea)
        {
            return rectangle != null && Physical(rectangle.X1) && Physical(rectangle.Y1) && Physical(rectangle.X2) && Physical(rectangle.Y2) &&
                (positiveArea ? rectangle.X2 > rectangle.X1 && rectangle.Y2 > rectangle.Y1 : rectangle.X2 >= rectangle.X1 && rectangle.Y2 >= rectangle.Y1);
        }

        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static bool Physical(double value) { return Finite(value) && Math.Abs(value) <= 2000; }
        private static float Clamp(float value, float low, float high) { return Math.Max(low, Math.Min(high, value)); }
        private static string Cm(double millimetres) { return (millimetres / 10).ToString("0.0", CultureInfo.InvariantCulture) + " cm"; }
        private static string Value(double? value, string format, string unit)
        {
            return value.HasValue && Finite(value.Value) ? value.Value.ToString(format, CultureInfo.InvariantCulture) + unit : "Unavailable";
        }

        private static string Clean(string value, string fallback, int limit)
        {
            if (string.IsNullOrWhiteSpace(value)) return fallback;
            string clean = new string(value.Where(c => !char.IsControl(c)).ToArray()).Trim();
            return clean.Length > limit ? clean.Substring(0, limit - 1) + "…" : clean;
        }
    }
}
