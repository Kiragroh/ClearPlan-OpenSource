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
using ClearPlan.Core.Localization;

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
        public string[] JawCoordinateLabels { get; internal set; }
        public bool UprightDisplay { get; internal set; }
        public double DisplayRotationDegrees { get; internal set; }
        public double DisplayExtentMm { get; internal set; }
        public string FrameDescription { get; internal set; }
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
        private static readonly Color ApertureGreen = Color.FromArgb(75, 245, 105);
        private static readonly Color JawColor = Color.FromArgb(255, 230, 96);
        private static readonly Color SimulationRed = Color.FromArgb(180, 35, 24);
        private const float LogicalWidth = 1500;
        private const float LogicalHeight = 1000;
        private static readonly RectangleF Viewport = new RectangleF(50, 54, 900, 900);

        /// <summary>Captured aperture in BLD coordinates only. No DRR, patient orientation
        /// or interpolation is implied; the host supplies the CP and collimator labels.</summary>
        public static byte[] RenderApertureInset(ReviewBeamAnalysis beam, ReviewControlPointSample cp, int size = 360)
        {
            if (size < 160 || size > 1024) throw new ArgumentOutOfRangeException(nameof(size));
            bool available = GeometryProblem(beam, cp) == null;
            using (var bitmap = new Bitmap(size, size, PixelFormat.Format24bppRgb))
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Canvas); g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                g.ScaleTransform(size / 940f, size / 940f);
                g.TranslateTransform(20 - Viewport.Left, 20 - Viewport.Top);
                g.SetClip(Viewport);
                if (available) DrawAperture(g, cp.Aperture, GeometryExtent(cp, true), true);
                else Text(g, "MLC unavailable", new RectangleF(70,400,860,100), 44, Color.White, false, StringAlignment.Center);
                var center = Map(0,0,1);
                using(var pen = new Pen(Color.White,3))
                {
                    g.DrawLine(pen,center.X-16,center.Y,center.X+16,center.Y);
                    g.DrawLine(pen,center.X,center.Y-16,center.X,center.Y+16);
                }
                Text(g,"+Y",new RectangleF(470,65,80,44),32,Color.White);
                Text(g,"+X",new RectangleF(852,465,88,44),32,Color.White);
                using(var stream = new MemoryStream()) { bitmap.Save(stream,ImageFormat.Png); return stream.ToArray(); }
            }
        }

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
                Text(graphics, state.FrameDescription, new RectangleF(367, 17, 583, 30), 20, Color.FromArgb(192, 205, 217), false, StringAlignment.Far);

                double extent = state.DisplayExtentMm;
                Fill(graphics, Color.FromArgb(25, 32, 40), Viewport);
                var clipping = graphics.Save();
                graphics.SetClip(Viewport);
                var bldTransform = graphics.Save();
                RotateAboutIso(graphics, state.DisplayRotationDegrees);
                if (state.ImageAvailable)
                {
                    using (var pixels = GrayscaleBitmap(image))
                    {
                        graphics.InterpolationMode = InterpolationMode.HighQualityBilinear;
                        var destination = Map(new ApertureRectangle(-image.ExtentMm, -image.ExtentMm, image.ExtentMm, image.ExtentMm), extent);
                        graphics.DrawImage(pixels, destination);
                    }
                }
                if (state.GeometryAvailable) DrawAperture(graphics, cp.Aperture, extent);
                graphics.Restore(bldTransform);
                DrawRuler(graphics, extent);
                graphics.Restore(clipping);
                using (var pen = new Pen(Color.FromArgb(104, 121, 136), 1)) graphics.DrawRectangle(pen, Viewport.X, Viewport.Y, Viewport.Width, Viewport.Height);

                // The native viewer exposes diagnostics in its accessible status control.
                // Reports keep the full in-image notice by default.
                if (!state.ImageAvailable && !compactStatus) DrawUnavailable(graphics, state.ImageMessage);
                if (state.GeometryAvailable) DrawApertureLegend(graphics, cp.Aperture, state);
                if (state.PhysicalJawsVisible) DrawJawLabels(graphics, cp.Aperture.Jaws, extent, state.DisplayRotationDegrees);
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
                JawCoordinateLabels = new string[0],
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
                state.JawCoordinateLabels = state.PhysicalJawsVisible ? JawCoordinateLabels(cp.Aperture.Jaws) : new string[0];
                state.LayoutDescription = (state.PhysicalJawsVisible ? "Physical jaws" : "Jawless") + " · " + cp.Aperture.Layers.Count + " MLC " + (cp.Aperture.Layers.Count == 1 ? "layer" : "layers");
                if (state.FixedBoundsVisible) state.LayoutDescription += " · maximum " + FieldSize(cp.Aperture.FixedBoundingBox);
                state.GeometryMessage = state.LayoutDescription;
            }
            state.UprightDisplay = state.ImageAvailable && image.BldToDisplayRotationDegrees.HasValue && Finite(image.BldToDisplayRotationDegrees.Value);
            state.DisplayRotationDegrees = state.UprightDisplay ? Math.IEEERemainder(image.BldToDisplayRotationDegrees.Value, 360) : 0;
            state.FrameDescription = state.UprightDisplay ? "Upright gantry (C0) · collimator " + Value(cp.CollimatorAngleDegrees, "0.#", "°") : "BLD frame · upright calibration unavailable";
            double extent = state.ImageAvailable ? image.ExtentMm : GeometryExtent(cp, state.GeometryAvailable);
            if (state.GeometryAvailable) extent = Math.Max(extent, GeometryExtent(cp, true));
            double radians = state.DisplayRotationDegrees * Math.PI / 180;
            state.DisplayExtentMm = extent * (Math.Abs(Math.Cos(radians)) + Math.Abs(Math.Sin(radians)));
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

        private static void DrawAperture(Graphics g, ApertureGeometry aperture, double extent, bool compact = false)
        {
            for (int layerIndex = 0; layerIndex < aperture.Layers.Count; layerIndex++)
            {
                var layer = aperture.Layers[layerIndex];
                Color color = LayerColor(layerIndex);
                using (var outline = new Pen(Color.FromArgb(190, color), compact ? 3f : 1.2f))
                using (var tint = new SolidBrush(Color.FromArgb(compact ? 22 : 8, color)))
                {
                    if (layerIndex % 2 != 0) outline.DashStyle = DashStyle.Dash;
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
            DrawEffectiveBoundary(g, aperture.EffectiveOpenings, extent);
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

        private static void DrawEffectiveBoundary(Graphics g, IList<ApertureRectangle> openings, double extent)
        {
            // Effective openings are non-overlapping cells from the aperture calculation. Cancel
            // coincident opposing edges, including partially shared edges, instead of drawing seams.
            var horizontal = new Dictionary<double, SortedDictionary<double, int>>();
            var vertical = new Dictionary<double, SortedDictionary<double, int>>();
            foreach (var rectangle in openings.Where(r => r.X2 > r.X1 && r.Y2 > r.Y1))
            {
                AddBoundaryEdge(horizontal, rectangle.Y1, rectangle.X1, rectangle.X2, 1);
                AddBoundaryEdge(horizontal, rectangle.Y2, rectangle.X1, rectangle.X2, -1);
                AddBoundaryEdge(vertical, rectangle.X1, rectangle.Y1, rectangle.Y2, 1);
                AddBoundaryEdge(vertical, rectangle.X2, rectangle.Y1, rectangle.Y2, -1);
            }
            using (var pen = new Pen(ApertureGreen, 2.8f))
            {
                DrawBoundaryEdges(g, horizontal, extent, pen, true);
                DrawBoundaryEdges(g, vertical, extent, pen, false);
            }
        }

        private static void AddBoundaryEdge(Dictionary<double, SortedDictionary<double, int>> edges, double fixedCoordinate, double low, double high, int sign)
        {
            SortedDictionary<double, int> events;
            if (!edges.TryGetValue(fixedCoordinate, out events)) edges[fixedCoordinate] = events = new SortedDictionary<double, int>();
            if (!events.ContainsKey(low)) events[low] = 0;
            if (!events.ContainsKey(high)) events[high] = 0;
            events[low] += sign; events[high] -= sign;
        }

        private static void DrawBoundaryEdges(Graphics g, Dictionary<double, SortedDictionary<double, int>> edges, double extent, Pen pen, bool horizontal)
        {
            foreach (var edge in edges)
            {
                double previous = 0; int active = 0;
                foreach (var change in edge.Value)
                {
                    if (active != 0)
                        g.DrawLine(pen, horizontal ? Map(previous, edge.Key, extent) : Map(edge.Key, previous, extent),
                            horizontal ? Map(change.Key, edge.Key, extent) : Map(edge.Key, change.Key, extent));
                    previous = change.Key; active += change.Value;
                }
            }
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
                }
                g.DrawEllipse(pen, centre.X - 7, centre.Y - 7, 14, 14);
            }
            OverlayLabel(g, "+Y", new RectangleF(centre.X + 12, Viewport.Top + 8, 60, 30), Color.FromArgb(244, 241, 214), 21);
            OverlayLabel(g, "+X", new RectangleF(Viewport.Right - 67, centre.Y - 49, 60, 30), Color.FromArgb(244, 241, 214), 21);
            OverlayLabel(g, "ISO", new RectangleF(centre.X+14,centre.Y-44,60,30),Color.White,21);
            OverlayLabel(g, "Ticks: 1 cm", new RectangleF(Viewport.Right-175,Viewport.Top+16,157,30),Color.White,20);
        }

        private static void DrawJawLabels(Graphics g, ApertureRectangle jaws, double extent, double rotationDegrees)
        {
            double midX = (jaws.X1 + jaws.X2) / 2, midY = (jaws.Y1 + jaws.Y2) / 2;
            var midpoints = new[] { Map(jaws.X1, midY, extent), Map(jaws.X2, midY, extent), Map(midX, jaws.Y1, extent), Map(midX, jaws.Y2, extent) };
            // Unit screen normals locate labels outside the actual rotated jaw, while text stays horizontal.
            var normals = new[] { new PointF(-1, 0), new PointF(1, 0), new PointF(0, 1), new PointF(0, -1) };
            var labels = JawCoordinateLabels(jaws);
            double radians = rotationDegrees * Math.PI / 180;
            for (int i = 0; i < midpoints.Length; i++)
            {
                PointF point = RotatePoint(midpoints[i], rotationDegrees);
                float nx = (float)(normals[i].X * Math.Cos(radians) - normals[i].Y * Math.Sin(radians));
                float ny = (float)(normals[i].X * Math.Sin(radians) + normals[i].Y * Math.Cos(radians));
                float x = Clamp(point.X - 80 + nx * 92, Viewport.Left + 8, Viewport.Right - 168);
                float y = Clamp(point.Y - 17 + ny * 32, Viewport.Top + 140, Viewport.Bottom - 80);
                OverlayLabel(g, labels[i], new RectangleF(x, y, 160, 34), JawColor, 22);
            }
        }

        private static string[] JawCoordinateLabels(ApertureRectangle jaws)
        {
            return new[] { "X1 " + Cm(jaws.X1), "X2 " + Cm(jaws.X2), "Y1 " + Cm(jaws.Y1), "Y2 " + Cm(jaws.Y2) };
        }

        private static void DrawApertureLegend(Graphics g, ApertureGeometry aperture, BeamEyeViewRenderState state)
        {
            float y = 72;
            for (int i = 0; i < aperture.Layers.Count; i++)
            {
                string label = Clean(aperture.Layers[i].Label, "MLC layer " + (i + 1), 40) + " · " + aperture.Layers[i].Bank1PositionsMm.Length + " pairs · " + (i % 2 == 0 ? "solid" : "dashed");
                OverlayLabel(g, label, new RectangleF(68, y, 438, 34), LayerColor(i), 22);
                OverlayLabel(g, LeafWidthLabel(aperture.Layers[i]),new RectangleF(68,y+33,438,30),LayerColor(i),20);
                y += 68;
            }
            OverlayLabel(g, "Bright green: effective aperture", new RectangleF(68, 836, 680, 32), ApertureGreen, 21);
            string boundary = state.PhysicalJawsVisible ? "Solid yellow: physical jaws" : (state.FixedBoundsVisible ? "Dashed white: maximum field " + FieldSize(aperture.FixedBoundingBox) + " · fixed" : "Jawless · maximum field not supplied");
            OverlayLabel(g, boundary, new RectangleF(68, 874, 680, 34), state.PhysicalJawsVisible ? JawColor : Color.White, 22);
            OverlayLabel(g, "CT proxy, not diagnostic. Leaf bodies extend to viewport edge.", new RectangleF(68, 913, 864, 29), Color.FromArgb(213, 222, 230), 20);
        }

        private static void DrawUnavailable(Graphics g, string reason)
        {
            // Keep the isocenter and central aperture inspectable even when there is no CT projection.
            Fill(g, Color.FromArgb(228, 23, 33, 43), new RectangleF(160, 705, 680, 153));
            Text(g, "CT projection unavailable", new RectangleF(182, 718, 636, 38), 28, Color.White, true, StringAlignment.Center);
            Text(g, Clean(ReviewLanguage.Text(reason), "No CT projection supplied.", 240), new RectangleF(182, 764, 636, 83), 21, Color.FromArgb(217, 227, 236), false, StringAlignment.Center);
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
                // Exact diagnostic translation only; never translate machine/beam IDs or whole drawing strings.
                string detail = !state.ImageAvailable ? ReviewLanguage.Text(state.ImageMessage) : (!state.GeometryAvailable ? ReviewLanguage.Text(state.GeometryMessage) : Clean(ReviewLanguage.Text(image == null ? null : image.ProjectionDescription), "Detached CT-derived line-integral projection.", 320));
                Text(g, detail, new RectangleF(x, 710, w, 55), 20, Muted);
            }
            Text(g, state.GeometryAvailable ? "Complete MLC / boundary geometry" : "MLC / boundary geometry unavailable", new RectangleF(x, compactStatus ? 640 : 675, w, 28), 21, state.GeometryAvailable ? Ink : Color.FromArgb(151, 88, 19));
            DrawOrientation(g, cp, state);
        }

        private static void DrawOrientation(Graphics g, ReviewControlPointSample cp, BeamEyeViewRenderState state)
        {
            const float gantryX = 1103, couchX = 1250, collimatorX = 1397, centreY = 861;
            Text(g, "Orientation schematic", new RectangleF(1030, 780, 440, 29), 23, Ink, true);
            double gantry = cp == null ? double.NaN : cp.GantryAngleDegrees;
            double couch = cp == null ? double.NaN : cp.PatientSupportAngleDegrees;
            double collimator = cp == null ? double.NaN : cp.CollimatorAngleDegrees;
            bool calibrated = state != null && state.UprightDisplay;
            DrawOrientationReference(g, gantryX, centreY, "0°");
            DrawOrientationReference(g, couchX, centreY, "0°");
            DrawOrientationReference(g, collimatorX, centreY, calibrated ? "Up" : "0°");
            if (Finite(gantry)) DrawGantryOrientation(g, gantryX, centreY, NormalizeAngle(gantry));
            if (Finite(couch)) DrawCouchOrientation(g, couchX, centreY, NormalizeAngle(couch));
            if (Finite(collimator)) DrawCollimatorOrientation(g, collimatorX, centreY, calibrated ? state.DisplayRotationDegrees : NormalizeAngle(collimator));
            DrawOrientationCaption(g, gantryX, "Gantry", gantry, "Room · front");
            DrawOrientationCaption(g, couchX, "Couch", couch, "Table · top");
            DrawOrientationCaption(g, collimatorX, "Collimator", collimator, calibrated ? "BLD · screen" : "BLD · nominal");
            // Only InspectState validates the image frame used by both aperture and inset.
            // Otherwise disclose the nominal convention; never infer a patient transform.
            Text(g, "Orientation only · not a collision check.", new RectangleF(1030, 946, 440, 20), 16, Muted, false, StringAlignment.Center);
        }

        private static void DrawOrientationReference(Graphics g, float centreX, float centreY, string reference)
        {
            using (var pen = new Pen(Border, 1.5f)) g.DrawEllipse(pen, centreX - 30, centreY - 30, 60, 60);
            using (var pen = new Pen(Muted, 1.5f)) g.DrawLine(pen, centreX, centreY - 43, centreX, centreY - 35);
            Text(g, reference, new RectangleF(centreX + 33, 815, 25, 19), 14, Muted);
        }

        private static void DrawGantryOrientation(Graphics g, float centreX, float centreY, double angle)
        {
            PointF source = LinacOrientationPoint(angle);
            using (var pen = new Pen(Color.FromArgb(119, 143, 156), 1))
            {
                g.DrawLine(pen, centreX - 34, centreY, centreX + 34, centreY);
                g.DrawLine(pen, centreX, centreY - 30, centreX, centreY + 30);
            }
            using (var pen = new Pen(Teal, 2.5f)) g.DrawLine(pen, centreX, centreY, centreX + source.X * 27, centreY + source.Y * 27);
            var transform = g.Save();
            g.TranslateTransform(centreX + source.X * 27, centreY + source.Y * 27);
            g.RotateTransform((float)angle);
            Fill(g, Color.FromArgb(22, 50, 74), new RectangleF(-10, -7, 20, 14));
            Fill(g, Teal, new RectangleF(-6, 5, 12, 4));
            g.Restore(transform);
            using (var brush = new SolidBrush(Teal)) g.FillEllipse(brush, centreX - 3, centreY - 3, 6, 6);
        }

        private static void DrawCouchOrientation(Graphics g, float centreX, float centreY, double angle)
        {
            // Separate top view: couch rotates the table and patient symbol, never the gantry.
            var transform = g.Save();
            g.TranslateTransform(centreX, centreY);
            g.RotateTransform((float)angle);
            Fill(g, Color.FromArgb(224, 232, 237), new RectangleF(-12, -28, 24, 56));
            using (var pen = new Pen(Color.FromArgb(117, 134, 148), 1.5f)) g.DrawRectangle(pen, -12, -28, 24, 56);
            using (var brush = new SolidBrush(Teal))
            {
                g.FillEllipse(brush, -4, -21, 8, 8);
                g.FillRectangle(brush, -4, -10, 8, 24);
            }
            g.Restore(transform);
            using (var pen = new Pen(Ink, 1))
            {
                g.DrawLine(pen, centreX - 5, centreY, centreX + 5, centreY);
                g.DrawLine(pen, centreX, centreY - 5, centreX, centreY + 5);
            }
        }

        private static void DrawCollimatorOrientation(Graphics g, float centreX, float centreY, double angle)
        {
            // Directed, labelled axes avoid the 180-degree ambiguity of a square/leaf icon.
            // The input is the validated BLD-to-screen rotation when available; otherwise
            // nominal C0 starts +X right / +Y above with positive rotation clockwise.
            // This is a direction indicator, not an inferred drawing of actual leaves.
            double radians = angle * Math.PI / 180;
            var x = new PointF((float)Math.Cos(radians), (float)Math.Sin(radians));
            var y = new PointF(x.Y, -x.X);
            using (var cap = new AdjustableArrowCap(3, 4))
            using (var xPen = new Pen(Teal, 2.3f))
            using (var yPen = new Pen(Ink, 1.7f))
            {
                xPen.CustomEndCap = cap;
                yPen.CustomEndCap = cap;
                yPen.DashStyle = DashStyle.Dash;
                g.DrawLine(xPen, centreX, centreY, centreX + x.X * 19, centreY + x.Y * 19);
                g.DrawLine(yPen, centreX, centreY, centreX + y.X * 19, centreY + y.Y * 19);
            }
            Text(g, "+X", new RectangleF(centreX + x.X * 32 - 13, centreY + x.Y * 32 - 9, 26, 19), 14, Teal, true, StringAlignment.Center);
            Text(g, "+Y", new RectangleF(centreX + y.X * 32 - 13, centreY + y.Y * 32 - 9, 26, 19), 14, Ink, true, StringAlignment.Center);
        }

        private static void DrawOrientationCaption(Graphics g, float centreX, string name, double angle, string view)
        {
            bool available = Finite(angle);
            string value = available ? NormalizeAngle(angle).ToString("0.#", CultureInfo.InvariantCulture) + "°" : "—";
            Text(g, name + " " + value, new RectangleF(centreX - 73, 906, 146, 24), 16, Ink, true, StringAlignment.Center);
            Text(g, available ? view : "Angle unavailable", new RectangleF(centreX - 73, 929, 146, 20), 16, Muted, false, StringAlignment.Center);
        }

        private static double NormalizeAngle(double angle) { return ((angle % 360) + 360) % 360; }

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

        private static void RotateAboutIso(Graphics g, double rotationDegrees)
        {
            // CreateFrame's X=u*cos-v*sin, Y=u*sin+v*cos requires +theta in GDI's Y-down
            // screen coordinates. The same transform applies to raster, leaves and all boundaries.
            var centre = Map(0, 0, 1);
            g.TranslateTransform(centre.X, centre.Y);
            g.RotateTransform((float)rotationDegrees);
            g.TranslateTransform(-centre.X, -centre.Y);
        }

        private static PointF RotatePoint(PointF point, double rotationDegrees)
        {
            var centre = Map(0, 0, 1);
            double radians = rotationDegrees * Math.PI / 180;
            double x = point.X - centre.X, y = point.Y - centre.Y;
            return new PointF(centre.X + (float)(x * Math.Cos(radians) - y * Math.Sin(radians)),
                centre.Y + (float)(x * Math.Sin(radians) + y * Math.Cos(radians)));
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
            return index % 2 == 0 ? Color.FromArgb(75, 245, 105) : Color.FromArgb(78, 155, 255);
        }

        private static string FieldSize(ApertureRectangle field)
        {
            return ((field.X2 - field.X1) / 10).ToString("0.#", CultureInfo.InvariantCulture) + " × " +
                ((field.Y2 - field.Y1) / 10).ToString("0.#", CultureInfo.InvariantCulture) + " cm";
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
