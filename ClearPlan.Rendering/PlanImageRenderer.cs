using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using ClearPlan.Core.Review;

namespace ClearPlan.Rendering
{
    /// <summary>Display-only CT renderer shared by the native workspace and reports. Input is detached, already-windowed data.</summary>
    public static class PlanImageRenderer
    {
        private static readonly Color Background = Color.FromArgb(13, 24, 40);
        private static readonly Color Orientation = Color.FromArgb(75, 220, 195);

        public static IList<ReviewImageOverlay> IsodoseScale(ReviewPlanImage image)
        {
            return (image.Overlays ?? new List<ReviewImageOverlay>()).Where(o => o != null && o.Kind == "isodose" &&
                o.SourceStatus == ReviewStatusCodes.Available && o.DoseGy.HasValue && o.Paths != null &&
                o.Paths.Any(p => p != null && p.Points != null && p.Points.Count >= 2))
                .OrderBy(o => o.DoseGy.Value).ToList().AsReadOnly();
        }
        public static int CanvasHeight(ReviewPlanImage image, bool showDose)
        {
            int count = showDose ? IsodoseScale(image).Count : 0;
            return 720 + (count == 0 ? 0 : 26 + ((count + 4) / 5) * 48) + (image.Synthetic ? 32 : 0);
        }

        public static bool IsRenderable(ReviewPlanImage image)
        {
            return image != null && image.SourceStatus == ReviewStatusCodes.Available &&
                image.WidthPixels >= 2 && image.HeightPixels >= 2 && image.WidthPixels <= 2048 && image.HeightPixels <= 2048 &&
                image.GrayscalePixels != null && image.GrayscalePixels.Length == (long)image.WidthPixels * image.HeightPixels &&
                Positive(image.PixelSpacingXMillimeters) && Positive(image.PixelSpacingYMillimeters);
        }

        public static PlanImageViewport CalculateViewport(ReviewPlanImage image, bool focusIsocenter)
        {
            if (focusIsocenter) return PlanImageViewport.Calculate(image);
            if (image == null) throw new ArgumentException("Missing detached CT geometry.");
            // Fit uses the same physical-coordinate implementation with image-centered geometry.
            // Never remove the source isocenter or dose region from the detached snapshot.
            return PlanImageViewport.Calculate(new ReviewPlanImage
            {
                WidthPixels = image.WidthPixels, HeightPixels = image.HeightPixels,
                PixelSpacingXMillimeters = image.PixelSpacingXMillimeters,
                PixelSpacingYMillimeters = image.PixelSpacingYMillimeters
            });
        }

        public static byte[] Render(ReviewPlanImage image, bool showStructures, bool showDose, bool focusIsocenter)
        { return RenderFiltered(image, showStructures, showDose, focusIsocenter, null); }

        public static byte[] RenderFiltered(ReviewPlanImage image, bool showStructures, bool showDose, bool focusIsocenter, IEnumerable<string> hiddenStructureIds)
        { return RenderFiltered(image, showStructures, showDose, focusIsocenter, hiddenStructureIds, true); }

        public static byte[] RenderFiltered(ReviewPlanImage image, bool showStructures, bool showDose, bool focusIsocenter, IEnumerable<string> hiddenStructureIds, bool includeRasterScale)
        {
            var hidden = new HashSet<string>(hiddenStructureIds ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            if (!IsRenderable(image)) throw new ArgumentException("Invalid or unavailable detached overview image.");
            int canvasHeight = CanvasHeight(image, showDose && includeRasterScale);
            using (var output = new Bitmap(1440, canvasHeight * 2))
            using (var graphics = Graphics.FromImage(output))
            using (var pixels = new Bitmap(image.WidthPixels, image.HeightPixels, PixelFormat.Format24bppRgb))
            {
                graphics.Clear(Background);
                graphics.ScaleTransform(2, 2);
                BitmapData data = pixels.LockBits(new Rectangle(0, 0, pixels.Width, pixels.Height), ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);
                try
                {
                    var buffer = new byte[data.Stride * data.Height];
                    for (int y = 0; y < pixels.Height; y++)
                    for (int x = 0; x < pixels.Width; x++)
                    {
                        byte value = image.GrayscalePixels[y * pixels.Width + x];
                        int offset = y * data.Stride + x * 3;
                        buffer[offset] = buffer[offset + 1] = buffer[offset + 2] = value;
                    }
                    System.Runtime.InteropServices.Marshal.Copy(buffer, 0, data.Scan0, buffer.Length);
                }
                finally { pixels.UnlockBits(data); }
                var viewport = CalculateViewport(image, focusIsocenter);
                var sourceViewport = PlanImageViewport.Calculate(image);
                var view = new RectangleF(30, 30, 660, 660);
                Func<double, float> mapX = x => view.Left + (float)viewport.MapX(x) * view.Width;
                Func<double, float> mapY = y => view.Top + (float)viewport.MapY(y) * view.Height;
                var target = RectangleF.FromLTRB(mapX(-0.5), mapY(-0.5),
                    mapX(image.WidthPixels - 0.5), mapY(image.HeightPixels - 0.5));
                var state = graphics.Save();
                graphics.SetClip(view);
                graphics.InterpolationMode = InterpolationMode.NearestNeighbor;
                graphics.PixelOffsetMode = PixelOffsetMode.Half;
                graphics.DrawImage(pixels, target);
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                DrawOverlays(graphics, image, target, showStructures, showDose, hidden);
                if (sourceViewport.HasIsocenter)
                    using (var pen = new Pen(Orientation, 2.0f))
                    {
                        float x = mapX(image.IsocenterPixelX.Value), y = mapY(image.IsocenterPixelY.Value);
                        graphics.DrawLine(pen, x - 22, y, x + 22, y);
                        graphics.DrawLine(pen, x, y - 22, x, y + 22);
                    }
                graphics.Restore(state);
                using (var font = new Font("Segoe UI", 21, FontStyle.Bold))
                using (var brush = new SolidBrush(Orientation))
                {
                    Text(graphics, image.LeftOrientation, font, brush, 15, 360);
                    Text(graphics, image.RightOrientation, font, brush, 705, 360);
                    Text(graphics, image.TopOrientation, font, brush, 360, 15);
                    Text(graphics, image.BottomOrientation, font, brush, 360, 703);
                }
                if (showDose && includeRasterScale) DrawIsodoseScale(graphics, image);
                if (image.Synthetic)
                {
                    using (var fill = new SolidBrush(Color.FromArgb(190, 123, 20, 31)))
                        graphics.FillRectangle(fill, 0, canvasHeight - 29, 720, 29);
                    using (var font = new Font("Segoe UI", 14, FontStyle.Bold))
                        Text(graphics, "SIMULATION - NOT PATIENT DATA", font, Brushes.White, 360, canvasHeight - 15);
                }
                using (var stream = new MemoryStream())
                { output.Save(stream, ImageFormat.Png); return stream.ToArray(); }
            }
        }

        private static void DrawIsodoseScale(Graphics graphics, ReviewPlanImage image)
        {
            var scale = IsodoseScale(image);
            if (scale.Count == 0) return;
            using (var heading = new Font("Segoe UI", 12, FontStyle.Bold))
            using (var value = new Font("Segoe UI", 13, FontStyle.Regular))
            using (var doseFont = new Font("Segoe UI", 11, FontStyle.Regular))
            using (var muted = new SolidBrush(Color.FromArgb(198, 213, 226)))
            {
                graphics.DrawString("Isodoses  |  % Rx / Gy", heading, muted, 30, 719);
                for (int i = 0; i < scale.Count; i++)
                {
                    var level = scale[i]; float x = 30 + (i % 5) * 132, y = 744 + (i / 5) * 48;
                    Color color; try { color = ColorTranslator.FromHtml(level.ColorHex); } catch (Exception) { color = Color.White; }
                    using (var brush = new SolidBrush(color)) graphics.FillRectangle(brush, x, y, 120, 5);
                    double? percent = IsodoseDisplayConfiguration.PercentFor(level);
                    graphics.DrawString(percent.HasValue ? percent.Value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + "% Rx" : "Dose", value, Brushes.White, x, y + 7);
                    graphics.DrawString(level.DoseGy.Value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + " Gy", doseFont, muted, x, y + 25);
                }
            }
        }

        private static void DrawOverlays(Graphics graphics, ReviewPlanImage image, RectangleF target, bool showStructures, bool showDose, ISet<string> hidden)
        {
            var state = graphics.Save();
            try
            {
                graphics.SetClip(target, CombineMode.Intersect);
                foreach (var overlay in (image.Overlays ?? new List<ReviewImageOverlay>())
                    .Where(item => item != null && item.SourceStatus == ReviewStatusCodes.Available &&
                        ((showStructures && item.Kind == "structure" && !hidden.Contains(item.Label)) || (showDose && item.Kind == "isodose")))
                    .OrderBy(item => item.Kind == "isodose" ? 0 : 1))
                {
                    Color color;
                    try { color = ColorTranslator.FromHtml(overlay.ColorHex); }
                    catch (Exception) { color = Color.White; }
                    using (var pen = new Pen(color, overlay.Kind == "isodose" ? 1.8f : 0.9f))
                    {
                        pen.DashStyle = DashStyle.Solid;
                        foreach (var path in overlay.Paths ?? new List<ReviewImagePath>())
                        {
                            if (path == null || path.Points == null || path.Points.Count < 2 || path.Points.Count > 200000 ||
                                path.Points.Any(point => point == null || !Finite(point.X) || !Finite(point.Y))) continue;
                            var points = path.Points.Select(point => new PointF(
                                target.Left + (float)((point.X + 0.5) / image.WidthPixels * target.Width),
                                target.Top + (float)((point.Y + 0.5) / image.HeightPixels * target.Height))).ToArray();
                            if (path.Closed && points.Length > 2) graphics.DrawPolygon(pen, points);
                            else graphics.DrawLines(pen, points);
                        }
                    }
                }
            }
            finally { graphics.Restore(state); }
        }

        private static void Text(Graphics graphics, string value, Font font, Brush brush, float x, float y)
        {
            using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                graphics.DrawString(value ?? "", font, brush, x, y, format);
        }
        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static bool Positive(double value) { return Finite(value) && value > 0; }
    }
}
