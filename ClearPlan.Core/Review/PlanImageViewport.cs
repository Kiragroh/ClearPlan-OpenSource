using System;
using System.Globalization;

namespace ClearPlan.Core.Review
{
    /// <summary>Display-only square physical viewport; input coordinates are native overview pixel centers.</summary>
    public sealed class PlanImageViewport
    {
        public bool HasIsocenter { get; private set; }
        public bool UsesDoseRegion { get; private set; }
        public double CenterPixelX { get; private set; }
        public double CenterPixelY { get; private set; }
        public double HalfExtentMillimeters { get; private set; }
        public string Description { get; private set; }
        private double spacingX, spacingY;

        public double MapX(double pixelCenter) { return 0.5 + (pixelCenter - CenterPixelX) * spacingX / (2 * HalfExtentMillimeters); }
        public double MapY(double pixelCenter) { return 0.5 + (pixelCenter - CenterPixelY) * spacingY / (2 * HalfExtentMillimeters); }

        public static PlanImageViewport Calculate(ReviewPlanImage image)
        {
            if (image == null || image.WidthPixels < 2 || image.HeightPixels < 2 ||
                !Positive(image.PixelSpacingXMillimeters) || !Positive(image.PixelSpacingYMillimeters))
                throw new ArgumentException("Invalid image viewport geometry.");
            bool hasIso = image.IsocenterPixelX.HasValue && image.IsocenterPixelY.HasValue &&
                Finite(image.IsocenterPixelX.Value) && Finite(image.IsocenterPixelY.Value) &&
                image.IsocenterPixelX >= -0.5 && image.IsocenterPixelX <= image.WidthPixels - 0.5 &&
                image.IsocenterPixelY >= -0.5 && image.IsocenterPixelY <= image.HeightPixels - 0.5;
            var result = new PlanImageViewport { HasIsocenter = hasIso,
                CenterPixelX = hasIso ? image.IsocenterPixelX.Value : (image.WidthPixels - 1.0) / 2,
                CenterPixelY = hasIso ? image.IsocenterPixelY.Value : (image.HeightPixels - 1.0) / 2,
                spacingX = image.PixelSpacingXMillimeters, spacingY = image.PixelSpacingYMillimeters };
            var region = image.DoseFocusRegion;
            result.UsesDoseRegion = hasIso && region != null && Positive(region.ThresholdGy) &&
                region.PrescriptionPercent == 2 && Finite(region.MinPixelX) && Finite(region.MaxPixelX) &&
                Finite(region.MinPixelY) && Finite(region.MaxPixelY) && region.MinPixelX >= 0 && region.MinPixelY >= 0 &&
                region.MaxPixelX <= image.WidthPixels - 1 && region.MaxPixelY <= image.HeightPixels - 1 &&
                region.MinPixelX <= region.MaxPixelX && region.MinPixelY <= region.MaxPixelY;
            double minX = result.UsesDoseRegion ? region.MinPixelX : -0.5;
            double maxX = result.UsesDoseRegion ? region.MaxPixelX : image.WidthPixels - 0.5;
            double minY = result.UsesDoseRegion ? region.MinPixelY : -0.5;
            double maxY = result.UsesDoseRegion ? region.MaxPixelY : image.HeightPixels - 0.5;
            result.HalfExtentMillimeters = Math.Max(
                Math.Max(Math.Abs(minX - result.CenterPixelX), Math.Abs(maxX - result.CenterPixelX)) * result.spacingX,
                Math.Max(Math.Abs(minY - result.CenterPixelY), Math.Abs(maxY - result.CenterPixelY)) * result.spacingY)
                + (result.UsesDoseRegion ? 10 : 0);
            result.Description = (hasIso ? "Isocenter-centered" : "Isocenter unavailable; image-centered") +
                (result.UsesDoseRegion ? "; zoom: sampled >=2% Rx region + 10 mm. " : "; full CT (2% dose region unavailable). ") +
                string.Format(CultureInfo.InvariantCulture, "Field of view {0:0} x {0:0} mm. ", result.HalfExtentMillimeters * 2) +
                (result.UsesDoseRegion && region.CoverageLimited ? "Dose-region extent may be truncated at source coverage. " : "") +
                "Outside CT coverage stays unfilled; no extrapolation.";
            return result;
        }

        /// <summary>Bound all finite samples >= the 2%-Rx threshold, including one adjacent sampling cell.
        /// This includes filled interiors even if no closed contour can be traced at a grid edge.</summary>
        public static ReviewImageDoseRegion CaptureDoseRegion(double[] samples, int columns, int rows,
            double twoPercentDoseGy, int imageWidth, int imageHeight)
        {
            if (columns < 2 || rows < 2 || columns > 512 || rows > 512 || samples == null ||
                samples.Length != columns * rows || imageWidth < 2 || imageHeight < 2 || !Positive(twoPercentDoseGy))
                throw new ArgumentException("Invalid dose-region sampling grid.");
            int minX = columns, minY = rows, maxX = -1, maxY = -1;
            bool limited = false;
            for (int y = 0; y < rows; y++)
            for (int x = 0; x < columns; x++)
            {
                double value = samples[y * columns + x];
                if (!Finite(value) || value < twoPercentDoseGy) continue;
                minX = Math.Min(minX, x); maxX = Math.Max(maxX, x);
                minY = Math.Min(minY, y); maxY = Math.Max(maxY, y);
                if (x == 0 || y == 0 || x == columns - 1 || y == rows - 1) limited = true;
                for (int yy = Math.Max(0, y - 1); yy <= Math.Min(rows - 1, y + 1); yy++)
                for (int xx = Math.Max(0, x - 1); xx <= Math.Min(columns - 1, x + 1); xx++)
                    if (!Finite(samples[yy * columns + xx]) || samples[yy * columns + xx] < 0) limited = true;
            }
            if (maxX < 0) return null;
            double dx = (imageWidth - 1.0) / (columns - 1), dy = (imageHeight - 1.0) / (rows - 1);
            return new ReviewImageDoseRegion { PrescriptionPercent = 2, ThresholdGy = twoPercentDoseGy,
                MinPixelX = Math.Max(0, minX - 1) * dx, MaxPixelX = Math.Min(imageWidth - 1, Math.Min(columns - 1, maxX + 1) * dx),
                MinPixelY = Math.Max(0, minY - 1) * dy, MaxPixelY = Math.Min(imageHeight - 1, Math.Min(rows - 1, maxY + 1) * dy),
                CoverageLimited = limited };
        }

        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static bool Positive(double value) { return Finite(value) && value > 0; }
    }
}
