using System;
using System.Collections.Generic;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Simulation
{
    /// <summary>Deterministic mathematical ellipsoid phantom, not generated or patient anatomy.</summary>
    public static class SyntheticPlanImageFactory
    {
        public static List<ReviewPlanImage> Create(string planKey)
        {
            var result = new List<ReviewPlanImage>();
            string[] kinds = { "transversal", "coronal", "sagittal" };
            string[] titles = { "Transversal", "Coronal (orthogonal)", "Sagittal" };
            for (int view = 0; view < 3; view++)
            {
                const int size = 320;
                var image = new ReviewPlanImage
                {
                    PlanKey = planKey, Kind = kinds[view], Title = titles[view],
                    Caption = "SIMULATION - mathematical ellipsoid phantom; not patient CT. Central plane; W400 / L40 HU-equivalent.",
                    Synthetic = true, SourceStatus = ReviewStatusCodes.Available,
                    WidthPixels = size, HeightPixels = size, GrayscalePixels = new byte[size * size],
                    PixelSpacingXMillimeters = 1.5, PixelSpacingYMillimeters = 1.5,
                    LeftOrientation = view == 2 ? "A" : "R", RightOrientation = view == 2 ? "P" : "L",
                    TopOrientation = view == 0 ? "A" : "S", BottomOrientation = view == 0 ? "P" : "I",
                    IsocenterPixelX = (size - 1) / 2.0, IsocenterPixelY = (size - 1) / 2.0
                };
                for (int row = 0; row < size; row++)
                for (int column = 0; column < size; column++)
                {
                    double u = (column - (size - 1) / 2.0) * 1.5;
                    double v = (row - (size - 1) / 2.0) * 1.5;
                    double x = view == 2 ? 0 : u;
                    double y = view == 0 ? v : view == 2 ? u : 0;
                    double z = view == 0 ? 0 : -v;
                    double body = x * x / (165 * 165) + y * y / (115 * 115) + z * z / (205 * 205);
                    double value = body <= 1 ? 35 + 10 * Math.Cos(x / 30) * Math.Cos(z / 40) : -1000;
                    if (body <= 1 && body > 0.86) value = -85;
                    if (x * x / 1600 + y * y / 625 + z * z / 3600 < 1) value = 110;
                    if ((x - 80) * (x - 80) + y * y < 225 && Math.Abs(z) < 145) value = 900;
                    if ((x + 80) * (x + 80) + y * y < 225 && Math.Abs(z) < 145) value = 900;
                    image.GrayscalePixels[row * size + column] = OrthogonalImageGeometry.WindowHu(value, 40, 400);
                }
                result.Add(image);
            }
            return result;
        }
    }
}
