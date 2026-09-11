using System;
using System.Linq;
using System.Reflection;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class PlanImageViewportTests
    {
        public static void RunAll()
        {
            var type = typeof(ReviewPlanImage).Assembly.GetType("ClearPlan.Core.Review.PlanImageViewport");
            TestAssert.NotNull(type, "Missing iso-centered, dose-bounded viewport.");
            var image = new ReviewPlanImage { WidthPixels = 101, HeightPixels = 81,
                PixelSpacingXMillimeters = 2, PixelSpacingYMillimeters = 3,
                IsocenterPixelX = 30.25, IsocenterPixelY = 45.75 };
            var samples = new double[25];
            samples[2 * 5 + 1] = 1; samples[2 * 5 + 2] = 2;
            // 2% of 50 Gy, not 20%; threshold equality is part of the display region.
            object region = type.GetMethod("CaptureDoseRegion").Invoke(null, new object[] { samples, 5, 5, 1.0, 101, 81 });
            TestAssert.NotNull(region);
            typeof(ReviewPlanImage).GetProperty("DoseFocusRegion").SetValue(image, region, null);
            dynamic viewport = type.GetMethod("Calculate").Invoke(null, new object[] { image });
            Close(0.5, (double)viewport.MapX(30.25)); Close(0.5, (double)viewport.MapY(45.75));
            Close((double)viewport.MapX(31.75) - (double)viewport.MapX(30.25),
                (double)viewport.MapY(46.75) - (double)viewport.MapY(45.75));
            TestAssert.True(viewport.UsesDoseRegion);
            TestAssert.True(viewport.Description.Contains("2%") && viewport.Description.Contains("10 mm"));
            // Bounds cover whole sampled cells, including region interiors, not just traced contours.
            dynamic extent = region;
            Close(0, (double)extent.MinPixelX); Close(75, (double)extent.MaxPixelX);
            Close(20, (double)extent.MinPixelY); Close(60, (double)extent.MaxPixelY);
            TestAssert.True(viewport.MapX(0.0) > 0 && viewport.MapX(75.0) < 1);
            TestAssert.True(viewport.MapY(20.0) > 0 && viewport.MapY(60.0) < 1);

            // Boundary/unknown coverage must be disclosed, never interpreted as zero dose.
            samples = Enumerable.Repeat(2.0, 25).ToArray(); samples[0] = double.NaN;
            dynamic edge = type.GetMethod("CaptureDoseRegion").Invoke(null, new object[] { samples, 5, 5, 1.0, 101, 81 });
            TestAssert.True(edge.CoverageLimited);
            TestAssert.True(type.GetMethod("CaptureDoseRegion").Invoke(null,
                new object[] { Enumerable.Repeat(double.NaN, 25).ToArray(), 5, 5, 1.0, 101, 81 }) == null);

            typeof(ReviewPlanImage).GetProperty("DoseFocusRegion").SetValue(image, null, null);
            viewport = type.GetMethod("Calculate").Invoke(null, new object[] { image });
            TestAssert.False(viewport.UsesDoseRegion);
            Close(0.5, (double)viewport.MapX(30.25)); Close(0.5, (double)viewport.MapY(45.75));
            TestAssert.True(viewport.MapX(-0.5) >= 0 && viewport.MapX(100.5) <= 1);
            TestAssert.True(viewport.MapY(-0.5) >= 0 && viewport.MapY(80.5) <= 1);
            image.IsocenterPixelX = double.NaN;
            viewport = type.GetMethod("Calculate").Invoke(null, new object[] { image });
            TestAssert.False(viewport.HasIsocenter);
            TestAssert.True(viewport.Description.Contains("unavailable"));

            // Floating-point resampling at the last source pixel must not invalidate its own bounds.
            image.WidthPixels = 384; image.HeightPixels = 256;
            image.IsocenterPixelX = 180; image.IsocenterPixelY = 120;
            region = type.GetMethod("CaptureDoseRegion").Invoke(null,
                new object[] { Enumerable.Repeat(2.0, 192 * 192).ToArray(), 192, 192, 1.0, 384, 256 });
            typeof(ReviewPlanImage).GetProperty("DoseFocusRegion").SetValue(image, region, null);
            viewport = type.GetMethod("Calculate").Invoke(null, new object[] { image });
            TestAssert.True(viewport.UsesDoseRegion, "Round-off at the source edge must not discard valid dose-region metadata.");
            TestAssert.True(viewport.Description.Contains("truncated"));
        }

        private static void Close(double expected, double actual)
        { TestAssert.True(Math.Abs(expected - actual) < 1e-9, "Viewport transform/extent mismatch."); }
    }
}
