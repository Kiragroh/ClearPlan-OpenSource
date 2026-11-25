using System;
using System.Collections.Generic;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;
using ClearPlan.Rendering;

namespace ClearPlan.Core.Tests
{
    /// <summary>GUI-only tight frames (2026-11-20): outermost shown isodose + 10 mm for CT planes,
    /// outermost opening or jaw + 10 mm over all control points for the BEV.</summary>
    internal static class GuiTightFrameTests
    {
        public static void IsodoseFit()
        {
            var image = new ReviewPlanImage { WidthPixels = 200, HeightPixels = 200,
                PixelSpacingXMillimeters = 1, PixelSpacingYMillimeters = 2, IsocenterPixelX = 100, IsocenterPixelY = 100 };
            TestAssert.True(PlanImageViewport.CalculateIsodoseFit(image, 10) == null, "Without isodoses the caller keeps the default frame.");
            var path = new ReviewImagePath { Closed = true };
            foreach (var p in new[] { new[] { 80.0, 90.0 }, new[] { 130.0, 90.0 }, new[] { 130.0, 110.0 }, new[] { 80.0, 110.0 } })
                path.Points.Add(new ReviewImagePoint { X = p[0], Y = p[1] });
            image.Overlays.Add(new ReviewImageOverlay { Kind = "isodose", SourceStatus = ReviewStatusCodes.Available, DoseGy = 10, Paths = new List<ReviewImagePath> { path } });
            image.Overlays.Add(new ReviewImageOverlay { Kind = "structure", SourceStatus = ReviewStatusCodes.Available,
                Paths = new List<ReviewImagePath> { new ReviewImagePath { Points = { new ReviewImagePoint { X = 0, Y = 0 }, new ReviewImagePoint { X = 199, Y = 199 } } } } });
            var fit = PlanImageViewport.CalculateIsodoseFit(image, 10);
            TestAssert.NotNull(fit);
            // Width 50 px x 1 mm = 50 mm, height 20 px x 2 mm = 40 mm -> half 25 mm + 10 mm margin; structures are ignored.
            TestAssert.True(Math.Abs(fit.HalfExtentMillimeters - 35) < 1e-9, "Half extent was " + fit.HalfExtentMillimeters);
            TestAssert.True(Math.Abs(fit.CenterPixelX - 105) < 1e-9 && Math.Abs(fit.CenterPixelY - 100) < 1e-9);
            TestAssert.True(fit.MapX(80) > 0 && fit.MapX(130) < 1 && fit.MapY(90) > 0 && fit.MapY(110) < 1, "Isodose must stay inside the frame.");
            // The GUI uses the tight frame; the report path keeps the 2%-region/full-CT frame.
            TestAssert.True(Math.Abs(PlanImageRenderer.CalculateGuiViewport(image, true).HalfExtentMillimeters - 35) < 1e-9);
            TestAssert.True(PlanImageRenderer.CalculateViewport(image, true).HalfExtentMillimeters > 35);
        }

        public static void ReferenceIsodoseFit()
        {
            var image = new ReviewPlanImage { WidthPixels = 200, HeightPixels = 200,
                PixelSpacingXMillimeters = 1, PixelSpacingYMillimeters = 1, IsocenterPixelX = 100, IsocenterPixelY = 100 };
            Func<double, double, ReviewImageOverlay> level = (percent, half) =>
            {
                var path = new ReviewImagePath { Closed = true };
                path.Points.Add(new ReviewImagePoint { X = 100 - half, Y = 100 - half });
                path.Points.Add(new ReviewImagePoint { X = 100 + half, Y = 100 - half });
                path.Points.Add(new ReviewImagePoint { X = 100 + half, Y = 100 + half });
                return new ReviewImageOverlay { Kind = "isodose", SourceStatus = ReviewStatusCodes.Available, PrescriptionPercent = percent,
                    DoseGy = percent / 100 * 60, Paths = new List<ReviewImagePath> { path } };
            };
            image.Overlays.Add(level(2, 90)); image.Overlays.Add(level(20, 40)); image.Overlays.Add(level(50, 20));
            // 20 % reference: half extent 40 mm + 10 mm, not the 2 % line at 90 mm.
            Close(50, PlanImageViewport.CalculateIsodoseFit(image, 10, 20).HalfExtentMillimeters);
            // 30 % is not shown: the lowest shown level above it (50 %) is used.
            Close(30, PlanImageViewport.CalculateIsodoseFit(image, 10, 30).HalfExtentMillimeters);
            // No shown level at or above 100 %: all shown isodoses (outermost 2 %).
            Close(100, PlanImageViewport.CalculateIsodoseFit(image, 10, 100).HalfExtentMillimeters);
            var previous = ClearPlan.Core.Settings.DisplayPreferences.Current;
            try
            {
                ClearPlan.Core.Settings.DisplayPreferences.Current = new ClearPlan.Core.Settings.DisplayPreferences();
                Close(50, PlanImageRenderer.CalculateGuiViewport(image, true).HalfExtentMillimeters);
            }
            finally { ClearPlan.Core.Settings.DisplayPreferences.Current = previous; }
        }

        public static void PreferencePrecedence()
        {
            string dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "clearplan-display-" + Guid.NewGuid().ToString("N"));
            try
            {
                var store = new ClearPlan.Core.Settings.DisplayPreferencesStore(System.IO.Path.Combine(dir, "DisplayDefaults.json"),
                    System.IO.Path.Combine(dir, "State", "Benutzer", ClearPlan.Core.Settings.DisplayPreferencesStore.UserFileName(@"UKL\Max Muster")));
                TestAssert.Equal(20.0, store.Load().ZoomIsodosePercent); TestAssert.Equal("default", store.LastSource);
                store.SaveGlobal(new ClearPlan.Core.Settings.DisplayPreferences { ZoomIsodosePercent = 30, BevMode = "mlc" });
                TestAssert.Equal(30.0, store.Load().ZoomIsodosePercent); TestAssert.Equal("global", store.LastSource);
                store.SaveUser(new ClearPlan.Core.Settings.DisplayPreferences { ZoomIsodosePercent = 50, BevMode = "ptv" });
                var user = store.Load();
                TestAssert.Equal(50.0, user.ZoomIsodosePercent); TestAssert.Equal("ptv", user.BevMode); TestAssert.Equal("user", store.LastSource);
                TestAssert.True(store.ResetUser());
                TestAssert.Equal("mlc", store.Load().BevMode); TestAssert.Equal("global", store.LastSource);
                TestAssert.False(store.ResetUser(), "Deleting twice must report that no user file exists.");
                TestAssert.Equal("uklmaxmuster.display.json", ClearPlan.Core.Settings.DisplayPreferencesStore.UserFileName("UKL Max Muster"));
                var odd = new ClearPlan.Core.Settings.DisplayPreferences { ZoomIsodosePercent = 900, ZoomMarginMillimeters = -3, BevMode = "x" }.Normalized();
                TestAssert.True(odd.ZoomIsodosePercent == 100 && odd.ZoomMarginMillimeters == 0 && odd.BevMode == "drr");
            }
            finally { try { System.IO.Directory.Delete(dir, true); } catch (Exception) { } }
        }

        private static void Close(double expected, double actual)
        {
            TestAssert.True(Math.Abs(expected - actual) < 1e-9, "Expected " + expected + ", got " + actual);
        }

        public static void CollimatorRotationSource()
        {
            var cp = new ReviewControlPointSample { CollimatorAngleDegrees = 15 };
            var beam = new ReviewBeamAnalysis();
            // Without native calibration the GUI uses the nominal collimator angle, never 0.
            Close(15, BeamEyeViewRenderer.GuiCollimatorRotation(beam, cp).Value);
            beam.BldToDisplayRotationDegrees = -15;
            Close(-15, BeamEyeViewRenderer.GuiCollimatorRotation(beam, cp).Value);
            cp.CollimatorAngleDegrees = 350; beam.BldToDisplayRotationDegrees = null;
            Close(-10, BeamEyeViewRenderer.GuiCollimatorRotation(beam, cp).Value);
            TestAssert.True(BeamEyeViewRenderer.GuiCollimatorRotation(null, null) == null);
            // The report path keeps its validated frame: no calibrated image means no rotation there.
            var state = BeamEyeViewRenderer.InspectState(beam, cp, null, false);
            Close(0, state.DisplayRotationDegrees);
        }

        public static void ApertureFit()
        {
            Func<double, double, ReviewControlPointSample> cp = (open, jaw) => new ReviewControlPointSample
            {
                Aperture = new ApertureGeometry
                {
                    Jaws = new ApertureRectangle(-jaw, -jaw, jaw, jaw),
                    EffectiveOpenings = new List<ApertureRectangle> { new ApertureRectangle(-open, -10, open, 10) }
                }
            };
            var beam = new ReviewBeamAnalysis();
            beam.ControlPoints.Add(cp(20, 30));
            beam.ControlPoints.Add(cp(42, 30));
            // Outermost over all CPs is the 42 mm opening -> 52 mm, rounded up to 55 mm.
            TestAssert.Equal(55.0, BeamEyeViewRenderer.ApertureFitExtentMm(beam, beam.ControlPoints[0], 10).Value);
            beam.ControlPoints.Clear();
            beam.ControlPoints.Add(cp(5, 8));
            TestAssert.Equal(30.0, BeamEyeViewRenderer.ApertureFitExtentMm(beam, beam.ControlPoints[0], 10).Value);
            TestAssert.True(BeamEyeViewRenderer.ApertureFitExtentMm(new ReviewBeamAnalysis(), new ReviewControlPointSample(), 10) == null);
        }
    }
}
