using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Core.Tests
{
    internal static class PlanImagesWorkspaceTests
    {
        public static void RunAll()
        {
            ThreePlanesAndSelection();
            MissingAndWrongPlanNeverBorrowImages();
            DisplayControlsLeaveSnapshotUntouched();
            LoadingAndReplacementAreExplicit();
            SharedViewportAndOverlayRendering();
        }

        private static Type WorkspaceType()
        {
            var type = typeof(ReviewWorkspaceViewModel).Assembly.GetType("ClearPlan.Presentation.ViewModels.PlanImagesViewModel");
            TestAssert.NotNull(type, "Missing detached three-plane CT workspace.");
            return type;
        }

        private static dynamic Workspace(ReviewPlanImage[] images, string planKey)
        { return Activator.CreateInstance(WorkspaceType(), new object[] { images, planKey }); }

        private static void ThreePlanesAndSelection()
        {
            dynamic vm = Workspace(SyntheticPlanImageFactory.Create("current").ToArray(), "current");
            TestAssert.True((bool)vm.HasImages);
            TestAssert.Equal(3, ((IEnumerable)vm.Planes).Cast<object>().Count());
            TestAssert.Equal(3, ((IEnumerable)vm.VisiblePlanes).Cast<object>().Count());
            TestAssert.Equal(3, (int)vm.PlaneColumns);
            vm.EnlargePlaneCommand.Execute("coronal");
            TestAssert.True((bool)vm.IsEnlarged);
            TestAssert.Equal("Schnittbilder · Koronal", (string)vm.HeadingText);
            TestAssert.Equal(1, ((IEnumerable)vm.VisiblePlanes).Cast<object>().Count());
            dynamic plane = ((IEnumerable)vm.VisiblePlanes).Cast<object>().Single();
            TestAssert.Equal("coronal", (string)plane.Kind);
            TestAssert.NotNull((object)plane.ImageSource);
            vm.ShowAllPlanesCommand.Execute(null);
            TestAssert.False((bool)vm.IsEnlarged);
            TestAssert.Equal(3, ((IEnumerable)vm.VisiblePlanes).Cast<object>().Count());
        }

        private static void MissingAndWrongPlanNeverBorrowImages()
        {
            dynamic vm = Workspace(SyntheticPlanImageFactory.Create("other").ToArray(), "current");
            TestAssert.False((bool)vm.HasImages, "Never show a prior plan's CT while current images are unavailable.");
            TestAssert.True(!string.IsNullOrWhiteSpace((string)vm.AvailabilityText));
            foreach (dynamic plane in vm.Planes) TestAssert.True(plane.ImageSource == null);
            var images = SyntheticPlanImageFactory.Create("current");
            images[0].GrayscalePixels = null;
            images[1].SourceStatus = ReviewStatusCodes.Unavailable;
            images[2].PixelSpacingYMillimeters = double.NaN;
            vm = Workspace(images.ToArray(), "current");
            TestAssert.False((bool)vm.HasImages);
            foreach (dynamic plane in vm.Planes)
            {
                TestAssert.True(plane.ImageSource == null);
                TestAssert.True(!string.IsNullOrWhiteSpace((string)plane.AvailabilityText));
            }
            images = SyntheticPlanImageFactory.Create("current");
            images.Add(images[0]);
            vm = Workspace(images.ToArray(), "current");
            dynamic ambiguous = ((IEnumerable)vm.Planes).Cast<dynamic>().Single(p => p.Kind == "transversal");
            TestAssert.False((bool)ambiguous.IsAvailable, "Ambiguous duplicate plane captures are not interchangeable.");
            images = SyntheticPlanImageFactory.Create("current");
            images[0].Overlays.Add(new ReviewImageOverlay { Kind = "isodose", Label = "Dose", SourceStatus = ReviewStatusCodes.Unavailable,
                UnavailableReason = "Synthetic source unavailable fixture" });
            vm = Workspace(images.ToArray(), "current");
            dynamic incomplete = ((IEnumerable)vm.Planes).Cast<dynamic>().Single(p => p.Kind == "transversal");
            TestAssert.True(((string)incomplete.AvailabilityText).Contains("1 Quelle fehlt"), "Missing overlay source must be named visibly, not look like zero contours.");
        }

        private static void DisplayControlsLeaveSnapshotUntouched()
        {
            var images = SyntheticPlanImageFactory.Create("current").ToArray();
            images[0].Overlays.Add(Overlay("isodose", "#FF0000", 80));
            images[0].Overlays.Add(Overlay("structure", "#00FF00", 220));
            byte[] before = images[0].GrayscalePixels.ToArray();
            dynamic vm = Workspace(images, "current");
            TestAssert.True((bool)vm.ShowDose && (bool)vm.ShowStructures && (bool)vm.FocusIsocenter);
            vm.ShowDose = false; vm.ShowStructures = false; vm.FocusIsocenter = false;
            TestAssert.False((bool)vm.ShowDose || (bool)vm.ShowStructures || (bool)vm.FocusIsocenter);
            TestAssert.True(images[0].GrayscalePixels.SequenceEqual(before));
            TestAssert.Equal(2, images[0].Overlays.Count);
            TestAssert.True(images[0].IsocenterPixelX.HasValue);
            TestAssert.True(((string)vm.WindowLevelDescription).Contains("400") && ((string)vm.WindowLevelDescription).Contains("40"));
        }

        private static void SharedViewportAndOverlayRendering()
        {
            var renderer = typeof(ClearPlan.Rendering.BeamEyeViewRenderer).Assembly.GetType("ClearPlan.Rendering.PlanImageRenderer");
            TestAssert.NotNull(renderer, "CT workspace needs shared physical-coordinate rendering.");
            var image = SyntheticPlanImageFactory.Create("current")[0];
            image.PixelSpacingXMillimeters = 2; image.PixelSpacingYMillimeters = 3;
            image.IsocenterPixelX = 100; image.IsocenterPixelY = 130;
            image.DoseFocusRegion = new ReviewImageDoseRegion { PrescriptionPercent = 2, ThresholdGy = 1,
                MinPixelX = 80, MaxPixelX = 130, MinPixelY = 120, MaxPixelY = 160 };
            dynamic focus = renderer.GetMethod("CalculateViewport").Invoke(null, new object[] { image, true });
            var expected = PlanImageViewport.Calculate(image);
            TestAssert.Equal(expected.MapX(42.5), (double)focus.MapX(42.5));
            TestAssert.Equal(expected.MapY(77.0), (double)focus.MapY(77.0));
            dynamic fit = renderer.GetMethod("CalculateViewport").Invoke(null, new object[] { image, false });
            TestAssert.False((bool)fit.UsesDoseRegion);
            TestAssert.Equal(0.5, (double)fit.MapX((image.WidthPixels - 1.0) / 2));
            TestAssert.True(fit.MapX(-0.5) >= 0 && fit.MapX(image.WidthPixels - 0.5) <= 1);
            TestAssert.True(fit.MapY(-0.5) >= 0 && fit.MapY(image.HeightPixels - 0.5) <= 1);
            image.Overlays.Add(Overlay("isodose", "#FF0000", 80));
            image.Overlays.Add(Overlay("structure", "#00FF00", 220));
            var render = renderer.GetMethod("Render");
            byte[] both = (byte[])render.Invoke(null, new object[] { image, true, true, false });
            byte[] ct = (byte[])render.Invoke(null, new object[] { image, false, false, false });
            byte[] dose = (byte[])render.Invoke(null, new object[] { image, false, true, false });
            byte[] structures = (byte[])render.Invoke(null, new object[] { image, true, false, false });
            TestAssert.False(both.SequenceEqual(ct));
            TestAssert.False(dose.SequenceEqual(ct));
            TestAssert.False(structures.SequenceEqual(ct));
            TestAssert.False(dose.SequenceEqual(structures));
            using (var stream = new MemoryStream(both))
            using (var bitmap = new Bitmap(stream))
            { TestAssert.True(bitmap.Width >= 720 && bitmap.Height >= 720); }
        }

        private static void LoadingAndReplacementAreExplicit()
        {
            var images = SyntheticPlanImageFactory.Create("current").ToArray();
            dynamic vm = Workspace(images, "current");
            vm.EnlargePlaneCommand.Execute("sagittal");
            vm.ShowDose = false;
            vm.SetLoading("Schnittbilder werden geladen …");
            TestAssert.True((bool)vm.IsLoading);
            TestAssert.False((bool)vm.HasImages);
            TestAssert.False((bool)vm.ReloadCommand.CanExecute(null));
            vm.ReplaceImages(images);
            TestAssert.False((bool)vm.IsLoading);
            TestAssert.True((bool)vm.HasImages);
            TestAssert.False((bool)vm.ShowDose);
            TestAssert.True((bool)vm.IsEnlarged);
            vm.SetFailure("CT konnte nicht geladen werden.");
            TestAssert.False((bool)vm.HasImages);
            TestAssert.True(((string)vm.StatusText).Contains("nicht geladen"));
            TestAssert.True((bool)vm.ReloadCommand.CanExecute(null));
        }

        private static ReviewImageOverlay Overlay(string kind, string color, double y)
        {
            var overlay = new ReviewImageOverlay { Kind = kind, Label = kind, ColorHex = color, SourceStatus = ReviewStatusCodes.Available };
            var path = new ReviewImagePath();
            path.Points.Add(new ReviewImagePoint { X = 20, Y = y });
            path.Points.Add(new ReviewImagePoint { X = 290, Y = y });
            overlay.Paths.Add(path); return overlay;
        }
    }
}
