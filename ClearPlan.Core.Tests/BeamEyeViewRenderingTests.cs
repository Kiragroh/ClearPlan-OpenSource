using System;
using System.Drawing;
using System.IO;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Simulation;
using ClearPlan.Rendering;

namespace ClearPlan.Core.Tests
{
    internal static class BeamEyeViewRenderingTests
    {
        public static void OrientationSymbolsDoNotTouchTheirHeading()
        {
            var draw = typeof(BeamEyeViewRenderer).GetMethod("DrawOrientation", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            foreach (double angle in new[] { 0d, 40d, 45d, 90d, 180d, 270d, 315d })
            using (var bitmap = new Bitmap(1500, 1000))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                draw.Invoke(null, new object[] { graphics, new ReviewControlPointSample { GantryAngleDegrees=angle, PatientSupportAngleDegrees=angle } });
                for (int y = 810; y <= 813; y++)
                for (int x = 1040; x <= 1470; x++)
                    TestAssert.True(bitmap.GetPixel(x,y).R > 245 && bitmap.GetPixel(x,y).G > 245 && bitmap.GetPixel(x,y).B > 245,
                        "Orientation symbol touches heading at angle " + angle);
            }
        }
        public static void PhysicalWidthsAndIsocenterScaleAreExplicit()
        {
            var method=typeof(BeamEyeViewRenderer).GetMethod("LeafWidthLabel",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Static);
            TestAssert.NotNull(method,"Physical leaf widths at isocenter need an explicit label.");
            var layer=new ApertureLayer { LeafBoundariesMm=new double[] {-5,-2.5,0,5} };
            TestAssert.Equal("Leaf widths at ISO: 2.5 / 5 mm",(string)method.Invoke(null,new object[] {layer}));
            var map=typeof(BeamEyeViewRenderer).GetMethod("Map",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Static,null,new[] {typeof(double),typeof(double),typeof(double)},null);
            var iso=(PointF)map.Invoke(null,new object[] {0.0,0.0,200.0});
            var x=(PointF)map.Invoke(null,new object[] {10.0,0.0,200.0});
            var y=(PointF)map.Invoke(null,new object[] {0.0,10.0,200.0});
            TestAssert.True(Math.Abs(iso.X-500)<0.001 && Math.Abs(iso.Y-504)<0.001);
            TestAssert.True(Math.Abs(x.X-iso.X-22.5)<0.001 && Math.Abs(iso.Y-y.Y-22.5)<0.001,
                "Ten mm must have the same scale on both axes around the native isocenter origin.");
        }
        public static void PngDimensionsAndSyntheticMarker()
        {
            var beam = SyntheticPlanAnalysisFactory.Create(true, true).Beams[0];
            var cp = beam.ControlPoints[15];
            foreach (var size in new[] { new Size(1500, 1000), new Size(1000, 667) })
            {
                byte[] png = BeamEyeViewRenderer.Render(beam, cp, Image(cp), true, size.Width, size.Height);
                using (var stream = new MemoryStream(png))
                using (var bitmap = new Bitmap(stream))
                {
                    TestAssert.Equal(size.Width, bitmap.Width);
                    TestAssert.Equal(size.Height, bitmap.Height);
                    Color mark = bitmap.GetPixel(size.Width - 15, size.Height - 12);
                    TestAssert.True(mark.R > mark.G * 2 && mark.R > mark.B * 2, "Synthetic banner must be visibly red, including compact renders.");
                }
            }
        }

        public static void RejectsDrrControlPointOrAngleMismatch()
        {
            var beam = SyntheticPlanAnalysisFactory.Create(false, true).Beams[0];
            var cp = beam.ControlPoints[15];
            var image = Image(cp);
            TestAssert.True(BeamEyeViewRenderer.InspectState(beam, cp, image, true).ImageAvailable);
            image.ControlPointIndex++;
            TestAssert.False(BeamEyeViewRenderer.InspectState(beam, cp, image, true).ImageAvailable);
            image.ControlPointIndex = cp.Index;
            image.CollimatorAngleDegrees += 1;
            TestAssert.False(BeamEyeViewRenderer.InspectState(beam, cp, image, true).ImageAvailable);
            image.CollimatorAngleDegrees = cp.CollimatorAngleDegrees + 360;
            TestAssert.True(BeamEyeViewRenderer.InspectState(beam, cp, image, true).ImageAvailable);
            image.GantryAngleDegrees += 1;
            TestAssert.False(BeamEyeViewRenderer.InspectState(beam, cp, image, true).ImageAvailable);
            image.GantryAngleDegrees = cp.GantryAngleDegrees;
            image.PatientSupportAngleDegrees += 1;
            TestAssert.False(BeamEyeViewRenderer.InspectState(beam, cp, image, true).ImageAvailable);
            image.PatientSupportAngleDegrees = cp.PatientSupportAngleDegrees;
            TestAssert.False(BeamEyeViewRenderer.InspectState(beam, cp, image, false).ImageAvailable);
            TestAssert.False(BeamEyeViewRenderer.InspectState(beam, cp, null, true).ImageAvailable);
            image.ControlPointIndex++;
            using (var stream = new MemoryStream(BeamEyeViewRenderer.Render(beam, cp, image, true)))
            using (var bitmap = new Bitmap(stream))
                TestAssert.True(bitmap.GetPixel(600, 180).GetBrightness() < 0.4, "Mismatched DRR pixels must not be drawn behind the leaves.");
        }

        public static void JawlessNeverLabelsVirtualBoundsAsJaws()
        {
            var labelPlacement = typeof(BeamEyeViewRenderer).GetMethod("YJawLabelLeft", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            TestAssert.NotNull(labelPlacement, "Physical Y-jaw labels need a collision-free ruler placement policy.");
            foreach (float middle in new[] { 400f, 500f, 560f })
            {
                float left = (float)labelPlacement.Invoke(null, new object[] { middle });
                TestAssert.True(left + 160 <= 488 || left >= 586, "Y-jaw label must not overlap the Y-axis or its tick-label band.");
            }
            var jawOnlyBeam = SyntheticPlanAnalysisFactory.Create(false, true).Beams[0];
            var jawOnlyCp = jawOnlyBeam.ControlPoints[0];
            jawOnlyCp.Aperture.Layers.Clear();
            jawOnlyBeam.MlcLayerCount = 0;
            var jawOnly = BeamEyeViewRenderer.InspectState(jawOnlyBeam, jawOnlyCp, Image(jawOnlyCp), true);
            TestAssert.True(jawOnly.GeometryAvailable && jawOnly.PhysicalJawsVisible, "Verified jaw-only apertures must retain their physical jaw labels.");
            jawOnlyCp.Aperture.Jaws = null;
            jawOnlyBeam.HasJaws = false;
            TestAssert.False(BeamEyeViewRenderer.InspectState(jawOnlyBeam, jawOnlyCp, Image(jawOnlyCp), true).GeometryAvailable,
                "Neither jaws nor MLC must not become a valid aperture.");
            var beam = SyntheticPlanAnalysisFactory.Create(true, true).Beams[0];
            var cp = beam.ControlPoints[15];
            cp.Aperture.FixedBoundingBox = new ApertureRectangle(-140, -140, 140, 140);
            var state = BeamEyeViewRenderer.InspectState(beam, cp, Image(cp), true);
            TestAssert.True(state.GeometryAvailable);
            TestAssert.True(state.FixedBoundsVisible);
            TestAssert.False(state.PhysicalJawsVisible);
            TestAssert.Equal(0, state.JawLabels.Length);
            TestAssert.True(state.LayoutDescription.Contains("Jawless"));
            cp.Aperture.Jaws = new ApertureRectangle(-100, -100, 100, 100);
            TestAssert.False(BeamEyeViewRenderer.InspectState(beam, cp, Image(cp), true).GeometryAvailable);
            beam = SyntheticPlanAnalysisFactory.Create(false, true).Beams[0];
            cp = beam.ControlPoints[15];
            state = BeamEyeViewRenderer.InspectState(beam, cp, Image(cp), true);
            TestAssert.True(state.PhysicalJawsVisible);
            TestAssert.True(state.JawLabels.SequenceEqual(new[] { "X1", "X2", "Y1", "Y2" }));
            cp.Aperture.Layers[0].Bank2PositionsMm = new double[1];
            TestAssert.False(BeamEyeViewRenderer.InspectState(beam, cp, Image(cp), true).GeometryAvailable);
        }

        public static void CardinalGantryOrientationDoesNotDependOnCouch()
        {
            double[] angles = { 0, 90, 180, 270, 360 };
            var expected = new[] { new PointF(0, -1), new PointF(1, 0), new PointF(0, 1), new PointF(-1, 0), new PointF(0, -1) };
            for (int i = 0; i < angles.Length; i++)
            {
                PointF position = BeamEyeViewRenderer.LinacOrientationPoint(angles[i]);
                TestAssert.True(Math.Abs(position.X - expected[i].X) < 0.00001 && Math.Abs(position.Y - expected[i].Y) < 0.00001);
            }
            TestAssert.Throws<ArgumentException>(() => BeamEyeViewRenderer.LinacOrientationPoint(double.NaN));
            var beam = SyntheticPlanAnalysisFactory.Create(true, true).Beams[0];
            var cp = beam.ControlPoints[15];
            cp.PatientSupportAngleDegrees = 0;
            byte[] initial = BeamEyeViewRenderer.Render(beam, cp, Image(cp), true);
            cp.PatientSupportAngleDegrees = 90;
            byte[] rotated = BeamEyeViewRenderer.Render(beam, cp, Image(cp), true);
            using (var firstStream = new MemoryStream(initial))
            using (var secondStream = new MemoryStream(rotated))
            using (var first = new Bitmap(firstStream))
            using (var second = new Bitmap(secondStream))
            {
                int couchChanges = 0;
                for (int y = 811; y < 890; y++)
                    for (int x = 1060; x < 1210; x++)
                        TestAssert.Equal(first.GetPixel(x, y), second.GetPixel(x, y), "Couch rotation must not move the fixed room-camera gantry schematic.");
                for (int y = 811; y < 890; y++)
                    for (int x = 1340; x < 1415; x++)
                        if (first.GetPixel(x, y) != second.GetPixel(x, y)) couchChanges++;
                TestAssert.True(couchChanges > 100, "Couch rotation must change the separate table/patient inset.");
            }
        }

        public static void SyntheticFactorySourceAndUnavailablePixelsAreGated()
        {
            var beam = SyntheticPlanAnalysisFactory.Create(true, true).Beams[0];
            var cp = beam.ControlPoints[0];
            var image = SyntheticDrrFactory.Create(beam, cp);
            TestAssert.True(BeamEyeViewRenderer.InspectState(beam, cp, image, true).ImageAvailable,
                "The actual synthetic DRR factory source must be accepted only in synthetic mode.");
            TestAssert.False(BeamEyeViewRenderer.InspectState(beam, cp, image, false).ImageAvailable);
            image.SourceStatus = "unavailable";
            image.UnavailableReason = "Example source failure.";
            TestAssert.False(BeamEyeViewRenderer.InspectState(beam, cp, image, true).ImageAvailable);
            image.SourceStatus = "available";
            image.GrayscalePixels = new byte[3];
            TestAssert.False(BeamEyeViewRenderer.InspectState(beam, cp, image, true).ImageAvailable);
        }

        public static void DrrOpeningStaysVisibleAndBothMlcLayersAreOutlined()
        {
            var beam = SyntheticPlanAnalysisFactory.Create(true, true).Beams[0];
            var cp = beam.ControlPoints[15];
            using (var stream = new MemoryStream(BeamEyeViewRenderer.Render(beam, cp, Image(cp), true)))
            using (var bitmap = new Bitmap(stream))
            {
                Color opening = bitmap.GetPixel(530, 530);
                TestAssert.Equal(185, (int)opening.R, "Open aperture must retain its original grayscale DRR pixels.");
                TestAssert.Equal(opening.R, opening.G);
                TestAssert.Equal(opening.G, opening.B);
                int cyan = 0, amber = 0;
                for (int y = 210; y < 850; y += 2)
                    for (int x = 60; x < 940; x += 2)
                    {
                        Color color = bitmap.GetPixel(x, y);
                        if (color.B - color.R > 35 && color.G - color.R > 35) cyan++;
                        if (color.R - color.G > 20 && color.G - color.B > 25) amber++;
                    }
                TestAssert.True(cyan > 100 && amber > 100, "Both physical MLC layers need visible distinct outlines, beyond legend labels.");
            }
            // A closed field-start aperture is valid; it must not be replaced by another CP.
            cp = beam.ControlPoints[0];
            foreach (var layer in cp.Aperture.Layers)
            {
                Array.Clear(layer.Bank1PositionsMm, 0, layer.Bank1PositionsMm.Length);
                Array.Clear(layer.Bank2PositionsMm, 0, layer.Bank2PositionsMm.Length);
            }
            cp.Aperture.EffectiveOpenings.Clear();
            cp.ApertureAreaCm2 = 0;
            TestAssert.True(BeamEyeViewRenderer.InspectState(beam, cp, Image(cp), true).GeometryAvailable);
            using (var stream = new MemoryStream(BeamEyeViewRenderer.Render(beam, cp, Image(cp), true)))
            using (var bitmap = new Bitmap(stream)) TestAssert.Equal(1500, bitmap.Width);
        }

        private static BeamEyeViewImage Image(ReviewControlPointSample cp)
        {
            return new BeamEyeViewImage
            {
                WidthPixels = 128, HeightPixels = 128, ExtentMm = 200,
                GrayscalePixels = Enumerable.Repeat((byte)185, 128 * 128).ToArray(),
                Synthetic = true, SourceStatus = "available", ProjectionDescription = "Synthetic constant-intensity test projection, not anatomy.",
                ControlPointIndex = cp.Index, GantryAngleDegrees = cp.GantryAngleDegrees,
                CollimatorAngleDegrees = cp.CollimatorAngleDegrees, PatientSupportAngleDegrees = cp.PatientSupportAngleDegrees
            };
        }
    }
}
