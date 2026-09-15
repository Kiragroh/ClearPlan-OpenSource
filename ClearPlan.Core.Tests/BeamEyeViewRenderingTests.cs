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
                draw.Invoke(null, new object[] { graphics, new ReviewControlPointSample { GantryAngleDegrees=angle, PatientSupportAngleDegrees=angle, CollimatorAngleDegrees=angle }, null });
                for (int y = 810; y <= 813; y++)
                for (int x = 1040; x <= 1470; x++)
                    TestAssert.True(bitmap.GetPixel(x,y).R > 245 && bitmap.GetPixel(x,y).G > 245 && bitmap.GetPixel(x,y).B > 245,
                        "Orientation symbol touches heading at angle " + angle);
            }
        }

        public static void CollimatorAngleChangesOnlyItsOwnSchematic()
        {
            using (var first = Orientation(Angles(40, 25, 0)))
            using (var second = Orientation(Angles(40, 25, 90)))
            {
                TestAssert.Equal(0, PixelChanges(first, second, OrientationInset(0)), "Collimator must not rotate the room-view gantry.");
                TestAssert.Equal(0, PixelChanges(first, second, OrientationInset(1)), "Collimator must not rotate the top-view couch.");
                TestAssert.True(PixelChanges(first, second, OrientationSymbol(2)) > 80, "A collimator-only change needs a third, rotating schematic, not only a numeric row.");
            }
        }

        public static void OrientationFootnoteStaysClearOfTheSyntheticBanner()
        {
            using (var bitmap = Orientation(Angles(40, 25, 35)))
                TestAssert.Equal(0, CountNonWhite(bitmap, new Rectangle(1030, 966, 440, 34)),
                    "Orientation content must end above the synthetic banner at y967, with a one-pixel gap.");
        }

        public static void ThreeOrientationAnglesVaryIndependently()
        {
            using (var original = Orientation(Angles(0, 0, 0)))
            for (int changed = 0; changed < 3; changed++)
            using (var rotated = Orientation(Angles(changed == 0 ? 90 : 0, changed == 1 ? 90 : 0, changed == 2 ? 90 : 0)))
            {
                for (int inset = 0; inset < 3; inset++)
                {
                    int differences = PixelChanges(original, rotated, OrientationInset(inset));
                    TestAssert.True(changed == inset ? differences > 80 : differences == 0,
                        "Changing axis " + changed + " must affect only its own independently labelled inset; inspected " + inset + ".");
                }
            }
        }

        public static void CollimatorCardinalAndObliqueAxesRemainUnambiguous()
        {
            double[] angles = { 0, 90, 180, 270, 35 };
            for (int firstIndex = 0; firstIndex < angles.Length; firstIndex++)
            using (var first = Orientation(Angles(0, 0, angles[firstIndex])))
            {
                // The positive BLD X direction starts right and rotates clockwise in this
                // explicitly nominal beam-axis schematic, not in patient coordinates.
                double radians = angles[firstIndex] * Math.PI / 180;
                int x = (int)Math.Round(1397 + 17 * Math.Cos(radians));
                int y = (int)Math.Round(861 + 17 * Math.Sin(radians));
                TestAssert.True(HasTealNear(first, x, y), "Positive BLD X arrow missing for collimator " + angles[firstIndex] + ".");
                for (int secondIndex = firstIndex + 1; secondIndex < angles.Length; secondIndex++)
                using (var second = Orientation(Angles(0, 0, angles[secondIndex])))
                    TestAssert.True(PixelChanges(first, second, OrientationSymbol(2)) > 60,
                        "Collimator symbols must distinguish " + angles[firstIndex] + " from " + angles[secondIndex] + " without numeric labels.");
            }
        }

        public static void OrientationWraparoundIsEquivalentForAllAxes()
        {
            foreach (double angle in new[] { 0d, 35d, 90d, 180d, 270d })
            using (var expected = Orientation(Angles(angle, angle, angle)))
            foreach (double wrapped in new[] { angle - 720, angle + 360, angle + 1080 })
            using (var actual = Orientation(Angles(wrapped, wrapped, wrapped)))
                TestAssert.Equal(0, PixelChanges(expected, actual, new Rectangle(1030, 814, 440, 135)),
                    "Equivalent wrapped angles must produce equivalent symbols and normalized numeric labels.");
        }

        public static void MissingOrientationAnglesDoNotHideTheOtherAxes()
        {
            using (var expected = Orientation(Angles(40, 25, 35)))
            foreach (double invalid in new[] { double.NaN, double.PositiveInfinity, double.NegativeInfinity })
            for (int missing = 0; missing < 3; missing++)
            using (var actual = Orientation(Angles(missing == 0 ? invalid : 40, missing == 1 ? invalid : 25, missing == 2 ? invalid : 35)))
            {
                for (int inset = 0; inset < 3; inset++)
                    if (inset != missing)
                        TestAssert.Equal(0, PixelChanges(expected, actual, OrientationInset(inset)), "Missing axis " + missing + " must not suppress valid axis " + inset + ".");
                TestAssert.True(CountTeal(actual, OrientationSymbol(missing)) == 0, "An unavailable angle must not retain a nominal zero-angle symbol.");
                TestAssert.True(PixelChanges(expected, actual, OrientationInset(missing)) > 60, "The unavailable inset must explicitly change its own status.");
            }
            using (var missingCp = Orientation(null))
                for (int inset = 0; inset < 3; inset++)
                    TestAssert.Equal(0, CountTeal(missingCp, OrientationSymbol(inset)), "No selected CP must not invent any angle.");
        }

        public static void OrientationKeepsFixedZeroReferencesAndNumericLabels()
        {
            foreach (double angle in new[] { 0d, 35d, 90d, 180d, 270d, 359.9d })
            using (var labelled = Orientation(Angles(angle, angle, angle)))
            for (int inset = 0; inset < 3; inset++)
                TestAssert.True(CountNonWhite(labelled, new Rectangle(OrientationInset(inset).X, 906, OrientationInset(inset).Width, 23)) > 50,
                    "Axis " + inset + " name and numeric angle must remain visible for " + angle + " degrees, including three digits and decimals.");
            using (var first = Orientation(Angles(0, 0, 0)))
            using (var second = Orientation(Angles(35, 35, 35)))
            for (int inset = 0; inset < 3; inset++)
            {
                int centre = new[] { 1103, 1250, 1397 }[inset];
                var zeroLabel = new Rectangle(centre + 33, 815, 25, 19);
                TestAssert.True(CountNonWhite(first, zeroLabel) > 10, "Every schematic needs a visible fixed zero reference.");
                TestAssert.Equal(0, PixelChanges(first, second, zeroLabel), "The zero reference must not rotate with its axis.");
                var label = new Rectangle(OrientationInset(inset).X, 906, OrientationInset(inset).Width, 24);
                TestAssert.True(PixelChanges(first, second, label) > 20, "Each inset must retain its current numeric angle beside its own name.");
            }
        }

        public static void OrientationSurvivesMissingDrrAndCompactOutput()
        {
            var beam = SyntheticPlanAnalysisFactory.Create(true, true).Beams[0];
            var cp = beam.ControlPoints[15];
            cp.GantryAngleDegrees = 40;
            cp.PatientSupportAngleDegrees = 25;
            cp.CollimatorAngleDegrees = 35;
            var image = Image(cp);
            image.BldToDisplayRotationDegrees = 35;
            foreach (var size in new[] { new Size(1500, 1000), new Size(1000, 667) })
            using (var availableStream = new MemoryStream(BeamEyeViewRenderer.Render(beam, cp, image, true, size.Width, size.Height)))
            using (var unavailableStream = new MemoryStream(BeamEyeViewRenderer.Render(beam, cp, null, true, size.Width, size.Height)))
            using (var available = new Bitmap(availableStream))
            using (var unavailable = new Bitmap(unavailableStream))
            {
                float scale = size.Width / 1500f;
                for (int inset = 0; inset < 3; inset++)
                {
                    Rectangle symbol = OrientationSymbol(inset);
                    // Calibration provenance intentionally changes the collimator's upper
                    // reference and lower caption, but identical rotations retain the axes.
                    if (inset == 2) symbol = new Rectangle(symbol.X, 834, symbol.Width, 69);
                    var scaledSymbol = new Rectangle((int)(symbol.X * scale), (int)(symbol.Y * scale + 1), (int)(symbol.Width * scale), (int)(symbol.Height * scale));
                    TestAssert.Equal(0, PixelChanges(available, unavailable, scaledSymbol), "Equivalent calibrated/nominal rotations must retain their symbols without a DRR.");
                    TestAssert.True(CountTeal(unavailable, scaledSymbol) > 15, "Every axis needs a visible symbol at output width " + size.Width + ".");
                }
                var modeCaption = new Rectangle((int)(1324 * scale), (int)(929 * scale + 1), (int)(146 * scale), (int)(17 * scale));
                TestAssert.True(PixelChanges(available, unavailable, modeCaption) > 15, "Calibrated screen alignment and nominal fallback must be visibly distinguished.");
            }
        }

        public static void CalibratedCollimatorAxesFollowTheValidatedScreenRotation()
        {
            var beam = SyntheticPlanAnalysisFactory.Create(true, true).Beams[0];
            var cp = beam.ControlPoints[15];
            cp.GantryAngleDegrees = 181;
            cp.PatientSupportAngleDegrees = 0;
            cp.CollimatorAngleDegrees = 90;
            var image = SyntheticDrrFactory.Create(beam, cp);
            image.BldToDisplayRotationDegrees = -90;
            var state = BeamEyeViewRenderer.InspectState(beam, cp, image, true);
            TestAssert.True(state.UprightDisplay && state.DisplayRotationDegrees == -90, "Fixture requires validated native-sign display calibration.");
            using (var calibrated = Rendered(beam, cp, image))
            using (var nominal = Rendered(beam, cp, null))
            {
                TestAssert.True(HasTealNear(calibrated, 1397, 844), "Native C90 with display rotation -90 must point BLD +X up, matching the visible aperture.");
                TestAssert.False(HasTealNear(calibrated, 1397, 878), "Calibrated BLD +X must not retain the opposite nominal +90 direction.");
                TestAssert.True(HasInkNear(calibrated, 1380, 861), "The calibrated BLD +Y arrow must point screen-left.");
                TestAssert.True(HasTealNear(nominal, 1397, 878), "Without a validated display frame C90 must use its labelled nominal clockwise convention.");
                TestAssert.True(HasInkNear(nominal, 1414, 861), "Nominal C90 puts BLD +Y screen-right.");
                TestAssert.Equal(0, PixelChanges(calibrated, nominal, new Rectangle(1324, 906, 146, 23)),
                    "Both captions must retain the native collimator 90 degrees, never the display transform -90 degrees.");
                TestAssert.True(PixelChanges(calibrated, nominal, new Rectangle(1430, 815, 25, 19)) > 10,
                    "A calibrated screen Up reference must differ from the nominal zero-angle reference.");
                TestAssert.True(PixelChanges(calibrated, nominal, new Rectangle(1324, 929, 146, 17)) > 15,
                    "The BLD screen and BLD nominal captions must disclose their different frames.");
                for (int inset = 0; inset < 2; inset++)
                    TestAssert.Equal(0, PixelChanges(calibrated, nominal, OrientationInset(inset)), "Display calibration must not change the gantry or couch inset.");
            }
        }

        public static void RejectedOrMissingCalibrationNeverControlsTheCollimatorInset()
        {
            var beam = SyntheticPlanAnalysisFactory.Create(true, true).Beams[0];
            var cp = beam.ControlPoints[15];
            cp.CollimatorAngleDegrees = 90;
            using (var nominal = Rendered(beam, cp, null))
            for (int reason = 0; reason < 8; reason++)
            {
                var image = Image(cp);
                image.BldToDisplayRotationDegrees = -90;
                if (reason == 0) image.ControlPointIndex++;
                if (reason == 1) image.CollimatorAngleDegrees++;
                if (reason == 2) image.SourceStatus = "unavailable";
                if (reason == 3) image.GrayscalePixels = new byte[3];
                if (reason == 4) image.BldToDisplayRotationDegrees = null;
                if (reason == 5) image.BldToDisplayRotationDegrees = double.NaN;
                if (reason == 6) image.BldToDisplayRotationDegrees = double.PositiveInfinity;
                if (reason == 7) image.Synthetic = false;
                TestAssert.False(BeamEyeViewRenderer.InspectState(beam, cp, image, true).UprightDisplay, "Fixture must reject unvalidated calibration " + reason + ".");
                using (var rejected = Rendered(beam, cp, image))
                    TestAssert.Equal(0, PixelChanges(nominal, rejected, OrientationInset(2)),
                        "Rejected or missing calibration " + reason + " must preserve the complete nominal inset, including its reference and mode label.");
            }
        }

        private static Bitmap Rendered(ReviewBeamAnalysis beam, ReviewControlPointSample cp, BeamEyeViewImage image)
        {
            using (var stream = new MemoryStream(BeamEyeViewRenderer.Render(beam, cp, image, true)))
            using (var bitmap = new Bitmap(stream)) return new Bitmap(bitmap);
        }

        private static bool HasInkNear(Bitmap bitmap, int centreX, int centreY)
        {
            for (int y = centreY - 2; y <= centreY + 2; y++)
            for (int x = centreX - 2; x <= centreX + 2; x++)
            {
                Color color = bitmap.GetPixel(x, y);
                if (color.R < 80 && color.G < 100 && color.B < 130 && color.B - color.R > 10 && color.G - color.R < 30) return true;
            }
            return false;
        }

        private static ReviewControlPointSample Angles(double gantry, double couch, double collimator)
        {
            return new ReviewControlPointSample { GantryAngleDegrees = gantry, PatientSupportAngleDegrees = couch, CollimatorAngleDegrees = collimator };
        }

        private static Rectangle OrientationInset(int index) { return new Rectangle(1030 + index * 147, 814, index == 2 ? 146 : 147, 135); }
        private static Rectangle OrientationSymbol(int index) { return new Rectangle(OrientationInset(index).X, 814, OrientationInset(index).Width, 89); }

        private static Bitmap Orientation(ReviewControlPointSample cp)
        {
            var bitmap = new Bitmap(1500, 1000);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                typeof(BeamEyeViewRenderer).GetMethod("DrawOrientation", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
                    .Invoke(null, new object[] { graphics, cp, null });
            }
            return bitmap;
        }

        private static int PixelChanges(Bitmap first, Bitmap second, Rectangle region)
        {
            int count = 0;
            for (int y = region.Top; y < region.Bottom; y++)
            for (int x = region.Left; x < region.Right; x++)
                if (first.GetPixel(x, y) != second.GetPixel(x, y)) count++;
            return count;
        }

        private static bool HasTealNear(Bitmap bitmap, int centreX, int centreY)
        {
            return CountTeal(bitmap, new Rectangle(centreX - 2, centreY - 2, 5, 5)) > 0;
        }

        private static int CountTeal(Bitmap bitmap, Rectangle region)
        {
            int count = 0;
            for (int y = region.Top; y < region.Bottom; y++)
            for (int x = region.Left; x < region.Right; x++)
            {
                Color color = bitmap.GetPixel(x, y);
                if (color.G - color.R > 45 && color.B - color.R > 35 && color.G < 180) count++;
            }
            return count;
        }

        private static int CountNonWhite(Bitmap bitmap, Rectangle region)
        {
            int count = 0;
            for (int y = region.Top; y < region.Bottom; y++)
            for (int x = region.Left; x < region.Right; x++)
                if (bitmap.GetPixel(x, y).GetBrightness() < 0.9) count++;
            return count;
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
                for (int y = 814; y < 903; y++)
                    for (int x = 1030; x < 1177; x++)
                        TestAssert.Equal(first.GetPixel(x, y), second.GetPixel(x, y), "Couch rotation must not move the fixed room-camera gantry schematic.");
                for (int y = 814; y < 903; y++)
                    for (int x = 1177; x < 1324; x++)
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
                int green = 0, lightGreen = 0;
                for (int y = 210; y < 850; y += 2)
                    for (int x = 60; x < 940; x += 2)
                    {
                        Color color = bitmap.GetPixel(x, y);
                        if (color.G - color.R > 35 && color.G - color.B > 25 && color.R < 130) green++;
                        if (color.G - color.R > 25 && color.G - color.B > 40 && color.R >= 130) lightGreen++;
                    }
                TestAssert.True(green > 100 && lightGreen > 100, "Both physical MLC layers need visible distinct green outlines, beyond legend labels.");
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
