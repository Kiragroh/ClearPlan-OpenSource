using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Simulation;
using ClearPlan.Rendering;

namespace ClearPlan.Core.Tests
{
    internal static class BevCollimatorDisplayTests
    {
        public static void PhysicalLayersUseGreenAndBlueWithExplicitMaximumField()
        {
            var color = typeof(BeamEyeViewRenderer).GetMethod("LayerColor", BindingFlags.NonPublic | BindingFlags.Static);
            var first = (Color)color.Invoke(null, new object[] { 0 });
            var second = (Color)color.Invoke(null, new object[] { 1 });
            TestAssert.True(first.G > first.R + 60 && first.G > first.B + 40, "Single/distal MLC must be green.");
            TestAssert.True(second.B > second.R + 60 && second.B > second.G + 20, "Proximal MLC must be visibly blue, not another green/yellow.");
            var beam = Beam(0);
            beam.ControlPoints[0].Aperture.FixedBoundingBox = new ApertureRectangle(-140, -140, 140, 140);
            var state = BeamEyeViewRenderer.InspectState(beam, beam.ControlPoints[0], null, false);
            TestAssert.True(state.FixedBoundsVisible && !state.PhysicalJawsVisible);
            TestAssert.True(state.LayoutDescription.Contains("28") && state.LayoutDescription.Contains("maximum"),
                "Jawless geometry must name its configured maximum field opening, not imply no boundary.");
        }

        public static void CalibratedRasterAndApertureStayRegisteredAtZeroNinetyAndThirtyDegrees()
        {
            foreach (double angle in new[] { 0d, 90d, 30d })
            {
                var beam = Beam(angle);
                var cp = beam.ControlPoints[0];
                var image = LandmarkImage(cp);
                SetCalibration(image, angle);
                var state = BeamEyeViewRenderer.InspectState(beam, cp, image, true);
                TestAssert.True(State<bool>(state, "UprightDisplay"));
                double extent = State<double>(state, "DisplayExtentMm");
                double radians = angle * Math.PI / 180;
                TestAssert.True(extent >= 200 * (Math.Abs(Math.Cos(radians)) + Math.Abs(Math.Sin(radians))) - 0.001,
                    "Rotating the captured DRR must not clip its corners.");
                using (var stream = new MemoryStream(BeamEyeViewRenderer.Render(beam, cp, image, true)))
                using (var bitmap = new Bitmap(stream))
                {
                    // The BLD landmark centre is (+70,+30) mm. CreateFrame defines BLD X=u*cos-v*sin,
                    // BLD Y=u*sin+v*cos, hence display (x*cos+y*sin, -x*sin+y*cos), not a guessed sign.
                    var centre = Display(70, 30, angle, extent);
                    Color landmark = bitmap.GetPixel((int)Math.Round(centre.X), (int)Math.Round(centre.Y));
                    TestAssert.True(landmark.R > 230 && Math.Abs(landmark.R - landmark.G) < 3,
                        "The DRR landmark must move with the calibrated aperture at C=" + angle);
                    var boundary = Display(60, 30, angle, extent);
                    TestAssert.True(HasGreen(bitmap, boundary, 4), "The green aperture must enclose the same DRR landmark at C=" + angle);
                    var background = Display(90, 30, angle, extent);
                    TestAssert.True(bitmap.GetPixel((int)background.X, (int)background.Y).R < 100,
                        "The test must distinguish the landmark from surrounding background.");
                }
            }
        }

        public static void UnknownCalibrationNeverGuessesFromCollimatorAngle()
        {
            var beam = Beam(90);
            var cp = beam.ControlPoints[0];
            var image = LandmarkImage(cp);
            foreach (double? calibration in new double?[] { null, double.NaN, double.PositiveInfinity })
            {
                SetCalibration(image, calibration);
                var state = BeamEyeViewRenderer.InspectState(beam, cp, image, true);
                TestAssert.True(state.ImageAvailable, "Missing display calibration must not discard valid BLD pixels.");
                TestAssert.False(State<bool>(state, "UprightDisplay"));
                TestAssert.Equal(0d, State<double>(state, "DisplayRotationDegrees"));
                TestAssert.True(State<string>(state, "FrameDescription").Contains("BLD"));
            }
            SetCalibration(image, -90);
            TestAssert.Equal(-90d, State<double>(BeamEyeViewRenderer.InspectState(beam, cp, image, true), "DisplayRotationDegrees"),
                "Native calibration, not the raw +90 degree collimator setting, determines the transform.");
            image.ControlPointIndex++;
            TestAssert.False(State<bool>(BeamEyeViewRenderer.InspectState(beam, cp, image, true), "UprightDisplay"),
                "Rejected images cannot lend their calibration to another control point.");
        }

        public static void NumericRulerAnnotationsAreRemovedButIsoAndTicksRemain()
        {
            var draw = typeof(BeamEyeViewRenderer).GetMethod("DrawRuler", BindingFlags.NonPublic | BindingFlags.Static);
            using (var bitmap = new Bitmap(1500, 1000))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Black);
                draw.Invoke(null, new object[] { graphics, 200d });
                for (int y = 520; y < 545; y++)
                for (int x = 262; x < 297; x++)
                    TestAssert.Equal(Color.FromArgb(0, 0, 0), bitmap.GetPixel(x, y), "Redundant -10 ruler text must not be drawn.");
                TestAssert.True(bitmap.GetPixel(275, 500).R > 0, "Physical 1 cm ticks remain.");
                TestAssert.True(bitmap.GetPixel(507, 504).R > 0, "The isocentre circle remains.");
            }
        }

        public static void PhysicalJawValuesFollowRotatedMidpointsAndJawlessHasNoJaws()
        {
            var beam = Beam(90);
            var cp = beam.ControlPoints[0];
            beam.HasJaws = true;
            cp.Aperture.Jaws = new ApertureRectangle(-100, -80, 100, 80);
            var image = LandmarkImage(cp);
            SetCalibration(image, 90);
            var state = BeamEyeViewRenderer.InspectState(beam, cp, image, true);
            TestAssert.True(State<string[]>(state, "JawCoordinateLabels").SequenceEqual(new[] { "X1 -10.0 cm", "X2 10.0 cm", "Y1 -8.0 cm", "Y2 8.0 cm" }));
            var draw = typeof(BeamEyeViewRenderer).GetMethod("DrawJawLabels", BindingFlags.NonPublic | BindingFlags.Static);
            using (var bitmap = new Bitmap(1500, 1000))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Black);
                draw.Invoke(null, new object[] { graphics, cp.Aperture.Jaws, 200d, 90d });
                // X2 midpoint (+100,0) rotates to (0,-100): label must be below ISO, still horizontal.
                int yellow = 0;
                for (int y = 725; y < 795; y++)
                for (int x = 405; x < 600; x++)
                { var color = bitmap.GetPixel(x, y); if (color.R > 150 && color.G > 130 && color.B < 120) yellow++; }
                TestAssert.True(yellow > 80, "The X2 numeric label must follow its rotated physical jaw midpoint.");
            }
            beam.HasJaws = false; beam.MlcLayerCount = 2; cp.Aperture.Jaws = null;
            cp.Aperture.Layers.Add(new ApertureLayer { Label = "Proximal", LeafBoundariesMm = new[] { 20d, 40d },
                Bank1PositionsMm = new[] { 60d }, Bank2PositionsMm = new[] { 80d } });
            state = BeamEyeViewRenderer.InspectState(beam, cp, image, true);
            TestAssert.True(state.GeometryAvailable);
            TestAssert.Equal(0, State<string[]>(state, "JawCoordinateLabels").Length);
        }

        public static void SyntheticFactoriesAndSnapshotPreserveDisplayCalibration()
        {
            var beam = SyntheticPlanAnalysisFactory.Create(false, true).Beams[0];
            var cp = beam.ControlPoints[0]; cp.CollimatorAngleDegrees = 30;
            cp.BevImage = SyntheticDrrFactory.Create(beam, cp);
            TestAssert.Equal((double?)30, Calibration(cp.BevImage));
            var analysis = new ReviewPlanAnalysis(); analysis.Beams.Add(beam);
            var copy = PlanAnalysisSnapshot.Copy(analysis);
            TestAssert.Equal((double?)30, Calibration(copy.Beams[0].ControlPoints[0].BevImage));
            beam.MachineId = "SYNTHETIC_C_ARM";
            TestAssert.Equal((double?)30, Calibration(SyntheticPublicationScenarioFactory.CreateDrr(beam, cp)));
        }

        public static void ProjectedWorldLandmarkRemainsUprightAcrossCollimatorAngles()
        {
            // Actual CT projection through the same production frame helper, independent of the
            // manually encoded raster test: a small synthetic sphere centred at patient (+50,0,+30).
            var volume = new CtVolume { SizeX = 45, SizeY = 45, SizeZ = 45, SpacingX = 5, SpacingY = 5, SpacingZ = 5,
                Origin = new BeamPoint3D(-110, -110, -110), XAxis = new BeamPoint3D(1, 0, 0),
                YAxis = new BeamPoint3D(0, 1, 0), ZAxis = new BeamPoint3D(0, 0, 1), HounsfieldUnits = new float[45 * 45 * 45] };
            for (int z = 0; z < 45; z++) for (int y = 0; y < 45; y++) for (int x = 0; x < 45; x++)
            {
                double dx = -110 + x * 5 - 50, dy = -110 + y * 5, dz = -110 + z * 5 - 30;
                volume.HounsfieldUnits[(z * 45 + y) * 45 + x] = dx * dx + dy * dy + dz * dz < 100 ? 1000 : -1000;
            }
            foreach (double angle in new[] { 0d, 90d, 30d })
            {
                var beam = Beam(angle); var cp = beam.ControlPoints[0];
                beam.HasJaws = true; beam.MlcLayerCount = 0; cp.Aperture.Layers.Clear();
                cp.Aperture.Jaws = new ApertureRectangle(-120, -120, 120, 120); cp.Aperture.EffectiveOpenings.Clear();
                var frame = SyntheticDrrFactory.CreateHfsFrame(0, angle, 0, 1000);
                var image = CtDrrProjector.Project(volume, frame, 192, 200, CancellationToken.None);
                image.Synthetic = true; image.ControlPointIndex = 0; image.CollimatorAngleDegrees = angle;
                SetCalibration(image, angle);
                double extent = State<double>(BeamEyeViewRenderer.InspectState(beam, cp, image, true), "DisplayExtentMm");
                var expected = Display(50, 30, 0, extent);
                using (var stream = new MemoryStream(BeamEyeViewRenderer.Render(beam, cp, image, true)))
                using (var bitmap = new Bitmap(stream))
                    TestAssert.True(bitmap.GetPixel((int)expected.X, (int)expected.Y).R > 200,
                        "A projected world-space landmark must remain upright at the same physical gantry-frame coordinate at C=" + angle);
            }
        }

        private static ReviewBeamAnalysis Beam(double angle)
        {
            var beam = new ReviewBeamAnalysis { BeamId = "SYNTHETIC LANDMARK", HasJaws = false, MlcLayerCount = 1 };
            var cp = new ReviewControlPointSample { Index = 0, CollimatorAngleDegrees = angle, Aperture = new ApertureGeometry() };
            cp.Aperture.Layers.Add(new ApertureLayer { Label = "Synthetic", LeafBoundariesMm = new[] { 20d, 40d },
                Bank1PositionsMm = new[] { 60d }, Bank2PositionsMm = new[] { 80d } });
            cp.Aperture.EffectiveOpenings.Add(new ApertureRectangle(60, 20, 80, 40));
            beam.ControlPoints.Add(cp);
            return beam;
        }

        private static BeamEyeViewImage LandmarkImage(ReviewControlPointSample cp)
        {
            var image = new BeamEyeViewImage { WidthPixels = 400, HeightPixels = 400, ExtentMm = 200,
                GrayscalePixels = Enumerable.Repeat((byte)45, 400 * 400).ToArray(), Synthetic = true,
                SourceStatus = "synthetic", ControlPointIndex = cp.Index, CollimatorAngleDegrees = cp.CollimatorAngleDegrees };
            for (int y = 164; y < 176; y++) for (int x = 264; x < 276; x++) image.GrayscalePixels[y * 400 + x] = 250;
            return image;
        }

        private static PointF Display(double x, double y, double angle, double extent)
        {
            double radians = angle * Math.PI / 180;
            return new PointF((float)(500 + (x * Math.Cos(radians) + y * Math.Sin(radians)) * 450 / extent),
                (float)(504 - (-x * Math.Sin(radians) + y * Math.Cos(radians)) * 450 / extent));
        }

        private static bool HasGreen(Bitmap bitmap, PointF point, int radius)
        {
            for (int y = (int)point.Y - radius; y <= (int)point.Y + radius; y++)
            for (int x = (int)point.X - radius; x <= (int)point.X + radius; x++)
            { var color = bitmap.GetPixel(x, y); if (color.G > 180 && color.G - color.R > 70 && color.G - color.B > 50) return true; }
            return false;
        }

        private static void SetCalibration(BeamEyeViewImage image, double? angle)
        {
            var property = typeof(BeamEyeViewImage).GetProperty("BldToDisplayRotationDegrees");
            TestAssert.NotNull(property, "The DRR must carry an explicit nullable calibrated BLD-to-display rotation.");
            property.SetValue(image, angle, null);
        }

        private static double? Calibration(BeamEyeViewImage image)
        {
            var property = typeof(BeamEyeViewImage).GetProperty("BldToDisplayRotationDegrees");
            TestAssert.NotNull(property, "Synthetic DRRs must declare their tested display convention.");
            return (double?)property.GetValue(image, null);
        }

        private static T State<T>(BeamEyeViewRenderState state, string name)
        {
            var property = typeof(BeamEyeViewRenderState).GetProperty(name);
            TestAssert.NotNull(property, "BEV state must expose " + name + " for accessible review.");
            return (T)property.GetValue(state, null);
        }
    }
}
