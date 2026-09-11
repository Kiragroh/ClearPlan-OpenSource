using System;
using System.Linq;
using ClearPlan.Core.Review;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Simulation;
using Newtonsoft.Json;

namespace ClearPlan.Core.Tests
{
    internal static class SyntheticPublicationScenarioTests
    {
        public static void FactoryContract()
        {
            var factory = typeof(ReviewSnapshot).Assembly.GetType("ClearPlan.Core.Simulation.SyntheticPublicationScenarioFactory");
            TestAssert.NotNull(factory, "A separate publication fixture factory must exist without changing the seven demonstration scenarios.");
            TestAssert.NotNull(factory.GetMethod("Create", new[] { typeof(bool) }));
        }

        public static void DeterministicAndSynthetic()
        {
            var first = SyntheticPublicationScenarioFactory.Create(false);
            var second = SyntheticPublicationScenarioFactory.Create(false);
            var validation = ReviewSnapshotValidator.ValidateForSimulator(first);
            TestAssert.True(validation.IsValid, string.Join("; ", validation.Issues.Select(i => i.Message)));
            TestAssert.Equal(JsonConvert.SerializeObject(first), JsonConvert.SerializeObject(second));
            TestAssert.True(first.Synthetic);
            TestAssert.True(first.PatientDisplayLabel.StartsWith("Synthetic", StringComparison.Ordinal));
            TestAssert.True(first.ProvenanceText.Contains("not dose calculated from apertures"));
            TestAssert.Equal(7, SyntheticScenarioFactory.ScenarioIds.Count);
            TestAssert.Equal(6, first.DvhSeries.Count);
            TestAssert.True(first.DvhSeries.Single(s => s.StructureId == "CTV_60").RequiredForTargetReview);
            TestAssert.True(first.DvhSeries.Single(s => s.StructureId == "PTV_60").RequiredForTargetReview);
            TestAssert.True(first.DvhSeries.Single(s => s.StructureId == "CTV_60").VolumeCc <
                first.DvhSeries.Single(s => s.StructureId == "PTV_60").VolumeCc);
            string serialized = JsonConvert.SerializeObject(first);
            foreach (string privateText in new[] { "private.example.invalid", "PatientId", "DICOM", "PRIVATE-PATIENT-SENTINEL", "PRIVATE-SOURCE-SENTINEL", "C:\\", "\\\\" })
                TestAssert.False(serialized.Contains(privateText), "Public fixture must not contain a private source or identifier: " + privateText);
            first.PlanImages[0].GrayscalePixels[0] = 99;
            TestAssert.False(second.PlanImages[0].GrayscalePixels[0] == 99, "Factories must not share mutable pixels.");
        }

        public static void GoalsAndQualityUseDisplayedDvh()
        {
            foreach (bool dualLayer in new[] { false, true })
            {
                var snapshot = SyntheticPublicationScenarioFactory.Create(dualLayer);
                foreach (string expectedStatus in new[] { ReviewStatusCodes.Pass, ReviewStatusCodes.Variation, ReviewStatusCodes.Fail })
                    TestAssert.True(snapshot.PqmRows.Any(r => r.Status == expectedStatus), "Publication fixture should calculate an example of " + expectedStatus + ".");
                foreach (var row in snapshot.PqmRows)
                {
                    var curve = snapshot.DvhSeries.Single(s => s.StructureId == row.ResolvedStructureId);
                    double expected = row.Objective == "Dmean" ? SyntheticDvhMetrics.MeanDoseGy(curve.Points) :
                        row.Objective == "V30" ? SyntheticDvhMetrics.RelativeVolumePercentAtDoseGy(curve.Points, 30) :
                        SyntheticDvhMetrics.DoseGyAtRelativeVolumePercent(curve.Points, row.Objective == "D98" ? 98 : 2);
                    Near(expected, row.AchievedValue, 1e-10);
                    TestAssert.Equal(SyntheticDvhMetrics.EvaluateThresholdStatus(expected, row.Comparator, row.Goal.Value, row.Variation.Value), row.Status);
                    TestAssert.False(string.IsNullOrWhiteSpace(row.SourceLabel));
                    TestAssert.True(curve.Selected, "All structures with goals must be selected initially.");
                }
                var coverage = snapshot.PqmRows.Single(r => r.ResolvedStructureId == "PTV_60" && r.Objective == "D98");
                Near(57, coverage.Goal, 0);
                TestAssert.Equal(">=", coverage.Comparator);
                TestAssert.True(coverage.Explanation.Contains("95%"));
                var ptv = snapshot.DvhSeries.Single(s => s.StructureId == "PTV_60");
                var external = snapshot.DvhSeries.Single(s => s.StructureId == "External");
                var actual = snapshot.PlanAnalysis.TargetQuality.Single();
                var expectedQuality = TargetQualityCalculator.Calculate("PTV_60", 60, ptv.VolumeCc,
                    ptv.VolumeCc * SyntheticDvhMetrics.RelativeVolumePercentAtDoseGy(ptv.Points, 60) / 100,
                    external.VolumeCc * SyntheticDvhMetrics.RelativeVolumePercentAtDoseGy(external.Points, 60) / 100,
                    external.VolumeCc * SyntheticDvhMetrics.RelativeVolumePercentAtDoseGy(external.Points, 30) / 100,
                    SyntheticDvhMetrics.DoseGyAtRelativeVolumePercent(ptv.Points, 2),
                    SyntheticDvhMetrics.DoseGyAtRelativeVolumePercent(ptv.Points, 98));
                Near(expectedQuality.PaddickCi.Value, actual.PaddickCi, 1e-10);
                Near(expectedQuality.PlanCheckCi.Value, actual.PlanCheckCi, 1e-10);
                Near(expectedQuality.GradientIndex.Value, actual.GradientIndex, 1e-10);
                Near(expectedQuality.HomogeneityIndex.Value, actual.HomogeneityIndex, 1e-10);
                TestAssert.True(actual.PaddickCi > 0 && actual.PaddickCi <= 1);
            }
        }

        public static void ThreePlanesUseSamePhantomDose()
        {
            var snapshot = SyntheticPublicationScenarioFactory.Create(false);
            TestAssert.Equal(3, snapshot.PlanImages.Count);
            foreach (var image in snapshot.PlanImages)
            {
                TestAssert.True(image.Synthetic && image.PlanKey == snapshot.ActivePlanKey);
                TestAssert.Equal(image.WidthPixels * image.HeightPixels, image.GrayscalePixels.Length);
                TestAssert.True(image.Overlays.Count(o => o.Kind == "structure" && o.Paths.Count > 0) >= 3);
                TestAssert.True(image.Overlays.Count(o => o.Kind == "isodose" && o.Paths.Count > 0) >= 5);
                TestAssert.NotNull(image.DoseFocusRegion);
                Near(2, image.DoseFocusRegion.PrescriptionPercent, 0);
                Near((image.WidthPixels - 1) / 2.0, image.IsocenterPixelX, 0);
                Near((image.HeightPixels - 1) / 2.0, image.IsocenterPixelY, 0);
                var ptv = image.Overlays.Single(o => o.Kind == "structure" && o.Label == "PTV_60");
                foreach (var point in ptv.Paths.SelectMany(p => p.Points))
                {
                    var world = World(image, point);
                    Near(1, world[0] * world[0] / 1600 + world[1] * world[1] / 900 + world[2] * world[2] / 3025, 1e-9);
                }
                foreach (var overlay in image.Overlays.Where(o => o.Kind == "isodose"))
                foreach (var point in overlay.Paths.SelectMany(p => p.Points))
                {
                    var world = World(image, point);
                    double r = Math.Sqrt(world[0] * world[0] / 1600 + world[1] * world[1] / 900 + world[2] * world[2] / 3025);
                    double dose = r <= 1 ? 63 - 2.4 * r * r : 60.6 * Math.Exp(-3 * (r - 1));
                    Near(overlay.DoseGy.Value, dose, 0.15);
                }
            }
            // Independent 3 mm lattice count: External encloses precisely this same ellipsoid.
            int count = 0;
            for (int z = -180; z <= 180; z += 3)
            for (int y = -105; y <= 105; y += 3)
            for (int x = -150; x <= 150; x += 3)
                if (x * x / 22500.0 + y * y / 11025.0 + z * z / 32400.0 <= 1) count++;
            Near(count * 0.027, snapshot.DvhSeries.Single(s => s.StructureId == "External").VolumeCc, 1e-8);
        }

        public static void LayersPamAndEstimatedDoseRate()
        {
            foreach (bool dualLayer in new[] { false, true })
            {
                var plan = SyntheticPublicationScenarioFactory.Create(dualLayer).PlanAnalysis;
                TestAssert.True(plan.Pam.HasValue && plan.Pam >= 0 && plan.Pam <= 1);
                TestAssert.True(plan.GeometryProvenance.Contains("not dose calculated from apertures"));
                foreach (var beam in plan.Beams)
                {
                    TestAssert.Equal(dualLayer ? 2 : 1, beam.MlcLayerCount);
                    TestAssert.Equal(!dualLayer, beam.HasJaws);
                    TestAssert.Equal("Estimated", beam.DoseRateEstimateStatus);
                    TestAssert.True(beam.DoseRateEstimateProfile.StartsWith("SYNTHETIC_", StringComparison.Ordinal));
                    TestAssert.True(beam.ControlPoints.All(p => !p.PlannedDoseRateMuPerMin.HasValue));
                    TestAssert.False(beam.ControlPoints[0].EstimatedDoseRateMuPerMin.HasValue);
                    TestAssert.True(beam.ControlPoints.Skip(1).Select(p => Math.Round(p.EstimatedDoseRateMuPerMin.Value, 3)).Distinct().Count() > 3);
                    foreach (var cp in beam.ControlPoints)
                    {
                        TestAssert.Equal(dualLayer ? 2 : 1, cp.Aperture.Layers.Count);
                        TestAssert.Equal(!dualLayer, cp.Aperture.Jaws != null);
                    }
                    Near(beam.MetersetMu.Value, beam.ControlPoints.Skip(1).Sum(p => p.EstimatedDoseRateMuPerMin.Value * p.EstimatedSegmentDurationSeconds.Value / 60), 1e-8);
                }
            }
        }

        public static void BeamViewsUsePhantomAndTarget()
        {
            var snapshot = SyntheticPublicationScenarioFactory.Create(true);
            foreach (var beam in snapshot.PlanAnalysis.Beams)
            {
                var cp = beam.ControlPoints[0];
                TestAssert.NotNull(cp.BevImage, "First-control-point BEV must include a projection of this publication phantom.");
                TestAssert.True(cp.BevImage.Synthetic);
                TestAssert.Equal(cp.Index, cp.BevImage.ControlPointIndex);
                Near(cp.GantryAngleDegrees, cp.BevImage.GantryAngleDegrees, 0);
                TestAssert.True(cp.BevImage.ProjectionDescription.Contains("same analytical phantom"));
                TestAssert.Equal(cp.BevImage.WidthPixels * cp.BevImage.HeightPixels, cp.BevImage.GrayscalePixels.Length);
                TestAssert.True(cp.BevImage.GrayscalePixels.Distinct().Count() > 30);
                TestAssert.True(beam.ControlPoints.All(p => p.TargetProjectionStrips.Count > 0 && p.TargetProjectionProvenance.Contains("PTV_60 ellipsoid")));
                // Source is on +/-Y at each beam start: expected perspective-projected X/Z ellipsoid area.
                double exactArea = Math.PI * 40 * 55 / (1 - 30.0 * 30 / (1000 * 1000)) / 100;
                Near(exactArea, cp.TargetAreaCm2, 1.0);
            }
        }

        private static double[] World(ReviewPlanImage image, ReviewImagePoint point)
        {
            double u = (point.X - image.IsocenterPixelX.Value) * image.PixelSpacingXMillimeters;
            double v = (point.Y - image.IsocenterPixelY.Value) * image.PixelSpacingYMillimeters;
            return image.Kind == "transversal" ? new[] { u, v, 0.0 } :
                image.Kind == "coronal" ? new[] { u, 0.0, -v } : new[] { 0.0, u, -v };
        }

        private static void Near(double expected, double? actual, double tolerance)
        {
            TestAssert.True(actual.HasValue && !double.IsNaN(actual.Value) && !double.IsInfinity(actual.Value) &&
                Math.Abs(expected - actual.Value) <= tolerance, "Expected " + expected + "; observed " + actual + ".");
        }
    }
}
