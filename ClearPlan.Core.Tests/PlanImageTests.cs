using System;
using System.Linq;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Reporting;

namespace ClearPlan.Core.Tests
{
    internal static class PlanImageTests
    {
        public static void ImageContractExists()
        {
            TestAssert.NotNull(typeof(ReviewSnapshot).Assembly.GetType("ClearPlan.Core.Review.ReviewPlanImage"));
        }

        public static void SyntheticImagesAreDeterministicAndMarked()
        {
            var first = SyntheticPlanImageFactory.Create("SYNTHETIC-PLAN");
            var second = SyntheticPlanImageFactory.Create("SYNTHETIC-PLAN");
            TestAssert.Equal(3, first.Count);
            TestAssert.Equal("transversal", first[0].Kind);
            TestAssert.Equal("coronal", first[1].Kind);
            TestAssert.Equal("sagittal", first[2].Kind);
            for (int i = 0; i < first.Count; i++)
            {
                TestAssert.True(first[i].Synthetic);
                TestAssert.True(first[i].Caption.Contains("SIMULATION"));
                TestAssert.Equal(first[i].WidthPixels * first[i].HeightPixels, first[i].GrayscalePixels.Length);
                TestAssert.True(first[i].GrayscalePixels.SequenceEqual(second[i].GrayscalePixels));
                TestAssert.True(first[i].GrayscalePixels.Distinct().Count() > 5);
            }
            TestAssert.Equal("R", first[0].LeftOrientation);
            TestAssert.Equal("L", first[0].RightOrientation);
            TestAssert.Equal("S", first[1].TopOrientation);
            TestAssert.Equal("A", first[2].LeftOrientation);
        }

        public static void PhysicalGeometryRejectsObliqueAndOutsidePositions()
        {
            TestAssert.Equal(-1, OrthogonalImageGeometry.AxisSign(-1, 0, 0, 0));
            TestAssert.Equal(2, OrthogonalImageGeometry.NearestIndex(6, 10, 2, -1, 6));
            TestAssert.Throws<ArgumentException>(() => OrthogonalImageGeometry.AxisSign(0.707, 0.707, 0, 0));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => OrthogonalImageGeometry.NearestIndex(99, 0, 1, 1, 10));
            TestAssert.Equal((byte)0, OrthogonalImageGeometry.WindowHu(-1000, 40, 400));
            TestAssert.Equal((byte)255, OrthogonalImageGeometry.WindowHu(1000, 40, 400));
        }

        public static void ReportOnlyContainsActivePlanImagesAndDetachedPixels()
        {
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            snapshot.PlanImages = SyntheticPlanImageFactory.Create(snapshot.ActivePlanKey);
            snapshot.PlanImages.AddRange(SyntheticPlanImageFactory.Create("OTHER-PLAN"));
            var report = new ReviewSnapshotReportMapper().Map(snapshot);
            TestAssert.Equal(3, report.PlanImages.Count);
            byte original = report.PlanImages[0].GrayscalePixels[0];
            snapshot.PlanImages[0].GrayscalePixels[0] = 99;
            TestAssert.Equal(original, report.PlanImages[0].GrayscalePixels[0]);
            TestAssert.True(report.Plans.All(plan => plan.PlanKey == snapshot.ActivePlanKey));
        }
    }
}
