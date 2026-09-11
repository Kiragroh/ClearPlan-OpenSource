using System;
using System.Collections.Generic;
using System.IO;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Core.Tests
{
    internal static class BevActivationTests
    {
        public static void RunAll()
        {
            NativeActivationRequestsOneMissingDrr();
            ExistingNativeImageDoesNotRequestAgain();
            MissingPixelCacheIsRegenerated();
            EmptySelectionDoesNotRequestDrr();
            FailureAndGeometryReasonsRemainVisible();
            LoadedViewReactivatesAfterSnapshotReplacement();
            NativePublicationHonorsCancellation();
        }

        public static void NativeActivationRequestsOneMissingDrr()
        {
            var workspace = Workspace();
            int requests = 0;
            workspace.GenerateDrrRequested += (sender, args) => requests++;
            workspace.Analysis.ActivateBevAsync().GetAwaiter().GetResult();
            workspace.Analysis.ActivateBevAsync().GetAwaiter().GetResult();
            TestAssert.NotNull(workspace.Analysis.BevImageSource, "Native MLC preview is rendered on activation.");
            TestAssert.Equal(1, requests, "Native BEV activation must request its missing CT-derived DRR once.");
        }

        public static void ExistingNativeImageDoesNotRequestAgain()
        {
            var workspace = Workspace();
            workspace.Analysis.SelectedControlPoint.BevImage = AvailableImage();
            int requests = 0;
            workspace.GenerateDrrRequested += (sender, args) => requests++;
            workspace.Analysis.ActivateBevAsync().GetAwaiter().GetResult();
            TestAssert.Equal(0, requests, "An existing matching in-memory DRR is reused.");
        }

        public static void MissingPixelCacheIsRegenerated()
        {
            var workspace = Workspace();
            var image = AvailableImage();
            image.GrayscalePixels = null;
            workspace.Analysis.SelectedControlPoint.BevImage = image;
            int requests = 0;
            workspace.GenerateDrrRequested += (sender, args) => requests++;
            workspace.Analysis.ActivateBevAsync().GetAwaiter().GetResult();
            TestAssert.Equal(1, requests, "JSON-restored metadata without pixels must not suppress native regeneration.");
        }

        public static void EmptySelectionDoesNotRequestDrr()
        {
            var workspace = new ReviewWorkspaceViewModel(new ReviewSnapshot { PlanAnalysis = new ReviewPlanAnalysis() });
            int requests = 0;
            workspace.GenerateDrrRequested += (sender, args) => requests++;
            workspace.Analysis.ActivateBevAsync().GetAwaiter().GetResult();
            TestAssert.Equal(0, requests);
        }

        public static void FailureAndGeometryReasonsRemainVisible()
        {
            var workspace = Workspace();
            const string reason = "ESAPI_BEV_NATIVE_FRAME: Native frame could not be verified.";
            workspace.Analysis.SelectedControlPoint.GeometryReason = "MLC model/profile unavailable.";
            workspace.Analysis.SelectedControlPoint.BevImage = new BeamEyeViewImage
            {
                SourceStatus = "unavailable", UnavailableReason = reason,
                ProjectionDescription = "Generic projection method, not the failure reason."
            };
            workspace.Analysis.RenderBevPreview();
            TestAssert.True(workspace.Analysis.BevStatusText.Contains(reason), "BEV status must expose the actual native failure.");
            TestAssert.True(workspace.Analysis.BevStatusText.Contains("MLC model/profile unavailable."),
                "An independent missing-geometry reason must also remain visible.");
        }

        public static void LoadedViewReactivatesAfterSnapshotReplacement()
        {
            string code = Source("ClearPlan.Presentation", "Views", "BeamEyeView.xaml.cs");
            TestAssert.True(code.Contains("DataContextChanged +=") && code.Contains("IsEnabledChanged +="),
                "An already-loaded BEV must reactivate when a snapshot is replaced or a native operation finishes.");
            TestAssert.True(code.Contains("Dispatcher.BeginInvoke") && code.Contains("DispatcherPriority.Loaded"),
                "Activation must run after the host commits its context and releases its operation guard.");
            TestAssert.True(code.Contains("IsLoaded") && code.Contains("IsVisible") && code.Contains("IsEnabled"),
                "Queued activation must not use hidden, unloaded or disabled native views.");
        }

        public static void NativePublicationHonorsCancellation()
        {
            string code = Source("ClearPlan.Script", "Review", "ClinicalReviewWorkspaceHost.cs");
            int start = code.IndexOf("var image = await new EsapiBevBuilder().BuildAsync", StringComparison.Ordinal);
            int publish = code.IndexOf("cp.BevImage = image;", start, StringComparison.Ordinal);
            TestAssert.True(start >= 0 && publish > start && code.Substring(start, publish - start)
                .Contains("analysisCancellation.Token.ThrowIfCancellationRequested();"),
                "A canceled native DRR must never be published even if the builder completed concurrently.");
        }

        private static ReviewWorkspaceViewModel Workspace()
        {
            return new ReviewWorkspaceViewModel(new ReviewSnapshot
            {
                PlanAnalysis = new ReviewPlanAnalysis
                {
                    Beams = new List<ReviewBeamAnalysis>
                    {
                        new ReviewBeamAnalysis { BeamNumber = 1, BeamId = "Synthetic test beam",
                            ControlPoints = new List<ReviewControlPointSample> { new ReviewControlPointSample() } }
                    }
                }
            });
        }

        private static BeamEyeViewImage AvailableImage()
        {
            return new BeamEyeViewImage { SourceStatus = "available", WidthPixels = 2, HeightPixels = 2,
                ExtentMm = 100, GrayscalePixels = new byte[] { 0, 80, 160, 255 }, ProjectionDescription = "Synthetic test pixels only." };
        }

        private static string Source(params string[] parts)
        {
            var root = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (root != null && !File.Exists(Path.Combine(root.FullName, "ClearPlan.sln"))) root = root.Parent;
            TestAssert.NotNull(root, "Repository root is required for native BEV wiring checks.");
            string path = root.FullName;
            foreach (string part in parts) path = Path.Combine(path, part);
            return File.ReadAllText(path);
        }
    }
}
