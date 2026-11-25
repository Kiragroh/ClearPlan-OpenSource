using System;
using ClearPlan.Core.Review;
using ClearPlan.Core.Settings;

namespace ClearPlan.Core.Tests
{
    internal static class CollisionIntegrationTests
    {
        public static void SourceModelIllustrationDoesNotInventCoordinateConfirmation()
        {
            var snapshot = ClearPlan.Core.Simulation.SyntheticScenarioFactory.Create("baseline-pass");
            var scene = ClearPlan.Core.Collision.SyntheticCollisionFactory.Create(snapshot.ActivePlanKey);
            // Synthetic test fixture exercising the native-mode contract, never a patient capture.
            snapshot.Synthetic = false; scene.Synthetic = false; scene.Profile = null;
            scene.NativeMachineId = "test-machine"; scene.PatientPosition = "HFS";
            scene.NominalCoordinatesSupported = true; scene.SupportedCoordinates = false;
            var catalog = new ClearPlan.Core.Collision.SourceCollisionModelCatalog {
                SchemaVersion = 1, Units = "mm",
                MachineBindings = new System.Collections.Generic.List<ClearPlan.Core.Collision.SourceCollisionMachineBinding> {
                    new ClearPlan.Core.Collision.SourceCollisionMachineBinding { MachineId = "test-machine", ModelId = "test-head" } },
                Models = new System.Collections.Generic.List<ClearPlan.Core.Collision.SourceCollisionModel> {
                    new ClearPlan.Core.Collision.SourceCollisionModel { ModelId = "test-head", Kind = "TrueBeamHeadSource",
                        SourcePath = "synthetic-unit-test", SourceSha256 = new string('a',64), SourceSnapshotDate = "2026-09-13",
                        EvidenceLevel = "source-recorded-measurement", MeasurementDate = "2026-07-30", Evidence = "Test fixture only, not a native capture.",
                        ReachRadiusMm = 700, WarningMarginMm = 30, SourceAngularMarginDegrees = 5,
                        HeadObstacles = new System.Collections.Generic.List<ClearPlan.Core.Collision.SourceHeadObstacle> {
                            new ClearPlan.Core.Collision.SourceHeadObstacle { FrontFaceFromIsoMm = 420, RadiusMm = 370 },
                            new ClearPlan.Core.Collision.SourceHeadObstacle { FrontFaceFromIsoMm = 400, RadiusMm = 200 } } } } };
            ClearPlan.Core.Collision.SourceCollisionSceneAdapter.Attach(scene, catalog, System.Threading.CancellationToken.None);
            TestAssert.Equal("unavailable", scene.SourceScreening.Status);
            TestAssert.False(scene.SourceScreening.ClinicalClearanceSupported);
            TestAssert.False(scene.SupportedCoordinates);
            var model = new ClearPlan.Presentation.ViewModels.CollisionViewModel(snapshot);
            model.SetScene(scene, null); model.IncludeInReport = true;
            TestAssert.True(model.SourceEnvelopeDrawable);
            TestAssert.False(model.DeviceEnvelopeDrawable);
            ClearPlan.Presentation.Views.CollisionView.CaptureReportIllustration(model);
            TestAssert.True(snapshot.CollisionPreviewPng.Length > 1000);
            TestAssert.True(snapshot.CollisionPreviewCaption.Contains("source-model envelope, nominal pose"));
            TestAssert.True(snapshot.CollisionPreviewCaption.Contains("does not confirm pitch/roll"));
            // Both catalogs may be configured; evidence and headline follow what is drawn.
            scene.Profile = new ClearPlan.Core.Collision.CollisionProfile {
                MachineId = "test-machine", Kind = "CArmSphere", Revision = "test-revision", Commissioned = true,
                CommissionedBy = "test-operator", Evidence = "commissioned-test-evidence", HeadCenterFromIsoMm = 500, HeadRadiusMm = 150 };
            model.SetScene(scene, null); model.StoreIllustration(new byte[] { 1 });
            TestAssert.True(model.SourceEnvelopeDrawable);
            TestAssert.True(model.ProfileEvidence.Contains("Test fixture only"));
            TestAssert.True(model.IllustrationStatus.StartsWith("Quellmodell"));
            scene.SupportedCoordinates = true;
            model.SetScene(scene, null); model.StoreIllustration(new byte[] { 1 });
            TestAssert.True(model.DeviceEnvelopeDrawable);
            TestAssert.False(model.SourceEnvelopeDrawable);
            TestAssert.Equal("commissioned-test-evidence", model.ProfileEvidence);
            TestAssert.True(model.IllustrationStatus.StartsWith("Gesamtbefund"));
            TestAssert.False(snapshot.CollisionPreviewCaption.Contains("Head extent clipped"));
            var serialized = ClearPlan.Core.Review.ReviewSnapshotJson.Serialize(snapshot);
            TestAssert.False(serialized.Contains("SourceModel") || serialized.Contains("SourceScreening"));
            model.ZeroPitchRollConfirmed = true;
            TestAssert.False(model.HasScene, "Confirmation requires fresh native capture, not an edit to existing results.");
        }
        public static void HeadlessIllustrationUsesTheGuiSceneWithoutAWindow()
        {
            var snapshot = ClearPlan.Core.Simulation.SyntheticScenarioFactory.Create("baseline-pass");
            var model = new ClearPlan.Presentation.ViewModels.CollisionViewModel(snapshot);
            model.ReloadCommand.Execute(null);
            model.IncludeInReport = true;
            var capture = typeof(ClearPlan.Presentation.Views.CollisionView).GetMethod("CaptureReportIllustration");
            TestAssert.NotNull(capture, "Headless reports must render the same detached 3D scene as the GUI without opening a window.");
            capture.Invoke(null, new object[] { model });
            TestAssert.True(snapshot.CollisionPreviewPng != null && snapshot.CollisionPreviewPng.Length > 1000);
            using (var stream = new System.IO.MemoryStream(snapshot.CollisionPreviewPng))
            {
                var bitmap = System.Windows.Media.Imaging.BitmapFrame.Create(stream);
                TestAssert.True(bitmap.PixelWidth >= 1000 && bitmap.PixelHeight >= 500);
                TestAssert.True(bitmap.PixelWidth * (long)bitmap.PixelHeight <= 4100000);
            }
            capture.Invoke(null, new object[] { model });
            using (var stream = new System.IO.MemoryStream(snapshot.CollisionPreviewPng))
            {
                var bitmap = new System.Windows.Media.Imaging.FormatConvertedBitmap(
                    System.Windows.Media.Imaging.BitmapFrame.Create(stream), System.Windows.Media.PixelFormats.Bgra32, null, 0);
                var pixels = new byte[bitmap.PixelWidth * bitmap.PixelHeight * 4];
                bitmap.CopyPixels(pixels, bitmap.PixelWidth * 4, 0);
                int whiteText = 0;
                for (int y = 15; y < 100; y++) for (int x = 20; x < bitmap.PixelWidth - 20; x++)
                {
                    int offset = (y * bitmap.PixelWidth + x) * 4;
                    if (pixels[offset] > 225 && pixels[offset + 1] > 225 && pixels[offset + 2] > 225) whiteText++;
                }
                TestAssert.True(whiteText > 100, "Repeated no-window capture must retain the 2D heading above the 3D geometry.");
            }
            model.SetFailure("No native geometry available.");
            capture.Invoke(null, new object[] { model });
            TestAssert.Equal(null, snapshot.CollisionPreviewPng, "A failed capture must not reuse an old illustration.");
        }
        public static void ReloadInvalidatesOldGeometryAndIllustration()
        {
            var snapshot = ClearPlan.Core.Simulation.SyntheticScenarioFactory.Create("baseline-pass");
            var model = new ClearPlan.Presentation.ViewModels.CollisionViewModel(snapshot);
            model.ReloadCommand.Execute(null); model.IncludeInReport = true;
            TestAssert.True(object.ReferenceEquals(model.BeamIds, model.BeamIds), "ItemsSource identity must remain stable between scene changes.");
            int notifications = 0;
            model.PropertyChanged += (s, e) => notifications++;
            model.SelectedBeamId = model.SelectedBeamId; model.PosePosition = model.PosePosition;
            TestAssert.Equal(0, notifications, "ComboBox/Slider writeback of unchanged values must not recursively notify bindings.");
            model.StoreIllustration(new byte[] { 1, 2, 3 });
            model.SetLoading();
            TestAssert.False(model.HasScene, "A pending refresh must not retain the previous collision scene.");
            TestAssert.Equal(null, model.Result);
            model.StoreIllustration(new byte[] { 4, 5, 6 });
            TestAssert.Equal(null, snapshot.CollisionPreviewPng);
        }
        public static void ReportIllustrationIsOptInAndDetached()
        {
            var snapshot = ClearPlan.Core.Simulation.SyntheticScenarioFactory.Create("baseline-pass");
            var model = new ClearPlan.Presentation.ViewModels.CollisionViewModel(snapshot);
            TestAssert.False(model.IncludeInReport);
            model.ReloadCommand.Execute(null);
            TestAssert.True(model.HasScene);
            model.StoreIllustration(new byte[] { 1, 2, 3 });
            var mapper = new ClearPlan.Reporting.ReviewSnapshotReportMapper();
            var image = typeof(ClearPlan.Reporting.ReviewReportDocument).GetProperty("CollisionPreviewPng");
            TestAssert.NotNull(image, "Optional collision illustration must reach the report renderer.");
            TestAssert.Equal(null, image.GetValue(mapper.Map(snapshot)));
            model.IncludeInReport = true;
            model.StoreIllustration(new byte[] { 1, 2, 3 });
            var report = mapper.Map(snapshot);
            TestAssert.Equal(3, ((byte[])image.GetValue(report)).Length);
            snapshot.CollisionPreviewPng[0] = 99;
            TestAssert.Equal((byte)1, ((byte[])image.GetValue(report))[0]);
            model.PosePosition = 1;
            TestAssert.Equal(null, snapshot.CollisionPreviewPng, "A changed pose must not export the previous illustration while rendering is pending.");
            var json = ReviewSnapshotJson.Serialize(snapshot);
            TestAssert.False(json.Contains("CollisionPreview") || json.Contains("CollisionScene"));
            TestAssert.Throws<ArgumentException>(() => model.SetScene(
                ClearPlan.Core.Collision.SyntheticCollisionFactory.Create("another-plan"), null));
        }
        public static void SettingsAndMemoryBoundary()
        {
            var nativeSnapshot = ClearPlan.Core.Simulation.SyntheticScenarioFactory.Create("baseline-pass");
            nativeSnapshot.Synthetic = false;
            var model = new ClearPlan.Presentation.ViewModels.CollisionViewModel(nativeSnapshot);
            var confirmation = model.GetType().GetProperty("ZeroPitchRollConfirmed");
            TestAssert.NotNull(confirmation, "Unknown native pitch/roll requires an explicit per-plan operator confirmation.");
            TestAssert.Equal(false, confirmation.GetValue(model));
            var sceneForConfirmation = ClearPlan.Core.Collision.SyntheticCollisionFactory.Create(nativeSnapshot.ActivePlanKey);
            sceneForConfirmation.Synthetic = false;
            model.SetScene(sceneForConfirmation, new ClearPlan.Core.Collision.CollisionResult { Status = "unavailable" });
            model.IncludeInReport = true; model.StoreIllustration(new byte[] { 1 });
            sceneForConfirmation.SupportedCoordinates = false;
            model.StoreIllustration(new byte[] { 1 });
            TestAssert.False(model.DeviceEnvelopeDrawable);
            TestAssert.True(nativeSnapshot.CollisionPreviewCaption.Contains("device envelope not displayed"),
                "The caption must not claim a hidden/unconfirmed device model was drawn.");
            confirmation.SetValue(model, true);
            TestAssert.False(model.HasScene, "Changing coordinate confirmation must invalidate captured geometry.");
            TestAssert.Equal(null, nativeSnapshot.CollisionPreviewPng);
            var path = typeof(ClearPlanPathOptions).GetProperty("CollisionProfilesJsonPath");
            TestAssert.NotNull(path, "Collision profiles need an editable settings.ini path.");
            var source = new ClearPlanSettingsModel();
            path.SetValue(source.Paths, "MachineGeometry/collision-profiles.json");
            var restored = new ClearPlanSettingsModel();
            PathSettingsIni.Apply(restored, PathSettingsIni.Serialize(source));
            TestAssert.Equal(path.GetValue(source.Paths), path.GetValue(restored.Paths));
            source.Paths.SourceCollisionModelsJsonPath = "MachineGeometry/source-screening.json";
            PathSettingsIni.Apply(restored, PathSettingsIni.Serialize(source));
            TestAssert.Equal(source.Paths.SourceCollisionModelsJsonPath, restored.Paths.SourceCollisionModelsJsonPath);
            var scene = typeof(ReviewSnapshot).GetProperty("CollisionScene");
            TestAssert.NotNull(scene, "Read-only collision scene must be attached to the active snapshot.");
            TestAssert.True(Attribute.IsDefined(scene, typeof(Newtonsoft.Json.JsonIgnoreAttribute)),
                "Patient surface geometry must not enter shareable review JSON.");
        }
    }
}
