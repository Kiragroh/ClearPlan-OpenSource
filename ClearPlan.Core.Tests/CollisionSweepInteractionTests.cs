using System;
using System.Collections;
using System.IO;
using System.Reflection;
using System.Windows.Input;
using ClearPlan.Core.Collision;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Core.Tests
{
    internal static class CollisionSweepInteractionTests
    {
        public static void BothAngleCirclesClearAndHalcyonNeverShowsCArm()
        {
            var model=Create(); model.Scene.PatientPosition="HFS";
            var view=new ClearPlan.Presentation.Views.CollisionView { DataContext=model };
            view.Measure(new System.Windows.Size(1340,800)); view.Arrange(new System.Windows.Rect(0,0,1340,800)); view.UpdateLayout();
            var draw=view.GetType().GetMethod("DrawScene",BindingFlags.NonPublic|BindingFlags.Instance);
            draw.Invoke(view,null);
            var gantry=view.FindName("GantryCompass") as System.Windows.Controls.Image;
            TestAssert.NotNull(gantry,"Gantry needs its own circular diagram.");
            var couch=(System.Windows.Controls.Image)view.FindName("RoomCompass");
            TestAssert.NotNull(gantry.Source); TestAssert.NotNull(couch.Source);
            foreach(var width in new[]{1340d,1184d,980d})
            {
                view.Width=width; view.Height=620;
                view.Measure(new System.Windows.Size(width,620)); view.Arrange(new System.Windows.Rect(0,0,width,620)); view.UpdateLayout();
                draw.Invoke(view,null); view.UpdateLayout();
                TestAssert.Equal(width,view.ActualWidth,"Compact layout regression must use the requested actual viewport width.");
                var illustration=(System.Windows.FrameworkElement)view.FindName("IllustrationRoot");
                var scene=(System.Windows.FrameworkElement)view.FindName("SceneViewport");
                var rail=(System.Windows.FrameworkElement)view.FindName("RoomReferenceRoot");
                var sceneBounds=scene.TransformToAncestor(illustration).TransformBounds(new System.Windows.Rect(0,0,scene.ActualWidth,scene.ActualHeight));
                var railBounds=rail.TransformToAncestor(illustration).TransformBounds(new System.Windows.Rect(0,0,rail.ActualWidth,rail.ActualHeight));
                TestAssert.True(sceneBounds.Right<=railBounds.Left,"Orientation must occupy a separate right column, never overlap the patient viewport at width "+width);
                TestAssert.True(railBounds.Right<=illustration.ActualWidth && railBounds.Bottom<=illustration.ActualHeight,"Orientation rail must remain inside the available compact surface.");
                TestAssert.True(gantry.ActualHeight>35 && couch.ActualHeight>35,"Both circles must remain readable in the bounded right rail.");
            }
            model.Scene.Profile.Kind="RingBore"; draw.Invoke(view,null);
            TestAssert.NotNull(gantry.Source,"A bore retains nominal angle circles.");
            TestAssert.Equal(0,((System.Windows.Controls.Viewport3D)view.FindName("RoomReferenceViewport")).Children.Count,
                "A bore must not receive a TrueBeam C-arm schematic.");
            model.Scene.Profile=null;
            model.Scene.SourceModel=new SourceCollisionModel {Kind="HalcyonBoreSource",BoreWarningRadiusMm=90,BoreLimitRadiusMm=100};
            draw.Invoke(view,null);
            TestAssert.NotNull(gantry.Source); TestAssert.NotNull(couch.Source);
            TestAssert.Equal(0,((System.Windows.Controls.Viewport3D)view.FindName("RoomReferenceViewport")).Children.Count,
                "Native Halcyon source geometry must not receive a TrueBeam C-arm schematic.");
            var staleProfile=Create().Scene.Profile; staleProfile.HeadRadiusMm=-1;
            model.Scene.Profile=staleProfile;
            TestAssert.False(model.UsesProfileGeometry,"Conflicting test profile must be unusable.");
            draw.Invoke(view,null);
            TestAssert.Equal(0,((System.Windows.Controls.Viewport3D)view.FindName("RoomReferenceViewport")).Children.Count,
                "A stale C-arm profile must not override Halcyon bore identity.");
            model.Scene.SourceModel=null; draw.Invoke(view,null);
            TestAssert.Equal(0,((System.Windows.Controls.Viewport3D)view.FindName("RoomReferenceViewport")).Children.Count,
                "An unusable C-arm profile alone must not fabricate a machine reference.");
            model.Scene.Profile=Create().Scene.Profile; model.Scene.Profile.Kind="RingBore";
            model.Scene.SourceModel=new SourceCollisionModel {Kind="TrueBeamHeadSource"};
            TestAssert.True(model.UsesProfileGeometry);
            draw.Invoke(view,null);
            TestAssert.Equal(0,((System.Windows.Controls.Viewport3D)view.FindName("RoomReferenceViewport")).Children.Count,
                "An active ring profile must veto a conflicting C-arm source reference.");
            model.SetLoading(); draw.Invoke(view,null);
            TestAssert.True(gantry.Source==null && couch.Source==null,"Loading must clear both old angle diagrams.");
        }

        public static void RoomReferenceHasFixedStandAndTurnsOnlyPatientAxis()
        {
            var method = typeof(ClearPlan.Presentation.Views.CollisionView).GetMethod("CreateRoomReference", BindingFlags.Static | BindingFlags.NonPublic);
            TestAssert.NotNull(method);
            var origin = new ClearPlan.Core.PlanAnalysis.BeamPoint3D(0,0,0);
            var a = new CollisionPose { Isocenter = origin, Source = new ClearPlan.Core.PlanAnalysis.BeamPoint3D(0,-1000,0) };
            var b = new CollisionPose { Isocenter = origin, Source = new ClearPlan.Core.PlanAnalysis.BeamPoint3D(0,-1000,0), PatientSupportAngleDegrees = 300 };
            var first = (System.Windows.Media.Media3D.Model3DGroup)method.Invoke(null,new object[] { a });
            var second = (System.Windows.Media.Media3D.Model3DGroup)method.Invoke(null,new object[] { b });
            for (int i=2; i<9; i++) TestAssert.Equal(first.Children[i].Bounds,second.Children[i].Bounds,"G0 machine reference must remain fixed when couch rotates.");
            TestAssert.False(first.Children[9].Bounds.Equals(second.Children[9].Bounds),"Patient-axis schematic must show couch rotation.");
            TestAssert.Equal(300d,b.PatientSupportAngleDegrees,"Display reference must not modify native pose.");
            var model=Create();
            model.Scene.PatientPosition="HFS";
            var view=new ClearPlan.Presentation.Views.CollisionView { DataContext=model };
            view.Measure(new System.Windows.Size(1340,800)); view.Arrange(new System.Windows.Rect(0,0,1340,800)); view.UpdateLayout();
            var draw=view.GetType().GetMethod("DrawScene",BindingFlags.NonPublic|BindingFlags.Instance);
            draw.Invoke(view,null);
            var compass=(System.Windows.Controls.Image)view.FindName("RoomCompass");
            TestAssert.NotNull(compass.Source,"C-arm room reference should have a current compass.");
            var savedScene=model.Scene;
            model.SetFailure("Test invalidation"); draw.Invoke(view,null);
            TestAssert.True(compass.Source==null,"Invalid geometry must not keep previous plan angles visible.");
            TestAssert.Equal(System.Windows.Visibility.Collapsed,((System.Windows.FrameworkElement)view.FindName("RoomReferenceRoot")).Visibility);
            model.SetScene(savedScene,null); draw.Invoke(view,null);
            TestAssert.NotNull(compass.Source);
            model.SetLoading(); draw.Invoke(view,null);
            TestAssert.True(compass.Source==null,"Loading must clear old room angles before the next capture arrives.");
        }

        public static void BodyOnlyClearAndMinimumStartRemainScoped()
        {
            var model = Create();
            model.Scene.Surfaces.RemoveAll(s => s.Role == "Support");
            var sweep = new CollisionSweepResult();
            foreach (var distance in new[] { 30d, 4d, double.NaN })
                sweep.Frames.Add(new CollisionSweepFrame { Pose = new CollisionPose { BeamId = distance == 4 ? "SECOND" : "FIRST" },
                    ModelBodyStatus = "model-clear", ModelTableStatus = "uncertain", Status = "uncertain",
                    BodyLowerBoundMm = distance, BodyClearanceMm = distance + 1 });
            model.Scene.Poses = new System.Collections.Generic.List<CollisionPose> { sweep.Frames[0].Pose, sweep.Frames[1].Pose };
            model.SetScene(model.Scene, null, sweep);
            TestAssert.Equal("SECOND", model.SelectedBeamId, "Start at global minimum, not the first field.");
            TestAssert.Equal("body-clear", model.SelectedFrame.StatusCode);
            TestAssert.Equal("Still clear", model.SelectedFrame.StatusText);
            TestAssert.Equal("#F2C45C", model.SelectedFrame.StatusColor, "Still clear is yellow, never full-clear green.");
            TestAssert.Equal("#FFF5DD", model.SelectedFrame.StatusSurface);
            var full = new CollisionSweepFrameViewModel(new CollisionSweepFrame { ModelBodyStatus="model-clear", ModelTableStatus="model-clear" });
            TestAssert.Equal("Full clear", full.StatusText);
            TestAssert.Equal("#69D391", full.StatusColor);
            TestAssert.Equal("#EAF6EE", full.StatusSurface);
            TestAssert.Equal("uncertain", model.SelectedFrame.Frame.Status, "Clinical evidence gate remains unchanged.");
            model.Scene.Surfaces.Add(new CollisionSurface { Role = "Support" }); // supplied but invalid is NOT absent
            model.SetScene(model.Scene, null, sweep);
            TestAssert.Equal("partial-clear", model.SelectedFrame.StatusCode);
            TestAssert.Equal("Still clear", model.SelectedFrame.StatusText);
            TestAssert.Equal("#F2C45C", model.SelectedFrame.StatusColor);
            sweep.Frames[1].ModelTableStatus = "model-hit";
            model.SetScene(model.Scene, null, sweep);
            TestAssert.Equal("model-hit", model.SelectedFrame.StatusCode);
            TestAssert.Equal("#FF7B72", model.SelectedFrame.StatusColor);
            model.Scene.Surfaces.RemoveAll(s => s.Role == "Support");
            sweep.Frames[1].ModelBodyStatus = "uncertain"; sweep.Frames[1].ModelTableStatus = "uncertain";
            model.SetScene(model.Scene, null, sweep);
            TestAssert.Equal("uncertain", model.SelectedFrame.StatusCode, "Missing table must not rescue uncertain body.");
            sweep.Frames[1].RadialReview = new CollisionRadialReview { BodyStatus="clear", TableStatus="uncertain", Status="uncertain", BodyClearanceMm=-4 };
            model.SetScene(model.Scene,null,sweep);
            // Status and numeric evidence must agree: an invalid clear label cannot rescue a negative gap.
            TestAssert.Equal("uncertain",model.SelectedFrame.StatusCode);
            sweep.Frames[1].RadialReview.BodyClearanceMm=1;
            TestAssert.Equal("body-clear",model.SelectedFrame.StatusCode,"Open-bore review may also be explicitly body-only.");
            TestAssert.Equal("uncertain",sweep.Frames[1].Status);
            sweep.Frames[1].RadialReview.BodyStatus="hit";
            sweep.Frames[1].RadialReview.Status="hit";
            TestAssert.Equal("hit",model.SelectedFrame.StatusCode);
            sweep.Frames[1].RadialReview.BodyStatus="clear";
            sweep.Frames[1].RadialReview.Status="model-hit";
            TestAssert.Equal("model-hit",model.SelectedFrame.StatusCode,"Native radial model-hit has precedence over body-only clear.");
            sweep.Frames[1].RadialReview.Status="uncertain"; sweep.Frames[1].RadialReview.TableStatus="model-hit";
            TestAssert.Equal("model-hit",model.SelectedFrame.StatusCode);
        }

        public static void PlaybackCrossesFieldsAndCanRepeat()
        {
            var model = Create();
            var second = SyntheticCollisionFactory.Create(model.Scene.PlanKey);
            foreach (var pose in second.Poses) { pose.BeamId = "SECOND"; model.Scene.Poses.Add(pose); }
            model.SetScene(model.Scene, null);
            model.PosePosition = model.MaximumPosePosition - 1;
            model.PlayPauseCommand.Execute(null);
            model.AdvancePlayback(); model.AdvancePlayback();
            TestAssert.Equal("SECOND", model.SelectedBeamId, "Playback must continue to the next field.");
            TestAssert.Equal(0, model.PosePosition);
            TestAssert.True(model.IsPlaying);
            var repeat = typeof(CollisionViewModel).GetProperty("RepeatPlayback");
            TestAssert.NotNull(repeat); repeat.SetValue(model, true);
            for (int i = 0; i < 37; i++) model.AdvancePlayback();
            TestAssert.Equal("SYNTHETIC ARC", model.SelectedBeamId);
            TestAssert.Equal(0, model.PosePosition); TestAssert.True(model.IsPlaying);
            model.SelectedBeamId = "SECOND";
            TestAssert.False(model.IsPlaying, "Manual field changes must stop playback.");
            model.PlayPauseCommand.Execute(null); model.SetFailure("No geometry");
            TestAssert.False(model.IsPlaying);
        }

        public static void MlcInsetKeepsCapturedCpIdentityAndClearsMissingData()
        {
            var model = Create();
            var cpProperty = typeof(CollisionViewModel).GetProperty("MlcControlPoint");
            TestAssert.NotNull(cpProperty, "Collision MLC needs an explicitly captured CP, not the renumbered sweep index.");
            var snapshot = (ClearPlan.Core.Review.ReviewSnapshot)typeof(CollisionViewModel)
                .GetField("snapshot", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(model);
            var beam = new ClearPlan.Core.PlanAnalysis.ReviewBeamAnalysis { BeamId = model.SelectedBeamId };
            foreach (var pose in model.Scene.Poses) beam.ControlPoints.Add(new ClearPlan.Core.PlanAnalysis.ReviewControlPointSample {
                Index = pose.ControlPointIndex, GantryAngleDegrees = pose.GantryDegrees, PatientSupportAngleDegrees = pose.PatientSupportAngleDegrees });
            snapshot.PlanAnalysis = new ClearPlan.Core.PlanAnalysis.ReviewPlanAnalysis(); snapshot.PlanAnalysis.Beams.Add(beam);
            model.PosePosition = 3; // 30 degrees = captured CP 6, not sampled row 3.
            var cp = (ClearPlan.Core.PlanAnalysis.ReviewControlPointSample)cpProperty.GetValue(model);
            TestAssert.NotNull(cp); TestAssert.Equal(6, cp.Index);
            beam.ControlPoints.Remove(cp);
            TestAssert.True(cpProperty.GetValue(model) == null, "Missing CP must not borrow another aperture.");
            model.SetLoading(); TestAssert.True(cpProperty.GetValue(model) == null);
        }

        public static void PlaybackPauseRemainsAvailableOnStaticField()
        {
            var model = Create();
            model.Scene.Poses.Add(new CollisionPose { BeamId="STATIC", ControlPointIndex=0, GantryDegrees=0,
                Source=model.Scene.Poses[0].Source, Isocenter=model.Scene.Poses[0].Isocenter });
            model.SetScene(model.Scene,null); model.SelectedBeamId="STATIC"; model.RepeatPlayback=true;
            model.PlayPauseCommand.Execute(null);
            for(int i=0;i<80 && model.SelectedBeamId != "STATIC";i++) model.AdvancePlayback();
            TestAssert.Equal("STATIC",model.SelectedBeamId,"Playback must reach the static field within one sweep.");
            model.AllFieldsPlayback=false;
            TestAssert.True(model.PlayPauseCommand.CanExecute(null),"Pause must remain usable in a one-state loop.");
            model.PlayPauseCommand.Execute(null); TestAssert.False(model.IsPlaying);
        }

        public static void MlcInsetRejectsChangedIsocenterAndNativeGeometryStamp()
        {
            var model=Create();
            model.PosePosition=0; // This test binds CP0; default entry now selects the minimum distance.
            var snapshot=(ClearPlan.Core.Review.ReviewSnapshot)typeof(CollisionViewModel)
                .GetField("snapshot",BindingFlags.NonPublic | BindingFlags.Instance).GetValue(model);
            var cp=new ClearPlan.Core.PlanAnalysis.ReviewControlPointSample { Index=0, GantryAngleDegrees=0,
                IsocenterMm=new[]{12d,0d,0d} };
            var beam=new ClearPlan.Core.PlanAnalysis.ReviewBeamAnalysis { BeamId=model.SelectedBeamId };
            beam.ControlPoints.Add(cp); snapshot.PlanAnalysis=new ClearPlan.Core.PlanAnalysis.ReviewPlanAnalysis(); snapshot.PlanAnalysis.Beams.Add(beam);
            TestAssert.True(model.MlcControlPoint==null,"Same field and angles do not permit a different isocenter.");
            cp.IsocenterMm=new[]{0d,0d,0d}; snapshot.Synthetic=false; model.Scene.Synthetic=false;
            TestAssert.True(model.MlcControlPoint==null,"Native aperture cannot be paired without matching capture stamps.");
            var stamp=typeof(CollisionScene).GetProperty("NativePlanFingerprint");
            var stamps=typeof(CollisionScene).GetProperty("NativeBeamFingerprints");
            TestAssert.NotNull(stamp); TestAssert.NotNull(stamps);
            snapshot.PlanAnalysis.NativePlanFingerprint="PLAN"; beam.NativeGeometryFingerprint="BEAM";
            stamp.SetValue(model.Scene,"PLAN");
            stamps.SetValue(model.Scene,new System.Collections.Generic.Dictionary<string,string>{{beam.BeamId,"BEAM"}});
            TestAssert.True(ReferenceEquals(cp,model.MlcControlPoint));
            beam.NativeGeometryFingerprint="CHANGED"; TestAssert.True(model.MlcControlPoint==null);
        }

        public static void ApertureInsetUsesPhysicalLayersWithoutDrr()
        {
            var renderer = typeof(ClearPlan.Rendering.BeamEyeViewRenderer).GetMethod("RenderApertureInset");
            TestAssert.NotNull(renderer, "Small MLC panel needs a legible aperture-only renderer.");
            foreach (int layers in new[] { 1, 2 })
            {
                var aperture = new ClearPlan.Core.PlanAnalysis.ApertureGeometry();
                aperture.FixedBoundingBox = new ClearPlan.Core.PlanAnalysis.ApertureRectangle(-140,-140,140,140);
                for (int i=0;i<layers;i++) aperture.Layers.Add(new ClearPlan.Core.PlanAnalysis.ApertureLayer {
                    Label = "Layer " + i, LeafBoundariesMm = new[] {-40d,0d,40d},
                    Bank1PositionsMm = new[] {-30d-i*5,-45d}, Bank2PositionsMm = new[] {30d,45d+i*5} });
                var beam = new ClearPlan.Core.PlanAnalysis.ReviewBeamAnalysis { MlcLayerCount = layers, HasJaws = false, GeometryStatus = "available" };
                var cp = new ClearPlan.Core.PlanAnalysis.ReviewControlPointSample { Aperture = aperture };
                var bytes = (byte[])renderer.Invoke(null,new object[] {beam,cp,360});
                using (var stream = new MemoryStream(bytes)) using(var image = new System.Drawing.Bitmap(stream))
                { TestAssert.Equal(360,image.Width); TestAssert.Equal(360,image.Height);
                  int green=0,blue=0;
                  for(int y=0;y<360;y++) for(int x=0;x<360;x++) { var p=image.GetPixel(x,y); if(p.G>p.R*1.4 && p.G>p.B*1.2) green++; if(p.B>p.R*1.3 && p.B>p.G*1.1) blue++; }
                  TestAssert.True(green>80,"Physical/effective MLC must be visible.");
                  if(layers==2) TestAssert.True(blue>30,"Second physical layer must be blue."); }
            }
        }

        public static void PlaybackStepsCachedFramesAndStopsAtTheEnd()
        {
            var model = Create();
            model.PosePosition = 0;
            var frames = (IList)Property(model, "BeamFrames");
            TestAssert.Equal(37, frames.Count);
            var sweep = Property(model, "Sweep");
            Command(model, "PlayPauseCommand").Execute(null);
            TestAssert.Equal(true, Property(model, "IsPlaying"));
            Method(model, "AdvancePlayback").Invoke(model, null);
            TestAssert.Equal(1, model.PosePosition);
            TestAssert.True(ReferenceEquals(sweep, Property(model, "Sweep")), "Playback must not recalculate geometry.");
            for (int i = 0; i < 50; i++) Method(model, "AdvancePlayback").Invoke(model, null);
            TestAssert.Equal(model.MaximumPosePosition, model.PosePosition);
            TestAssert.Equal(false, Property(model, "IsPlaying"));
            Command(model, "PreviousPoseCommand").Execute(null);
            TestAssert.Equal(model.MaximumPosePosition - 1, model.PosePosition);
        }

        public static void ContextChangesStopPlaybackAndClearExportFrames()
        {
            var model = Create();
            model.Scene.SourceModel = new SourceCollisionModel { ModelId = "SHOULD-NOT-REPLACE-PROFILE" };
            string provenance = model.ProfileText;
            model.EnvelopeVisible = false;
            TestAssert.Equal(provenance, model.ProfileText, "Visibility must not change distance provenance.");
            model.IncludeInReport = true; model.StoreIllustration(new byte[] { 1 });
            TestAssert.False(model.IllustrationPng == null);
            // Caption evidence must describe the same effective model when the envelope is hidden.
            var captionField = typeof(CollisionViewModel).GetField("snapshot", BindingFlags.NonPublic | BindingFlags.Instance);
            var sourceSnapshot = (ClearPlan.Core.Review.ReviewSnapshot)captionField.GetValue(model);
            TestAssert.False(sourceSnapshot.CollisionPreviewCaption.Contains("SHOULD-NOT-REPLACE-PROFILE"));
            Command(model, "PlayPauseCommand").Execute(null);
            model.SetLoading();
            TestAssert.Equal(false, Property(model, "IsPlaying"));
            TestAssert.Equal(0, ((IList)Property(model, "BeamFrames")).Count);
            TestAssert.False(Command(model, "PlayPauseCommand").CanExecute(null));
            model.SetFailure("No geometry");
            TestAssert.Equal(false, Property(model, "IsPlaying"));
        }

        public static void ViewIncludesPlaybackGhostsAndStatusDistanceTable()
        {
            string root = FindRoot();
            string xaml = File.ReadAllText(Path.Combine(root, "ClearPlan.Presentation/Views/CollisionView.xaml"));
            string view = File.ReadAllText(Path.Combine(root, "ClearPlan.Presentation/Views/CollisionView.xaml.cs"));
            foreach (string binding in new[] { "PlayPauseCommand", "PreviousPoseCommand", "NextPoseCommand", "AllFrames", "Frame.Pose.BeamId", "GhostsVisible" })
                TestAssert.True(xaml.Contains(binding), "Collision GUI missing " + binding);
            TestAssert.True(xaml.Contains("BodyDistanceText") && xaml.Contains("TableDistanceText"));
            TestAssert.True(view.Contains("DispatcherTimer") && view.Contains("StopPlayback"));
            TestAssert.True(view.Contains("CaptureReportSweep"), "Reports need each beam overview, not only the last GUI frame.");
            TestAssert.True(view.Contains("BoreOutline"), "A stationary opaque bore must use solid edges so it does not hide the patient.");
            var outline = typeof(ClearPlan.Presentation.Views.CollisionView).GetMethod("SourceHeadOutline", BindingFlags.NonPublic | BindingFlags.Static);
            TestAssert.NotNull(outline, "Moving source-model ghosts need boundary contours, not stacked filled head volumes.");
            var front = (System.Windows.Media.Media3D.MeshGeometry3D)outline.Invoke(null, new object[] {
                new System.Windows.Media.Media3D.Point3D(), new System.Windows.Media.Media3D.Vector3D(0,0,1), 400.0, 100.0, 700.0, null, true });
            TestAssert.True(front.TriangleIndices.Count > 0);
            foreach (var point in front.Positions)
            {
                TestAssert.True(Math.Abs(point.Z - 400) < 0.001, "Ghost contour belongs to the captured front-face plane.");
                TestAssert.True(Math.Sqrt(point.X*point.X + point.Y*point.Y) > 96,
                    "No filled front-face triangles may cover the anatomy between the contours.");
            }
            var nominal = new CollisionSweepFrameViewModel(new CollisionSweepFrame { Pose = new CollisionPose(),
                Status = "uncertain", BodyLowerBoundMm = 200, BodyClearanceMm = 203 });
            var availability = typeof(CollisionSweepFrameViewModel).GetProperty("EvidenceText");
            TestAssert.NotNull(availability);
            TestAssert.True(((string)availability.GetValue(nominal)).Contains("Tisch: Abstand nicht verfügbar"),
                "A missing distance alone does not prove the surface was not captured.");
            TestAssert.Equal("uncertain", nominal.StatusCode, "Available body distance cannot approve a missing table.");
            var display = typeof(ClearPlan.Presentation.Views.CollisionView).GetMethod("RoomDisplayTransform", BindingFlags.NonPublic | BindingFlags.Static);
            TestAssert.NotNull(display, "Room view must rotate patient and native source together, not apply couch twice.");
            var fixturePose = new CollisionPose { Isocenter = new ClearPlan.Core.PlanAnalysis.BeamPoint3D(10,20,30), PatientSupportAngleDegrees = 300 };
            var transform = (System.Windows.Media.Media3D.Transform3D)display.Invoke(null, new object[] { fixturePose });
            var source90 = new System.Windows.Media.Media3D.Point3D(510,20,30 + Math.Sqrt(3)*500);
            var roomSource = transform.Transform(source90);
            TestAssert.True(Math.Abs(roomSource.X - 1010) < 1e-6 && Math.Abs(roomSource.Y - 20) < 1e-6 && Math.Abs(roomSource.Z - 30) < 1e-6);
            var bodyPoint = new System.Windows.Media.Media3D.Point3D(50,60,70);
            TestAssert.True(Math.Abs((transform.Transform(bodyPoint)-roomSource).Length - (bodyPoint-source90).Length) < 1e-6,
                "Display transform cannot change patient/head distances.");
            TestAssert.True(view.Contains("OrthographicCamera") && view.Contains("DrawTrajectory") && xaml.Contains("MouseWheel"),
                "Readable review needs fitted room geometry, trajectory markers and direct zoom.");
            var samePose = typeof(ClearPlan.Presentation.Views.CollisionView).GetMethod("SameEnvelopePose", BindingFlags.NonPublic | BindingFlags.Static);
            var a = new CollisionPose { GantryDegrees = 0, Isocenter = new ClearPlan.Core.PlanAnalysis.BeamPoint3D(), Source = new ClearPlan.Core.PlanAnalysis.BeamPoint3D(0,-1000,0) };
            var b = new CollisionPose { GantryDegrees = 0, Isocenter = new ClearPlan.Core.PlanAnalysis.BeamPoint3D(), Source = new ClearPlan.Core.PlanAnalysis.BeamPoint3D(1000,0,0) };
            TestAssert.Equal(false, samePose.Invoke(null, new object[] { a,b }), "Distinct captured source directions must not disappear.");
            TestAssert.False(view.Contains("meshes.Add(LeadingBoundary"), "An open-bore radial review must not display an axial end wall.");
            var leading = typeof(ClearPlan.Presentation.Views.CollisionView).GetMethod("LeadingBoundary", BindingFlags.NonPublic | BindingFlags.Static);
            var disk = (System.Windows.Media.Media3D.MeshGeometry3D)leading.Invoke(null, new object[] { new System.Windows.Media.Media3D.Point3D(5,7,200), 100.0, false });
            foreach (var point in disk.Positions) TestAssert.Equal(200.0, point.Z, "Leading boundary must lie at the configured physical z plane.");
            TestAssert.True(disk.TriangleIndices.Count > 0);
            TestAssert.False(xaml.Contains("Negativ = Modellpenetration"), "Negative lower bounds alone do not prove overlap.");
            TestAssert.True(CollisionSweepFrameViewModel.Distance(null, 4.4).StartsWith("≤"), "Unknown lower bound must not look like exact clearance.");
        }

        public static void NativeCoverageComesFromTheImageNotTheReportCrop()
        {
            string source = File.ReadAllText(Path.Combine(FindRoot(), "ClearPlan.Script/Review/EsapiCollisionBuilder.cs"));
            TestAssert.True(source.Contains("CtCoverageKnown") && source.Contains("ExternalTruncatedAtCtBoundary"));
            TestAssert.True(source.Contains("plan.StructureSet.Image") && source.Contains("ZSize"));
            TestAssert.False(source.Contains("BeginModifications") || source.Contains("SaveModifications"));
        }

        public static void NativeLayoutKeepsLargeViewport()
        {
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass"); snapshot.Synthetic = false;
            var scene = SyntheticCollisionFactory.Create(snapshot.ActivePlanKey); scene.Synthetic = false; scene.Profile = null;
            scene.SourceModel = new SourceCollisionModel { ModelId = "synthetic-ring-source", Kind = "HalcyonBoreSource",
                BoreWarningRadiusMm = 90, BoreLimitRadiusMm = 100, LeadingWarningFromIsoMm = 190, LeadingLimitFromIsoMm = 200 };
            var model = new CollisionViewModel(snapshot); model.SetScene(scene, null);
            var view = new ClearPlan.Presentation.Views.CollisionView { DataContext = model };
            view.Measure(new System.Windows.Size(1184, 620));
            view.Arrange(new System.Windows.Rect(0, 0, 1184, 620)); view.UpdateLayout();
            var viewport = (System.Windows.FrameworkElement)view.FindName("SceneViewport");
            TestAssert.True(viewport.ActualWidth >= 500 && viewport.ActualHeight >= 300,
                "Native 3D viewport must remain at least 500 x 300, got " + viewport.ActualWidth + " x " + viewport.ActualHeight);
        }

        public static void PlaybackRetainsTheRoomCamera()
        {
            var model = Create();
            var view = new ClearPlan.Presentation.Views.CollisionView { DataContext = model };
            view.Measure(new System.Windows.Size(1340, 800));
            view.Arrange(new System.Windows.Rect(0, 0, 1340, 800)); view.UpdateLayout();
            var draw = view.GetType().GetMethod("DrawScene", BindingFlags.NonPublic | BindingFlags.Instance);
            var viewport = (System.Windows.Controls.Viewport3D)view.FindName("SceneViewport");
            draw.Invoke(view, null);
            var poseHeading=view.FindName("ScenePoseHeading") as System.Windows.Controls.TextBlock;
            TestAssert.NotNull(poseHeading,"Scene needs a rendered-pose heading, not a live binding racing ahead of geometry.");
            string renderedHeading=poseHeading.Text;
            model.PosePosition=model.PosePosition==0 ? 1 : 0;
            view.UpdateLayout();
            TestAssert.Equal(renderedHeading,poseHeading.Text,"Queued geometry must keep its matching heading until DrawScene renders the next pose.");
            draw.Invoke(view,null); view.UpdateLayout();
            TestAssert.Equal(model.PoseText,poseHeading.Text,"A completed draw must expose the same pose in scene and orientation.");
            var focus = (System.Windows.Media.Media3D.OrthographicCamera)viewport.Camera;
            model.RepeatPlayback=true; model.AllFieldsPlayback=false; model.PlayPauseCommand.Execute(null); draw.Invoke(view,null);
            var first = (System.Windows.Media.Media3D.OrthographicCamera)viewport.Camera;
            TestAssert.True(focus.Width < first.Width,"Manual minimum-position review must use a closer fit than the full angular sweep.");
            for (int position=0; position<36; position++)
            {
                model.AdvancePlayback(); draw.Invoke(view, null);
                var current = (System.Windows.Media.Media3D.OrthographicCamera)viewport.Camera;
                TestAssert.True(Math.Abs(current.Width-first.Width) < 1e-6, "Playback must not auto-zoom between gantry angles.");
                TestAssert.True((current.Position-first.Position).Length < 1e-6, "Playback must not pan the room camera.");
            }
            model.Zoom = 1.5; draw.Invoke(view, null);
            TestAssert.True(((System.Windows.Media.Media3D.OrthographicCamera)viewport.Camera).Width < first.Width,
                "Explicit zoom must still change the stable camera.");
            model.StopPlayback();
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            var scene = SyntheticCollisionFactory.Create(snapshot.ActivePlanKey); scene.Poses.Clear();
            var unavailable = new CollisionViewModel(snapshot); unavailable.SetScene(scene, null);
            var unavailableView = new ClearPlan.Presentation.Views.CollisionView { DataContext = unavailable };
            unavailableView.Measure(new System.Windows.Size(1340, 800));
            unavailableView.Arrange(new System.Windows.Rect(0, 0, 1340, 800)); unavailableView.UpdateLayout();
            draw.Invoke(unavailableView, null);
            var points = (IList)unavailableView.GetType().GetField("framingPoints", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(unavailableView);
            TestAssert.True(points.Count > 0, "Captured surfaces must remain framed when no pose can be evaluated.");
        }

        public static void WorstPoseUsesDisplayedRadialDistance()
        {
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            var scene = SyntheticCollisionFactory.Create(snapshot.ActivePlanKey);
            var a = scene.Poses[0]; a.BeamId = "A";
            var b = scene.Poses[1]; b.BeamId = "B";
            scene.Poses = new System.Collections.Generic.List<CollisionPose> { a, b };
            var sweep = new CollisionSweepResult { Frames = new System.Collections.Generic.List<CollisionSweepFrame> {
                new CollisionSweepFrame { Pose = a, BodyLowerBoundMm = -1000, RadialReview = new CollisionRadialReview { BodyClearanceMm = 90, TableClearanceMm = 50 } },
                new CollisionSweepFrame { Pose = b, BodyLowerBoundMm = 50, RadialReview = new CollisionRadialReview { BodyClearanceMm = -5, TableClearanceMm = 50 } } } };
            var model = new CollisionViewModel(snapshot); model.SetScene(scene, null, sweep);
            TestAssert.NotNull(typeof(CollisionViewModel).GetProperty("AllFrames"), "GUI table must contain all fields, not just the active beam.");
            model.SelectedFrame = model.BeamFrames[0];
            model.WorstPoseCommand.Execute(null);
            TestAssert.Equal("B", model.SelectedBeamId, "Smallest-distance navigation must use displayed radial distance, not retired axial bounds.");
            var all = (System.Collections.Generic.IList<CollisionSweepFrameViewModel>)typeof(CollisionViewModel).GetProperty("AllFrames").GetValue(model);
            TestAssert.Equal(2, all.Count);
            model.SelectedFrame = all[0];
            TestAssert.Equal("A", model.SelectedBeamId, "Selecting a different beam's table row must switch the active envelope.");
            TestAssert.True(ReferenceEquals(all, model.AllFrames), "Selection must not reset the table's collection view.");
        }

        private static CollisionViewModel Create()
        {
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            var model = new CollisionViewModel(snapshot);
            model.SetScene(SyntheticCollisionFactory.Create(snapshot.ActivePlanKey), null);
            return model;
        }
        private static object Property(object value, string name)
        {
            var property = value.GetType().GetProperty(name);
            TestAssert.NotNull(property, "Missing collision property " + name);
            return property.GetValue(value);
        }
        private static MethodInfo Method(object value, string name)
        {
            var method = value.GetType().GetMethod(name);
            TestAssert.NotNull(method, "Missing collision action " + name);
            return method;
        }
        private static ICommand Command(object value, string name) { return (ICommand)Property(value, name); }
        private static string FindRoot()
        {
            for (var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory); dir != null; dir = dir.Parent)
                if (File.Exists(Path.Combine(dir.FullName, "ClearPlan.sln"))) return dir.FullName;
            throw new InvalidOperationException("Repository root unavailable.");
        }
    }
}
