using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Threading;
using ClearPlan.Core.Collision;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Presentation.Views
{
    public partial class CollisionView : UserControl
    {
        private CollisionViewModel model;
        private bool pending;
        private readonly DispatcherTimer playback = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
        private CollisionScene cachedScene;
        private readonly Dictionary<CollisionSurface, MeshGeometry3D> surfaceMeshes = new Dictionary<CollisionSurface, MeshGeometry3D>();
        private readonly Dictionary<CollisionPose, List<MeshGeometry3D>> envelopeMeshes = new Dictionary<CollisionPose, List<MeshGeometry3D>>();
        private List<Point3D> framingPoints;
        private string framingBeam;
        private bool framingBody, framingTable, framingEnvelope, framingRoom;
        private bool framingPlayback, renderedPlayback;
        private CollisionPose framingPose;
        private bool captureRequested;
        private bool reportOrientationLayout;
        private byte[] orientationPng;
        private ReviewBeamAnalysis insetBeam;
        private ReviewControlPointSample insetCp;
        private CollisionScene insetScene;
        private ApertureGeometry insetAperture;
        private Point? orbitStart;
        private void OrbitStart(object sender, MouseButtonEventArgs e)
        {
            if (model == null || !model.HasScene) return;
            model.StopPlayback(); orbitStart = e.GetPosition(IllustrationRoot); IllustrationRoot.CaptureMouse(); e.Handled = true;
        }
        private void OrbitStop(object sender, MouseButtonEventArgs e) { orbitStart = null; IllustrationRoot.ReleaseMouseCapture(); }
        private void OrbitMove(object sender, MouseEventArgs e)
        {
            if (!orbitStart.HasValue || model == null || e.LeftButton != MouseButtonState.Pressed) return;
            var current = e.GetPosition(IllustrationRoot); var delta = current - orbitStart.Value; orbitStart = current;
            model.Yaw = Math.IEEERemainder(model.Yaw + delta.X * 0.35, 360);
            model.Elevation = Math.Max(-75, Math.Min(75, model.Elevation + delta.Y * 0.3));
        }
        private void ZoomScene(object sender, MouseWheelEventArgs e)
        { if (model == null || !model.HasScene) return; model.Zoom = Math.Max(0.6, Math.Min(2, model.Zoom * (e.Delta > 0 ? 1.12 : 1/1.12))); e.Handled = true; }
        public static void CaptureReportSweep(CollisionViewModel source)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (System.Threading.Thread.CurrentThread.GetApartmentState() != System.Threading.ApartmentState.STA)
                throw new InvalidOperationException("WPF collision capture requires the owner STA.");
            source.StopPlayback(); source.StoreReportSweep(null);
            if (source.Sweep == null || source.Sweep.Frames.Count == 0) { source.StoreIllustration(null); return; }
            if (!source.IncludeInReport || !source.HasScene) return;
            string selectedBeam = source.SelectedBeamId; int selectedPosition = source.PosePosition;
            bool ghosts = source.GhostsVisible;
            var reports = new List<CollisionReportBeam>();
            var view = new CollisionView { DataContext = source, captureRequested = true };
            try
            {
                view.Measure(new Size(1340, 800)); view.Arrange(new Rect(0, 0, 1340, 800)); view.UpdateLayout();
                source.GhostsVisible = true;
                foreach (string beam in source.BeamIds)
                {
                    source.SelectedBeamId = beam;
                    var minimum = source.BeamFrames.Where(f => f.MinimumDisplayedDistance.HasValue).OrderBy(f => f.MinimumDisplayedDistance.Value).FirstOrDefault();
                    source.SelectedFrame = minimum ?? source.BeamFrames.FirstOrDefault();
                    view.DrawScene();
                    var rows = source.BeamFrames.Select(f => new CollisionReportRow {
                        GantryDegrees = f.Frame.Pose.GantryDegrees, Interpolated = f.Frame.Interpolated,
                        CouchDegrees = f.Frame.Pose.PatientSupportAngleDegrees, MinimumDistanceMm = f.MinimumDisplayedDistance,
                        CapturedControlPointIndex = f.Frame.CapturedControlPointIndex,
                        Status = f.StatusCode, BodyDistance = f.BodyDistanceText, TableDistance = f.TableDistanceText,
                        BodyStatus = f.BodyStatusCode, TableStatus = f.TableStatusCode, Reason = f.Reason }).ToList();
                    string status = CollisionSweepFrameViewModel.Aggregate(rows.Select(r => r.Status));
                    reports.Add(new CollisionReportBeam { BeamId = beam, Status = status, RadialOnly = source.IsRadialReview, BodyOnly = !source.HasSeparateSupport, OverviewPng = source.IllustrationPng,
                        OrientationPng = minimum == null ? null : view.orientationPng,
                        Summary = source.IsRadialReview ? (source.HasSeparateSupport ? "Open bore: radial gap of captured body/support surfaces; no axial end wall."
                            : "Open bore: radial gap of captured body only; no axial end wall. No separate table model; table clearance not assessed.")
                            : source.HasSeparateSupport ? "Captured body and support surfaces; nominal source model."
                            : "Captured body only; no separate table model. Table clearance is not assessed.", Rows = rows });
                }
            }
            finally
            {
                source.SelectedBeamId = selectedBeam; source.PosePosition = selectedPosition; source.GhostsVisible = ghosts;
                view.DataContext = null;
            }
            CaptureReportIllustration(source); // retain the selected-image compatibility field after restoring GUI state
            source.StoreReportSweep(reports);
        }
        /// <summary>Render the same detached scene as the GUI, without showing a window or accessing ESAPI.</summary>
        public static void CaptureReportIllustration(CollisionViewModel source)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (System.Threading.Thread.CurrentThread.GetApartmentState() != System.Threading.ApartmentState.STA)
                throw new InvalidOperationException("WPF collision capture requires the owner STA.");
            source.StoreIllustration(null);
            if (!source.IncludeInReport || !source.HasScene || source.IsLoading) return;
            var view = new CollisionView { DataContext = source, captureRequested = true };
            try
            {
                // Explicit layout replaces Loaded/SizeChanged, which a headless export never raises.
                view.Measure(new Size(1340, 800));
                view.Arrange(new Rect(0, 0, 1340, 800));
                view.UpdateLayout();
                view.DrawScene();
            }
            finally { view.DataContext = null; }
        }
        public CollisionView()
        {
            InitializeComponent();
            playback.Tick += (s, e) => { if (model == null || !IsVisible) { playback.Stop(); if (model != null) model.StopPlayback(); } else model.AdvancePlayback(); };
            DataContextChanged += ContextChanged;
            Loaded += (s, e) => { Connect(); if (model != null && model.WorstPoseCommand.CanExecute(null)) model.WorstPoseCommand.Execute(null); ScheduleRender(); };
            Unloaded += (s, e) => { playback.Stop(); if (model != null) { model.StopPlayback(); model.SceneChanged -= Changed; model.PropertyChanged -= ModelPropertyChanged; } model = null; };
            IsVisibleChanged += (s, e) => { if (!IsVisible) { playback.Stop(); if (model != null) model.StopPlayback(); } };
            SizeChanged += (s, e) => ScheduleRender();
            Loaded += (s, e) => ClearPlan.Core.Localization.ReviewLanguage.Changed += LanguageChanged;
            Unloaded += (s, e) => ClearPlan.Core.Localization.ReviewLanguage.Changed -= LanguageChanged;
        }
        private void LanguageChanged(object sender, EventArgs args) { ScheduleRender(); }
        private void ContextChanged(object sender, DependencyPropertyChangedEventArgs args) { Connect(); ScheduleRender(); }
        private void Connect()
        {
            playback.Stop();
            if (model != null) { model.StopPlayback(); model.SceneChanged -= Changed; model.PropertyChanged -= ModelPropertyChanged; }
            model = DataContext as CollisionViewModel;
            if (model != null) { model.SceneChanged += Changed; model.PropertyChanged += ModelPropertyChanged; }
        }
        private void ModelPropertyChanged(object sender, PropertyChangedEventArgs args)
        {
            if (model != null && renderedPlayback != model.IsPlaying) { renderedPlayback=model.IsPlaying; ScheduleRender(); }
            if (model != null && model.IsPlaying && IsLoaded && IsVisible) playback.Start();
            else playback.Stop();
        }
        private void Changed(object sender, EventArgs args) { ScheduleRender(); }
        private void ScheduleRender()
        {
            if (pending) return;
            pending = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                pending = false;
                if (!IsLoaded) return;
                try { DrawScene(); }
                catch (Exception)
                {
                    SceneViewport.Children.Clear();
                    ClearRoomReference();
                    if (model != null) model.SetFailure("3D-Geometrie konnte nicht sicher dargestellt werden. Daten prüfen und erneut laden.");
                }
            }), DispatcherPriority.Render);
        }
        private void DrawMlcInset()
        {
            var beam = model == null ? null : model.MlcBeam;
            var cp = model == null ? null : model.MlcControlPoint;
            var aperture = cp == null ? null : cp.Aperture;
            var scene = model == null ? null : model.Scene;
            if (ReferenceEquals(insetScene,scene) && ReferenceEquals(insetAperture,aperture) &&
                ReferenceEquals(insetBeam,beam) && ReferenceEquals(insetCp,cp) && MlcInset.Source != null) return;
            insetScene = scene; insetAperture = aperture;
            insetBeam = beam; insetCp = cp; MlcInset.Source = null;
            var bytes = ClearPlan.Rendering.BeamEyeViewRenderer.RenderApertureInset(beam,cp);
            using(var stream = new MemoryStream(bytes,false))
            {
                var image = new BitmapImage(); image.BeginInit(); image.CacheOption = BitmapCacheOption.OnLoad;
                image.StreamSource = stream; image.EndInit(); image.Freeze(); MlcInset.Source = image;
            }
        }
        private void DrawScene()
        {
            SceneViewport.Children.Clear();
            ClearRoomReference();
            DrawMlcInset();
            // Live bindings can advance before the queued 3D draw. Keep every
            // label inside the scene on the pose rendered in this same pass.
            ScenePoseHeading.Text = model == null ? "" : ClearPlan.Core.Localization.ReviewLanguage.Display(model.PoseText);
            SceneSummary.Text = model == null ? "" : ClearPlan.Core.Localization.ReviewLanguage.Display(model.IllustrationSummary);
            SceneSummary.Foreground = model == null ? Brushes.White : (Brush)new BrushConverter().ConvertFromString(model.IllustrationSummaryColor);
            SceneActiveStatus.Text = model == null || model.SelectedFrame == null ? "—" : ClearPlan.Core.Localization.ReviewLanguage.Display(model.SelectedFrame.StatusText);
            SceneActiveStatus.Foreground = model == null || model.SelectedFrame == null ? Brushes.White : (Brush)new BrushConverter().ConvertFromString(model.SelectedFrame.StatusColor);
            SceneSampling.Text = model == null ? "" : ClearPlan.Core.Localization.ReviewLanguage.Display(model.SamplingText);
            if (model == null || !model.HasScene) return;
            if (!ReferenceEquals(cachedScene, model.Scene)) { cachedScene = model.Scene; surfaceMeshes.Clear(); envelopeMeshes.Clear(); framingPoints = null; }
            var pose = model.SelectedPose;
            // Conflicting ring/bore identity must never acquire a C-arm schematic.
            bool boreReference = (model.HasSourceModel && model.Scene.SourceModel.Kind == "HalcyonBoreSource") ||
                (model.Scene.Profile != null && model.Scene.Profile.Kind == "RingBore");
            bool cArm = model.RoomView && pose != null && !boreReference &&
                ((model.HasSourceModel && model.Scene.SourceModel.Kind == "TrueBeamHeadSource") ||
                (model.UsesProfileGeometry && model.Scene.Profile.Kind == "CArmSphere"));
            RoomReferenceRoot.Visibility = model.RoomView && pose != null ? Visibility.Visible : Visibility.Collapsed;
            RoomReferenceUnavailable.Visibility = cArm ? Visibility.Collapsed : Visibility.Visible;
            ConfigureOrientationLayout();
            SceneViewport.Margin = IsoOverlay.Margin = new Thickness(0,50,0,42);
            IllustrationRoot.UpdateLayout();
            if (!captureRequested)
            {
                // Wrapped labels remain outside the patient viewport even in
                // the compact layout. The report retains its fixed capture fit.
                SceneViewport.Margin = IsoOverlay.Margin = new Thickness(0, Math.Max(50,SceneHeadingPanel.ActualHeight+24),
                    0,Math.Max(42,SceneLegend.ActualHeight+20));
                IllustrationRoot.UpdateLayout();
            }
            if (cArm)
            {
                RoomReferenceViewport.Children.Add(new ModelVisual3D { Content = CreateRoomReference(pose) });
                var eye = new Vector3D(1.1,-0.75,-1.4); eye.Normalize();
                double roomAspect = RoomReferenceViewport.ActualWidth / Math.Max(1, RoomReferenceViewport.ActualHeight);
                RoomReferenceViewport.Camera = new OrthographicCamera(new Point3D(0,140,-40)+eye*5000, -eye, new Vector3D(0,-1,0), Math.Max(2600, 2450*roomAspect));
            }
            if (model.RoomView && pose != null)
            {
                GantryCompass.Source = CreateAngleCompass(pose.GantryDegrees, false);
                RoomCompass.Source = CreateAngleCompass(pose.PatientSupportAngleDegrees, true);
            }
            var origin = pose == null ? new Point3D() : Point(pose.Isocenter);
            var group = new Model3DGroup();
            group.Children.Add(new AmbientLight(Color.FromRgb(140, 140, 140)));
            group.Children.Add(new DirectionalLight(Colors.White, new Vector3D(-1, 1, -2)));
            double radius = 350;
            double lowZ = origin.Z, highZ = origin.Z;
            foreach (var surface in model.Scene.Surfaces)
            {
                if (surface.Mesh == null || surface.Mesh.Vertices == null) continue;
                foreach (var vertex in surface.Mesh.Vertices)
                {
                    radius = Math.Max(radius, (Point(vertex) - origin).Length);
                    lowZ = Math.Min(lowZ, vertex.Z); highZ = Math.Max(highZ, vertex.Z);
                }
                if ((surface.Role == "External" && !model.BodyVisible) || (surface.Role == "Support" && !model.TableVisible)) continue;
                MeshGeometry3D mesh;
                if (!surfaceMeshes.TryGetValue(surface, out mesh))
                {
                    mesh = new MeshGeometry3D { Positions = new Point3DCollection(surface.Mesh.Vertices.Select(Point)),
                        TriangleIndices = new Int32Collection(surface.Mesh.TriangleIndices) };
                    mesh.Freeze(); surfaceMeshes.Add(surface, mesh);
                }
                Add(group, mesh, surface.Role == "External" ? Color.FromRgb(115, 198, 190) : Color.FromRgb(157, 177, 199));
            }
            bool ring = (model.DeviceEnvelopeDrawable && model.Scene.Profile.Kind == "RingBore") ||
                (model.SourceEnvelopeDrawable && model.Scene.SourceModel.Kind == "HalcyonBoreSource");
            // A ring housing is stationary: do not stack coincident transparent cylinders
            // or suggest that the bore rotates with the gantry. The angle rows retain all states.
            if (model.GhostsVisible && !ring)
                DrawTrajectory(group, origin);
            if (model.HullsVisible && !ring)
            {
                var distinct = new List<CollisionPose>();
                foreach (var frame in model.BeamFrames)
                {
                    var ghost = frame.Frame.Pose;
                    if (ReferenceEquals(frame, model.SelectedFrame) || distinct.Any(previous => SameEnvelopePose(previous, ghost)) || SameEnvelopePose(ghost, pose)) continue;
                    distinct.Add(ghost);
                    radius = Math.Max(radius, DrawEnvelope(group, ghost, frame.StatusCode, false, lowZ, highZ) + (Point(ghost.Isocenter) - origin).Length);
                }
            }
            if (model.SelectedFrame != null)
                radius = Math.Max(radius, DrawEnvelope(group, pose, model.SelectedFrame.StatusCode, true, lowZ, highZ));
            Add(group, Box(origin, 70, 3, 3), Colors.White);
            Add(group, Box(origin, 3, 70, 3), Colors.White);
            Add(group, Box(origin, 3, 3, 70), Colors.White);
            var display = model.RoomView ? RoomDisplayTransform(pose) : Transform3D.Identity;
            group.Transform = display;
            SceneViewport.Children.Add(new ModelVisual3D { Content = group });
            double yaw = model.Yaw * Math.PI / 180, tilt = model.Elevation * Math.PI / 180;
            var outward = new Vector3D(Math.Sin(yaw)*Math.Cos(tilt), -Math.Sin(tilt), Math.Cos(yaw)*Math.Cos(tilt));
            var right = Vector3D.CrossProduct(-outward, new Vector3D(0,-1,0)); right.Normalize();
            var up = Vector3D.CrossProduct(right, -outward); up.Normalize();
            var framingOrigin = model.BeamFrames.Count == 0 ? origin : Point(model.BeamFrames[0].Frame.Pose.Isocenter);
            double minX=0, maxX=0, minY=0, maxY=0;
            var fitPoints = GetFramingPoints(lowZ, highZ);
            foreach (var point in fitPoints)
            {
                    var relative = point - framingOrigin;
                    // Camera focus only: distant support extensions still render
                    // and all captured surfaces remain in the geometric calculation.
                    if (relative.Length > 1000) continue;
                    double x=Vector3D.DotProduct(relative,right), y=Vector3D.DotProduct(relative,up);
                    minX=Math.Min(minX,x); maxX=Math.Max(maxX,x); minY=Math.Min(minY,y); maxY=Math.Max(maxY,y);
            }
            double aspect = Math.Max(0.1, SceneViewport.ActualWidth / Math.Max(1, SceneViewport.ActualHeight));
            double centerX=(minX+maxX)/2, centerY=(minY+maxY)/2;
            double span = Math.Max(300, Math.Max(maxX-minX,(maxY-minY)*aspect))*1.06/model.Zoom;
            var target = framingOrigin + right*centerX + up*centerY;
            double framingRadius = Math.Max(350, fitPoints.Select(p=>(p-framingOrigin).Length).DefaultIfEmpty(350).Max());
            double distance = framingRadius*3+2000;
            SceneViewport.Camera = new OrthographicCamera(target + outward*distance, -outward, up, span)
                { NearPlaneDistance=1, FarPlaneDistance=distance+framingRadius*5 };
            var isoOffset = origin-framingOrigin;
            IsoMarker.RenderTransform = new TranslateTransform((Vector3D.DotProduct(isoOffset,right)-centerX)/span*SceneViewport.ActualWidth,
                (centerY-Vector3D.DotProduct(isoOffset,up))/(span/aspect)*SceneViewport.ActualHeight);
            // Playback changes geometry/material state only, not PNG compression on each tick.
            if (!model.IncludeInReport || !captureRequested || model.IsPlaying) return;
            IllustrationRoot.UpdateLayout();
            // Report-only views retain the horizontal strip and full-width
            // main viewport; the GUI's right rail never changes this capture fit.
            orientationPng = RoomReferenceRoot.Visibility == Visibility.Visible ? CapturePng(RoomReferenceRoot) : null;
            var visibility = RoomReferenceRoot.Visibility;
            try
            {
                RoomReferenceRoot.Visibility = Visibility.Hidden;
                IllustrationRoot.UpdateLayout();
                model.StoreIllustration(CapturePng(IllustrationRoot));
            }
            finally { RoomReferenceRoot.Visibility = visibility; }
        }
        private void ConfigureOrientationLayout()
        {
            RoomReferenceColumn.Width = new GridLength(!captureRequested && RoomReferenceRoot.Visibility == Visibility.Visible
                ? Math.Max(124,Math.Min(156,IllustrationRoot.ActualWidth*0.22)) : 0);
            if (!captureRequested || reportOrientationLayout) return;
            reportOrientationLayout = true;
            Grid.SetColumn(RoomReferenceRoot,0); Grid.SetColumnSpan(RoomReferenceRoot,2);
            RoomReferenceRoot.Width=380; RoomReferenceRoot.Height=132;
            RoomReferenceRoot.HorizontalAlignment=HorizontalAlignment.Center; RoomReferenceRoot.VerticalAlignment=VerticalAlignment.Bottom;
            RoomReferenceRoot.Margin=new Thickness(0,0,0,45);
            Grid.SetColumnSpan(SceneHeadingPanel,2); Grid.SetColumnSpan(SceneLegend,2);
            ScenePoseHeading.TextWrapping=TextWrapping.NoWrap;
            ((TextBlock)SceneHeadingPanel.Children[1]).TextWrapping=TextWrapping.NoWrap;
            OrientationLayout.RowDefinitions.Clear();
            OrientationLayout.RowDefinitions.Add(new RowDefinition {Height=GridLength.Auto});
            OrientationLayout.RowDefinitions.Add(new RowDefinition {Height=new GridLength(1,GridUnitType.Star)});
            OrientationLayout.RowDefinitions.Add(new RowDefinition {Height=GridLength.Auto});
            OrientationLayout.ColumnDefinitions.Clear();
            foreach(double width in new[]{158d,104d,104d}) OrientationLayout.ColumnDefinitions.Add(new ColumnDefinition {Width=new GridLength(width)});
            Grid.SetColumnSpan(OrientationTitle,3); Grid.SetColumnSpan(OrientationFooter,3);
            Grid.SetRow(GantryCompass,1); Grid.SetColumn(GantryCompass,1);
            Grid.SetRow(RoomCompass,1); Grid.SetColumn(RoomCompass,2);
            Grid.SetRow(OrientationFooter,2);
            OrientationTitle.Text=ClearPlan.Core.Localization.ReviewLanguage.Text("Raumorientierung · Schema");
            OrientationFooter.Text=ClearPlan.Core.Localization.ReviewLanguage.Text("Nominal · 0° gestrichelt · kein Kollisionsmodell");
        }
        private static byte[] CapturePng(FrameworkElement element)
        {
            int width = (int)Math.Ceiling(element.ActualWidth), height = (int)Math.Ceiling(element.ActualHeight);
            if (width < 2 || height < 2) return null;
            double scale = Math.Min(2, Math.Sqrt(4000000.0 / ((double)width * height)));
            var bitmap = new RenderTargetBitmap((int)Math.Ceiling(width * scale), (int)Math.Ceiling(height * scale), 96 * scale, 96 * scale, PixelFormats.Pbgra32);
            // Render in the element's local coordinates, excluding its parent offset.
            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
                dc.DrawRectangle(new VisualBrush(element) { Stretch = Stretch.Fill }, null, new Rect(0,0,width,height));
            bitmap.Render(visual);
            var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap));
            using (var stream = new MemoryStream()) { encoder.Save(stream); return stream.ToArray(); }
        }
        private void ClearRoomReference()
        {
            RoomReferenceViewport.Children.Clear(); RoomCompass.Source=null; GantryCompass.Source=null; orientationPng=null;
            RoomReferenceRoot.Visibility=Visibility.Collapsed; RoomReferenceColumn.Width=new GridLength(0);
        }
        private static bool SameEnvelopePose(CollisionPose a, CollisionPose b)
        {
            return a != null && b != null && Math.Abs(Math.IEEERemainder(a.GantryDegrees - b.GantryDegrees, 360)) < 1e-6 &&
                (Point(a.Isocenter) - Point(b.Isocenter)).Length < 1e-6 &&
                (Point(a.Source) - Point(b.Source)).Length < 1e-6;
        }
        private IList<Point3D> GetFramingPoints(double lowZ, double highZ)
        {
            // Manual review/report: fit the selected pose, not empty space for
            // every other gantry angle. Playback: one stable field-wide camera.
            // These points are display-only and never enter distance checks.
            if (framingPoints != null && framingBeam == model.SelectedBeamId && framingBody == model.BodyVisible &&
                framingTable == model.TableVisible && framingEnvelope == model.EnvelopeVisible && framingRoom == model.RoomView &&
                framingPlayback == model.IsPlaying && (model.IsPlaying || ReferenceEquals(framingPose,model.SelectedPose)))
                return framingPoints;
            framingPoints = new List<Point3D>(); framingBeam = model.SelectedBeamId;
            framingBody = model.BodyVisible; framingTable = model.TableVisible;
            framingEnvelope = model.EnvelopeVisible; framingRoom = model.RoomView;
            framingPlayback=model.IsPlaying; framingPose=model.SelectedPose;
            var bodyTransforms = new HashSet<string>();
            if (model.BeamFrames.Count == 0)
                foreach (var surface in surfaceMeshes)
                {
                    if ((surface.Key.Role == "External" && !model.BodyVisible) || (surface.Key.Role == "Support" && !model.TableVisible)) continue;
                    framingPoints.AddRange(surface.Value.Positions);
                }
            var fittedFrames = model.IsPlaying ? (IEnumerable<CollisionSweepFrameViewModel>)model.BeamFrames : model.SelectedFrame == null ? new CollisionSweepFrameViewModel[0] : new[] { model.SelectedFrame };
            foreach (var frame in fittedFrames)
            {
                var pose = frame.Frame.Pose;
                var transform = model.RoomView ? RoomDisplayTransform(pose) : Transform3D.Identity;
                string transformKey = transform.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (bodyTransforms.Add(transformKey))
                    foreach (var surface in surfaceMeshes)
                    {
                        if ((surface.Key.Role == "External" && !model.BodyVisible) || (surface.Key.Role == "Support" && !model.TableVisible)) continue;
                        framingPoints.AddRange(surface.Value.Positions.Select(p=>transform.Transform(p)));
                    }
                var envelope = new Model3DGroup();
                DrawEnvelope(envelope, pose, frame.StatusCode, true, lowZ, highZ);
                foreach (var geometry in envelope.Children.OfType<GeometryModel3D>())
                    framingPoints.AddRange(((MeshGeometry3D)geometry.Geometry).Positions.Select(p=>transform.Transform(p)));
                framingPoints.Add(Point(pose.Isocenter));
            }
            return framingPoints;
        }
        private static Transform3D RoomDisplayTransform(CollisionPose pose)
        {
            if (pose == null) return Transform3D.Identity;
            // Display-only rigid frame change. Apply it ONCE to the body, support,
            // native source and model together; native distance inputs are untouched.
            return new RotateTransform3D(new AxisAngleRotation3D(new Vector3D(0,1,0), -pose.PatientSupportAngleDegrees), Point(pose.Isocenter));
        }
        // Display-only room reference, deliberately separate from the patient close-up.
        // The fixed stand and axis never rotate with the couch. Schematic parts
        // are never added to CollisionScene.Surfaces or used for distance checks.
        private static Model3DGroup CreateRoomReference(CollisionPose pose)
        {
            var group = new Model3DGroup();
            group.Children.Add(new AmbientLight(Color.FromRgb(175,175,175)));
            group.Children.Add(new DirectionalLight(Colors.White, new Vector3D(-1,-1,-2)));
            Add(group, Box(new Point3D(0,20,1090), 700,1500,180), Color.FromRgb(96,112,129));
            Add(group, Box(new Point3D(0,810,970), 950,70,650), Color.FromRgb(80,95,112));
            Add(group, Segment(new Point3D(0,0,-200), new Point3D(0,0,920), 10), Color.FromRgb(186,204,216));
            var iso = Point(pose.Isocenter);
            var direction = RoomDisplayTransform(pose).Transform(Point(pose.Source)) - iso;
            double sad=direction.Length;
            if (sad > 1e-6) direction.Normalize();
            var head = new Point3D() + direction*sad;
            var rear = head + new Vector3D(0,0,930);
            var rotor = new Model3DGroup();
            for(int a=10;a<360;a+=20) { double t=a*Math.PI/180, end=(a+10)*Math.PI/180;
                Add(rotor,Segment(new Point3D(Math.Sin(t)*sad,-Math.Cos(t)*sad,0),new Point3D(Math.Sin(end)*sad,-Math.Cos(end)*sad,0),5),Color.FromArgb(180,125,184,222)); }
            Add(rotor, Bar(new Point3D(0,0,930), rear, 360, 150), Color.FromRgb(180,197,210));
            group.Children.Add(rotor);
            Add(group, Bar(rear,head,380,150), Color.FromRgb(169,188,202));
            Add(group, Bar(head-direction*160,head+direction*110,290,300), Color.FromRgb(218,230,238));
            Add(group, Segment(new Point3D(),head,3), Color.FromArgb(170,238,215,148));
            var patientAxis = new Model3DGroup();
            Add(patientAxis, Box(new Point3D(0,150,-340),380,65,1650), Color.FromRgb(66,150,147));
            Add(patientAxis, Box(new Point3D(0,480,-880),270,600,280), Color.FromRgb(66,150,147));
            var torso = Sphere(new Point3D(),1);
            for(int i=0;i<torso.Positions.Count;i++) { var p=torso.Positions[i]; torso.Positions[i]=new Point3D(p.X*130,p.Y*95+25,p.Z*330-110); }
            Add(patientAxis, torso, Color.FromRgb(130,215,200));
            Add(patientAxis, Sphere(new Point3D(0,18,320),90), Color.FromRgb(174,238,221));
            Add(patientAxis, Bar(new Point3D(-65,50,-420),new Point3D(-65,65,-1010),70,90), Color.FromRgb(130,215,200));
            Add(patientAxis, Bar(new Point3D(65,50,-420),new Point3D(65,65,-1010),70,90), Color.FromRgb(130,215,200));
            patientAxis.Transform = new RotateTransform3D(new AxisAngleRotation3D(new Vector3D(0,1,0),-pose.PatientSupportAngleDegrees));
            group.Children.Add(patientAxis);
            Add(group, Marker(new Point3D(),25), Colors.White);
            return group;
        }
        private static MeshGeometry3D Bar(Point3D a, Point3D b, double width, double height)
        {
            var axis=b-a; if(axis.Length<1e-6)return new MeshGeometry3D(); axis.Normalize();
            var u=Vector3D.CrossProduct(axis, Math.Abs(axis.Y)<0.9 ? new Vector3D(0,1,0) : new Vector3D(0,0,1)); u.Normalize();
            var v=Vector3D.CrossProduct(axis,u); var mesh=new MeshGeometry3D();
            foreach(var end in new[]{a,b}) foreach(var sign in new[]{new[]{-1,-1},new[]{1,-1},new[]{1,1},new[]{-1,1}})
                mesh.Positions.Add(end+u*(sign[0]*width/2)+v*(sign[1]*height/2));
            Quad(mesh,0,1,2,3); Quad(mesh,4,7,6,5);
            for(int i=0;i<4;i++) Quad(mesh,i,(i+1)%4,(i+1)%4+4,i+4);
            return mesh;
        }
        private static ImageSource CreateAngleCompass(double degrees, bool couch)
        {
            var drawing=new DrawingGroup();
            using(var dc=drawing.Open())
            {
                var center=new Point(52,53); var muted=new SolidColorBrush(Color.FromRgb(191,207,217));
                var accent=new SolidColorBrush(couch ? Color.FromRgb(130,215,200) : Color.FromRgb(218,230,238));
                dc.DrawRectangle(Brushes.Transparent,null,new Rect(0,0,104,98));
                bool finite=!double.IsNaN(degrees) && !double.IsInfinity(degrees);
                var title=new FormattedText((couch ? "Couch " : "Gantry ")+(finite ? degrees.ToString("0.#",System.Globalization.CultureInfo.InvariantCulture)+"°" : "—"),
                    System.Globalization.CultureInfo.InvariantCulture,FlowDirection.LeftToRight,new Typeface("Segoe UI"),11,Brushes.White,1);
                dc.DrawText(title,new Point((104-title.Width)/2,0));
                dc.DrawEllipse(null,new Pen(muted,1),center,28,28);
                dc.DrawLine(new Pen(muted,1){DashStyle=DashStyles.Dash},new Point(52,20),new Point(52,85));
                dc.DrawLine(new Pen(muted,1),new Point(21,53),new Point(25,53));
                dc.DrawLine(new Pen(muted,1),new Point(79,53),new Point(83,53));
                var zero=new FormattedText("0°",System.Globalization.CultureInfo.InvariantCulture,FlowDirection.LeftToRight,new Typeface("Segoe UI"),9,muted,1);
                dc.DrawText(zero,new Point(84,16));
                if (finite)
                {
                    dc.PushTransform(new RotateTransform(couch ? -degrees : degrees,52,53));
                    if (couch)
                    {
                        dc.DrawRoundedRectangle(accent,null,new Rect(46,34,12,44),2,2);
                        dc.DrawEllipse(Brushes.White,null,new Point(52,37),3,3);
                    }
                    else
                    {
                        dc.DrawLine(new Pen(accent,1.5),center,new Point(52,25));
                        dc.DrawRectangle(accent,null,new Rect(45,22,14,10));
                    }
                    dc.Pop();
                }
                dc.DrawEllipse(Brushes.White,null,center,2.3,2.3);
                var caption=new FormattedText(ClearPlan.Core.Localization.ReviewLanguage.Text(couch ? "Draufsicht" : "Frontansicht"),System.Globalization.CultureInfo.InvariantCulture,
                    FlowDirection.LeftToRight,new Typeface("Segoe UI"),10,muted,1);
                dc.DrawText(caption,new Point((104-caption.Width)/2,85));
            }
            drawing.Freeze(); return new DrawingImage(drawing);
        }
        private void DrawTrajectory(Model3DGroup group, Point3D origin)
        {
            Point3D? previous = null;
            foreach (var frame in model.BeamFrames)
            {
                var direction=Point(frame.Frame.Pose.Source)-Point(frame.Frame.Pose.Isocenter);
                if (direction.Length < 1e-6) continue;
                direction.Normalize();
                double depth;
                if (model.SourceEnvelopeDrawable && model.Scene.SourceModel.Kind == "TrueBeamHeadSource")
                    depth = model.Scene.SourceModel.HeadObstacles.Min(o=>o.FrontFaceFromIsoMm);
                else if (model.DeviceEnvelopeDrawable && model.Scene.Profile.Kind == "CArmSphere")
                    depth = model.Scene.Profile.HeadCenterFromIsoMm-model.Scene.Profile.HeadRadiusMm;
                else continue; // no inferred device dimensions when no model is drawable
                var marker=Point(frame.Frame.Pose.Isocenter)+direction*depth;
                var color=CollisionSweepFrameViewModel.IsFullyClear(frame.StatusCode) ? Color.FromRgb(105,211,145)
                    : frame.StatusCode == "hit" || frame.StatusCode == "model-hit" ? Color.FromRgb(255,123,114) : Color.FromRgb(242,196,92);
                if (previous.HasValue) Add(group, Segment(previous.Value,marker,1.8), Color.FromArgb(150,180,195,205));
                Add(group, Marker(marker,ReferenceEquals(frame,model.SelectedFrame) ? 12 : 7),color);
                if (ReferenceEquals(frame,model.SelectedFrame)) Add(group,Segment(origin,marker,1.3),Color.FromArgb(180,190,216,225));
                previous=marker;
            }
        }
        private static MeshGeometry3D Marker(Point3D center, double size)
        {
            var mesh=new MeshGeometry3D();
            foreach(var axis in new[] {new Vector3D(size,0,0),new Vector3D(0,size,0),new Vector3D(0,0,size)})
            { mesh.Positions.Add(center+axis); mesh.Positions.Add(center-axis); }
            foreach(int x in new[]{0,1}) foreach(int y in new[]{2,3}) foreach(int z in new[]{4,5})
            {mesh.TriangleIndices.Add(x);mesh.TriangleIndices.Add(y);mesh.TriangleIndices.Add(z);}
            return mesh;
        }
        private static MeshGeometry3D Segment(Point3D a,Point3D b,double radius)
        {
            var mesh=new MeshGeometry3D(); var axis=b-a; if(axis.Length<1e-6)return mesh; axis.Normalize();
            var u=Vector3D.CrossProduct(axis,Math.Abs(axis.Z)<0.9?new Vector3D(0,0,1):new Vector3D(0,1,0)); u.Normalize();
            var v=Vector3D.CrossProduct(axis,u);
            foreach(var point in new[]{a,b}) for(int i=0;i<8;i++)
            {double angle=i*Math.PI/4;mesh.Positions.Add(point+(u*Math.Cos(angle)+v*Math.Sin(angle))*radius);}
            for(int i=0;i<8;i++)Quad(mesh,i,(i+1)%8,(i+1)%8+8,i+8);
            return mesh;
        }
        private double DrawEnvelope(Model3DGroup group, CollisionPose pose, string status, bool active, double lowZ, double highZ)
        {
            if (pose == null || (!model.DeviceEnvelopeDrawable && !model.SourceEnvelopeDrawable)) return 0;
            var origin = Point(pose.Isocenter);
            var profile = model.Scene.Profile;
            double radius = 0;
            if (model.DeviceEnvelopeDrawable)
                radius = profile.Kind == "CArmSphere" ? profile.HeadCenterFromIsoMm + profile.HeadRadiusMm :
                    Math.Sqrt(profile.BoreRadiusMm * profile.BoreRadiusMm + profile.BoreHalfLengthMm * profile.BoreHalfLengthMm);
            else if (model.Scene.SourceModel.Kind == "TrueBeamHeadSource") radius = model.Scene.SourceModel.ReachRadiusMm;
            else radius = Math.Sqrt(Math.Pow(Math.Max(100, (highZ - lowZ) / 2 + 50), 2) + Math.Pow(model.Scene.SourceModel.BoreLimitRadiusMm, 2)) + Math.Abs((lowZ + highZ) / 2 - origin.Z);
            List<MeshGeometry3D> meshes;
            if (!envelopeMeshes.TryGetValue(pose, out meshes))
            {
                meshes = new List<MeshGeometry3D>();
                var direction = Point(pose.Source) - origin;
                if (direction.Length < 1e-6) return radius;
                direction.Normalize();
                if (model.DeviceEnvelopeDrawable)
                {
                    if (profile.Kind == "CArmSphere") meshes.Add(Sphere(origin + direction * profile.HeadCenterFromIsoMm, profile.HeadRadiusMm));
                    else { meshes.Add(Bore(origin, profile.BoreRadiusMm, profile.BoreHalfLengthMm));
                        meshes.Add(BoreOutline(origin, profile.BoreRadiusMm, profile.BoreHalfLengthMm)); }
                }
                else
                {
                    var source = model.Scene.SourceModel;
                    if (source.Kind == "TrueBeamHeadSource")
                        foreach (var obstacle in source.HeadObstacles)
                        {
                            double? coveredFrom = source.HeadObstacles.Where(other => other.FrontFaceFromIsoMm > obstacle.FrontFaceFromIsoMm && other.RadiusMm >= obstacle.RadiusMm)
                                .Select(other => (double?)other.FrontFaceFromIsoMm).Min();
                            meshes.Add(SourceHead(origin, direction, obstacle.FrontFaceFromIsoMm, obstacle.RadiusMm, source.ReachRadiusMm, coveredFrom));
                            meshes.Add(SourceHeadOutline(origin, direction, obstacle.FrontFaceFromIsoMm, obstacle.RadiusMm, source.ReachRadiusMm, coveredFrom, false));
                            meshes.Add(SourceHeadOutline(origin, direction, obstacle.FrontFaceFromIsoMm, obstacle.RadiusMm, source.ReachRadiusMm, coveredFrom, true));
                        }
                    else
                    {
                        var center = new Point3D(origin.X, origin.Y, (lowZ + highZ) / 2);
                        double length = Math.Max(100, (highZ - lowZ) / 2 + 50);
                        meshes.Add(Bore(center, source.BoreLimitRadiusMm, length));
                        meshes.Add(BoreOutline(center, source.BoreLimitRadiusMm, length));
                    }
                }
                foreach (var mesh in meshes) mesh.Freeze();
                envelopeMeshes.Add(pose, meshes);
            }
            var color = CollisionSweepFrameViewModel.IsFullyClear(status) ? Color.FromRgb(70, 190, 115) : status == "hit" || status == "model-hit" ? Color.FromRgb(242, 95, 86) : Color.FromRgb(242, 196, 92);
            color.A = active ? (byte)255 : (byte)28;
            bool ring = model.DeviceEnvelopeDrawable ? profile.Kind == "RingBore" : model.Scene.SourceModel.Kind == "HalcyonBoreSource";
            bool sourceHead = !model.DeviceEnvelopeDrawable && model.Scene.SourceModel.Kind == "TrueBeamHeadSource";
            for (int i = 0; i < meshes.Count; i++)
            {
                var materialColor = color;
                if (sourceHead)
                {
                    // All other angles show only a thin front-face boundary. The
                    // active model has opaque contours and a quiet volume tint.
                    // No geometry or status in the distance calculation is changed.
                    if ((!active && i % 3 != 2) || (active && i % 3 == 2)) continue;
                    materialColor = active && i % 3 == 0 ? Color.FromRgb(126,145,159) : materialColor;
                    materialColor.A = active ? (i % 3 == 0 ? (byte)72 : (byte)255) : (byte)75;
                }
                // The fixed ring is a cage, not an opaque wall over the anatomy.
                // Its active edges remain fully opaque; the shell only indicates the no-fly boundary.
                if (ring && (i == 0 || i == 2)) materialColor.A = 22;
                Add(group, meshes[i], materialColor);
            }
            return radius;
        }
        private static Point3D Point(BeamPoint3D point) { return new Point3D(point.X, point.Y, point.Z); }
        private static void Add(Model3DGroup group, MeshGeometry3D mesh, Color color)
        {
            var material = new DiffuseMaterial(new SolidColorBrush(color));
            group.Children.Add(new GeometryModel3D(mesh, material) { BackMaterial = material });
        }
        private static MeshGeometry3D Sphere(Point3D center, double radius)
        {
            var mesh = new MeshGeometry3D(); const int slices = 48, rings = 24;
            for (int ring = 0; ring <= rings; ring++)
            for (int slice = 0; slice <= slices; slice++)
            {
                double p = Math.PI * ring / rings, t = 2 * Math.PI * slice / slices;
                mesh.Positions.Add(center + new Vector3D(Math.Sin(p) * Math.Cos(t), Math.Cos(p), Math.Sin(p) * Math.Sin(t)) * radius);
            }
            for (int ring = 0; ring < rings; ring++)
            for (int slice = 0; slice < slices; slice++) Quad(mesh, ring * (slices + 1) + slice, (ring + 1) * (slices + 1) + slice, (ring + 1) * (slices + 1) + slice + 1, ring * (slices + 1) + slice + 1);
            return mesh;
        }
        private static MeshGeometry3D SourceHead(Point3D iso, Vector3D axis, double face, double radius, double reach, double? coveredFrom)
        {
            // Semi-infinite source obstacle intersected with its stated spherical reach.
            // No finite rear-head depth is invented for this illustration.
            var mesh = new MeshGeometry3D(); const int slices = 64, rings = 12;
            double radial = Math.Min(radius, Math.Sqrt(Math.Max(0, reach * reach - face * face)));
            if (radial <= 0) return mesh;
            var u = Vector3D.CrossProduct(axis, Math.Abs(axis.Z) < 0.9 ? new Vector3D(0, 0, 1) : new Vector3D(0, 1, 0));
            u.Normalize(); var v = Vector3D.CrossProduct(axis, u);
            mesh.Positions.Add(iso + axis * face);
            for (int i = 0; i <= slices; i++)
            {
                double angle = i * 2 * Math.PI / slices;
                mesh.Positions.Add(iso + axis * face + (u * Math.Cos(angle) + v * Math.Sin(angle)) * radial);
            }
            for (int i = 1; i <= slices; i++) { mesh.TriangleIndices.Add(0); mesh.TriangleIndices.Add(i + 1); mesh.TriangleIndices.Add(i); }
            int cap = mesh.Positions.Count;
            if (coveredFrom.HasValue && coveredFrom.Value < Math.Sqrt(reach * reach - radial * radial))
            {
                for (int i = 0; i <= slices; i++)
                {
                    double angle = i * 2 * Math.PI / slices;
                    mesh.Positions.Add(iso + axis * coveredFrom.Value + (u * Math.Cos(angle) + v * Math.Sin(angle)) * radial);
                }
                for (int i = 0; i < slices; i++) Quad(mesh, i + 1, i + 2, cap + i + 1, cap + i);
                return mesh;
            }
            for (int ring = 0; ring <= rings; ring++) for (int i = 0; i <= slices; i++)
            {
                double r = radial * ring / rings, angle = i * 2 * Math.PI / slices;
                mesh.Positions.Add(iso + axis * Math.Sqrt(Math.Max(0, reach * reach - r * r)) +
                    (u * Math.Cos(angle) + v * Math.Sin(angle)) * r);
            }
            for (int ring = 0; ring < rings; ring++) for (int i = 0; i < slices; i++)
                Quad(mesh, cap + ring * (slices + 1) + i, cap + ring * (slices + 1) + i + 1,
                    cap + (ring + 1) * (slices + 1) + i + 1, cap + (ring + 1) * (slices + 1) + i);
            int last = cap + rings * (slices + 1);
            for (int i = 0; i < slices; i++) Quad(mesh, i + 1, i + 2, last + i + 1, last + i);
            return mesh;
        }
        private static MeshGeometry3D SourceHeadOutline(Point3D iso, Vector3D axis, double face, double radius, double reach, double? coveredFrom, bool frontOnly)
        {
            // Display contours of the same source obstacle, not a substitute
            // collision solid. No invented rear housing or patient geometry.
            var mesh = new MeshGeometry3D();
            double radial = Math.Min(radius, Math.Sqrt(Math.Max(0, reach * reach - face * face)));
            if (radial <= 0 || axis.Length < 1e-6) return mesh;
            axis.Normalize();
            var u = Vector3D.CrossProduct(axis, Math.Abs(axis.Z) < 0.9 ? new Vector3D(0,0,1) : new Vector3D(0,1,0));
            u.Normalize(); var v = Vector3D.CrossProduct(axis, u);
            double rear = Math.Sqrt(Math.Max(0, reach * reach - radial * radial));
            if (coveredFrom.HasValue) rear = Math.Min(rear, coveredFrom.Value);
            const int segments = 64; const double halfWidth = 1.5;
            foreach (double depth in frontOnly ? new[] { face } : new[] { face, rear })
            {
                int offset = mesh.Positions.Count;
                for (int i=0; i<=segments; i++)
                {
                    double angle = i * 2 * Math.PI / segments;
                    var direction = u * Math.Cos(angle) + v * Math.Sin(angle);
                    mesh.Positions.Add(iso + axis * depth + direction * Math.Max(0, radial-halfWidth));
                    mesh.Positions.Add(iso + axis * depth + direction * (radial+halfWidth));
                }
                for (int i=0; i<segments; i++) Quad(mesh, offset+2*i, offset+2*i+1, offset+2*i+3, offset+2*i+2);
            }
            if (!frontOnly) for (int i=0; i<4; i++)
            {
                double angle = i * Math.PI/2;
                var direction = u * Math.Cos(angle) + v * Math.Sin(angle);
                var tangent = (v * Math.Cos(angle) - u * Math.Sin(angle)) * halfWidth;
                int offset=mesh.Positions.Count;
                foreach(double depth in new[] { face, rear })
                {
                    mesh.Positions.Add(iso + axis * depth + direction * radial - tangent);
                    mesh.Positions.Add(iso + axis * depth + direction * radial + tangent);
                }
                Quad(mesh, offset, offset+1, offset+3, offset+2);
            }
            return mesh;
        }

        private static MeshGeometry3D Bore(Point3D center, double radius, double length)
        {
            var mesh = new MeshGeometry3D(); const int slices = 80;
            for (int i = 0; i <= slices; i++)
            {
                double angle = 2 * Math.PI * i / slices;
                mesh.Positions.Add(center + new Vector3D(radius * Math.Cos(angle), radius * Math.Sin(angle), -length));
                mesh.Positions.Add(center + new Vector3D(radius * Math.Cos(angle), radius * Math.Sin(angle), length));
            }
            for (int i = 0; i < slices; i++) Quad(mesh, 2 * i, 2 * i + 1, 2 * i + 3, 2 * i + 2);
            return mesh;
        }
        private static MeshGeometry3D BoreOutline(Point3D center, double radius, double length)
        {
            var mesh = new MeshGeometry3D(); const int slices = 80; const double width = 5;
            radius += 1; // lift opaque outlines off the translucent shell to avoid z-fighting
            foreach (double z in new[] { -length, 0.0, length })
            {
                int offset = mesh.Positions.Count;
                for (int i = 0; i <= slices; i++)
                {
                    double angle = 2 * Math.PI * i / slices;
                    foreach (double edge in new[] { -width / 2, width / 2 })
                        mesh.Positions.Add(center + new Vector3D(radius * Math.Cos(angle), radius * Math.Sin(angle), z + edge));
                }
                for (int i = 0; i < slices; i++) Quad(mesh, offset + 2*i, offset + 2*i+1, offset + 2*i+3, offset + 2*i+2);
            }
            for (int i = 0; i < 8; i++)
            {
                double angle = i * Math.PI / 4, delta = width / radius / 2;
                int offset = mesh.Positions.Count;
                foreach (double z in new[] { -length, length }) foreach (double edge in new[] { -delta, delta })
                    mesh.Positions.Add(center + new Vector3D(radius * Math.Cos(angle + edge), radius * Math.Sin(angle + edge), z));
                Quad(mesh, offset, offset + 1, offset + 3, offset + 2);
            }
            return mesh;
        }
        private static MeshGeometry3D LeadingBoundary(Point3D center, double radius, bool outline)
        {
            var mesh = new MeshGeometry3D(); const int slices = 80;
            if (!outline)
            {
                mesh.Positions.Add(center);
                for (int i = 0; i <= slices; i++)
                {
                    double angle = 2 * Math.PI * i / slices;
                    mesh.Positions.Add(center + new Vector3D(radius * Math.Cos(angle), radius * Math.Sin(angle), 0));
                }
                for (int i = 1; i <= slices; i++) { mesh.TriangleIndices.Add(0); mesh.TriangleIndices.Add(i); mesh.TriangleIndices.Add(i+1); }
            }
            else
            {
                center.Z += 1;
                for (int i = 0; i <= slices; i++) foreach (double r in new[] { radius - 2.5, radius + 2.5 })
                {
                    double angle = 2 * Math.PI * i / slices;
                    mesh.Positions.Add(center + new Vector3D(r * Math.Cos(angle), r * Math.Sin(angle), 0));
                }
                for (int i = 0; i < slices; i++) Quad(mesh, 2*i, 2*i+1, 2*i+3, 2*i+2);
            }
            return mesh;
        }
        private static MeshGeometry3D Box(Point3D center, double x, double y, double z)
        {
            var mesh = new MeshGeometry3D();
            foreach (int k in new[] { -1, 1 }) foreach (int j in new[] { -1, 1 }) foreach (int i in new[] { -1, 1 })
                mesh.Positions.Add(center + new Vector3D(i * x / 2, j * y / 2, k * z / 2));
            Quad(mesh, 0, 1, 3, 2); Quad(mesh, 4, 6, 7, 5); Quad(mesh, 0, 4, 5, 1);
            Quad(mesh, 2, 3, 7, 6); Quad(mesh, 0, 2, 6, 4); Quad(mesh, 1, 5, 7, 3); return mesh;
        }
        private static void Quad(MeshGeometry3D mesh, int a, int b, int c, int d)
        { foreach (int index in new[] { a, b, c, a, c, d }) mesh.TriangleIndices.Add(index); }
    }
}
