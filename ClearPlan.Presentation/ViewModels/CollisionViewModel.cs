using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows.Input;
using ClearPlan.Core.Collision;
using ClearPlan.Core.Review;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Presentation.ViewModels
{
    public sealed class CollisionViewModel : INotifyPropertyChanged
    {
        private readonly ReviewSnapshot snapshot;
        private string beamId;
        private string[] beamIds = new string[0];
        private int posePosition;
        private bool loading, bodyVisible = true, tableVisible = true, envelopeVisible = true;
        private bool zeroPitchRollConfirmed;
        private bool playing, advancingPlayback, ghostsVisible = true, hullsVisible;
        private bool allFieldsPlayback = true, repeatPlayback;
        private List<CollisionSweepFrameViewModel> frames = new List<CollisionSweepFrameViewModel>();
        private IList<CollisionSweepFrameViewModel> allFrames;
        private List<CollisionSweepFrameViewModel> beamFrames = new List<CollisionSweepFrameViewModel>();
        private double yaw = -135, elevation = 24, zoom = 1;
        private string message = "Geometrie noch nicht geladen. Geräteprofil unter Einstellungen konfigurieren und Geometrie laden.";
        public event PropertyChangedEventHandler PropertyChanged;
        public event EventHandler ReloadRequested;
        public event EventHandler SceneChanged;

        public CollisionViewModel(ReviewSnapshot source)
        {
            snapshot = source;
            allFrames = frames.AsReadOnly();
            ReloadCommand = new RelayCommand(p =>
            {
                if (snapshot.Synthetic)
                    SetScene(SyntheticCollisionFactory.Create(snapshot.ActivePlanKey, DemoKind), null);
                else { var handler = ReloadRequested; if (handler != null) handler(this, EventArgs.Empty); }
            }, p => !IsLoading);
            ResetCameraCommand = new RelayCommand(p => { yaw = -135; elevation = 24; zoom = 1; Notify(); ChangeScene(); });
            PlayPauseCommand = new RelayCommand(p =>
            {
                if (playing) StopPlayback();
                else
                {
                    if (PosePosition == MaximumPosePosition && (!AllFieldsPlayback || SelectedBeamId == BeamIds.LastOrDefault()))
                    { if (AllFieldsPlayback) SelectedBeamId = BeamIds.FirstOrDefault(); PosePosition = 0; }
                    playing = true; Notify();
                }
            }, p => playing || !loading && (AllFieldsPlayback ? frames.Count > 1 : BeamFrames.Count > 1));
            PreviousPoseCommand = new RelayCommand(p => { StopPlayback(); PosePosition--; }, p => !loading && PosePosition > 0);
            NextPoseCommand = new RelayCommand(p => { StopPlayback(); PosePosition++; }, p => !loading && PosePosition < MaximumPosePosition);
            WorstPoseCommand = new RelayCommand(p =>
            {
                StopPlayback();
                var worst = frames.Where(f => f.MinimumDisplayedDistance.HasValue)
                    .OrderBy(f => f.MinimumDisplayedDistance.Value).FirstOrDefault();
                if (worst == null) return;
                SelectedBeamId = worst.Frame.Pose.BeamId;
                SelectedFrame = worst;
            }, p => frames.Any(f => f.MinimumDisplayedDistance.HasValue));
            if (source.CollisionScene != null) SetScene(source.CollisionScene, null);
        }

        public ICommand ReloadCommand { get; private set; }
        public ICommand ResetCameraCommand { get; private set; }
        public ICommand WorstPoseCommand { get; private set; }
        public ICommand PlayPauseCommand { get; private set; }
        public ICommand PreviousPoseCommand { get; private set; }
        public ICommand NextPoseCommand { get; private set; }
        public bool IsPlaying { get { return playing; } }
        public bool AllFieldsPlayback { get { return allFieldsPlayback; } set { allFieldsPlayback = value; Notify(); CommandManager.InvalidateRequerySuggested(); } }
        public bool RepeatPlayback { get { return repeatPlayback; } set { repeatPlayback = value; Notify(); } }
        public string PlayText { get { return playing ? "Pause" : "Abspielen"; } }
        public CollisionSweepResult Sweep { get; private set; }
        public List<CollisionSweepFrameViewModel> BeamFrames { get { return beamFrames; } }
        public IList<CollisionSweepFrameViewModel> AllFrames { get { return allFrames; } }
        public CollisionSweepFrameViewModel SelectedFrame
        {
            get { return beamFrames.ElementAtOrDefault(posePosition); }
            set
            {
                if (value == null || !frames.Contains(value)) return;
                SelectedBeamId = value.Frame.Pose.BeamId;
                PosePosition = beamFrames.IndexOf(value);
            }
        }
        public bool GhostsVisible { get { return ghostsVisible; } set { if (ghostsVisible == value) return; ghostsVisible = value; Notify(); ChangeScene(); } }
        public bool HullsVisible { get { return hullsVisible; } set { if (hullsVisible == value) return; hullsVisible = value; Notify(); ChangeScene(); } }
        public bool RoomView { get { return Scene != null && Scene.PatientPosition == "HFS" &&
            (Scene.SupportedCoordinates || Scene.NominalCoordinatesSupported); } }
        public string ViewFrameText { get { return RoomView ? "Raumsicht · nominelle Tischrotation" : "Patientenkoordinaten"; } }
        public void StopPlayback() { if (!playing) return; playing = false; Notify(); }
        public void AdvancePlayback()
        {
            if (!playing) return;
            advancingPlayback = true;
            try
            {
                if (PosePosition < MaximumPosePosition) PosePosition++;
                else if (AllFieldsPlayback && Array.IndexOf(BeamIds, SelectedBeamId) < BeamIds.Length - 1)
                    SelectedBeamId = BeamIds[Array.IndexOf(BeamIds, SelectedBeamId) + 1];
                else if (RepeatPlayback)
                { if (AllFieldsPlayback) SelectedBeamId = BeamIds.FirstOrDefault(); PosePosition = 0; }
                else StopPlayback();
            }
            finally { advancingPlayback = false; }
            if (!RepeatPlayback && PosePosition >= MaximumPosePosition &&
                (!AllFieldsPlayback || SelectedBeamId == BeamIds.LastOrDefault())) StopPlayback();
        }
        public bool IsSynthetic { get { return snapshot.Synthetic; } }
        public bool CanConfirmCoordinates { get { return !snapshot.Synthetic && !loading; } }
        public bool ZeroPitchRollConfirmed
        {
            get { return zeroPitchRollConfirmed; }
            set
            {
                if (zeroPitchRollConfirmed == value) return;
                zeroPitchRollConfirmed = value;
                SetFailure("Lagerungsbestätigung geändert. Geometrie für diesen Plan erneut laden.");
            }
        }
        public string DemoKind { get; set; } = "CArmSphere";
        public string[] DemoKinds { get { return new[] { "CArmSphere", "RingBore" }; } }
        public CollisionScene Scene { get { return snapshot.CollisionScene; } }
        public CollisionResult Result { get; private set; }
        public bool HasScene { get { return Scene != null && Scene.Surfaces != null && Scene.Surfaces.Count > 0; } }
        public bool IsLoading { get { return loading; } }
        public string StatusText { get { return Result == null ? message :
            string.Format(CultureInfo.InvariantCulture, "{0} geprüfte 10°-Stichproben · {1} Oberflächen · {2}.",
                frames.Count, Scene.Surfaces.Count, StatusLabel); } }
        public string StatusDetail { get { return message; } }
        public string StatusLabel { get { return CollisionSweepFrameViewModel.Label(CollisionSweepFrameViewModel.Aggregate(frames.Select(f => f.StatusCode))); } }
        public bool IsRadialReview { get { return SelectedFrame != null && SelectedFrame.Frame.RadialReview != null; } }
        public string ShortLimitations { get { return IsRadialReview ? "Radialer Abstand der erfassten Geometrie · Offene Röhre · Außerhalb des CT nicht bewertet"
            : "Frei: Abstand der erfassten Flächen außerhalb des Warnabstands · 10°-Stichproben · Keine klinische Kollisionsfreigabe"; } }
        public string AssuranceText { get {
            if (Scene == null) return string.Empty;
            var gaps = new List<string>();
            if (!Scene.Synthetic && !UsesProfileGeometry) gaps.Add("Modell nicht kommissioniert");
            if (!Scene.SupportedCoordinates) gaps.Add("Lagerung nur nominell");
            if (Scene.ExternalTruncatedAtCtBoundary) gaps.Add("CT-Abdeckung unvollständig");
            else if (!Scene.Synthetic && !Scene.CtCoverageKnown) gaps.Add("CT-Abdeckung unbekannt");
            if (!Scene.Surfaces.Any(s => s.Role == "External" && s.Mesh != null && s.Mesh.Vertices != null && s.Mesh.Vertices.Count > 0)) gaps.Add("Körperoberfläche fehlt");
            if (HasSeparateSupport && !Scene.Surfaces.Any(s => s.Role == "Support" && s.Mesh != null && s.Mesh.Vertices != null && s.Mesh.Vertices.Count > 0)) gaps.Add("Tischoberfläche nicht auswertbar");
            return gaps.Count == 0 ? "Nur erfasste Geometrie; keine klinische Freigabe." : "Nachweise offen: " + string.Join(" · ", gaps) + ".";
        } }
        public string ProfileText
        {
            get
            {
                var profile = Scene == null ? null : Scene.Profile;
                if (HasSourceModel && !UsesProfileGeometry)
                    return "Quellmodell · " + Scene.SourceModel.ModelId + " · nicht kommissioniert";
                return profile == null ? "Kein passendes Geräteprofil · Abmessungen werden nicht geraten"
                    : profile.Kind + " · Revision " + profile.Revision + (snapshot.Synthetic ? " · ausschließlich Demo-Maße" : " · " + profile.MachineId + " · geprüft: " + profile.CommissionedBy);
            }
        }
        public string ProfileEvidence { get { return HasSourceModel && !UsesProfileGeometry
            ? Scene.SourceModel.Evidence + " · SHA-256 " + Scene.SourceModel.SourceSha256
            : Scene != null && Scene.Profile != null ? Scene.Profile.Evidence : "Kein Kommissionierungsnachweis geladen."; } }
        public bool HasSourceModel { get { return Scene != null && Scene.SourceModel != null; } }
        public string SourceScreeningLabel
        {
            get
            {
                var result = Scene == null ? null : Scene.SourceScreening;
                if (result == null || result.Status == "unavailable" || result.Status == "cancelled") return "Punkt-Screening nicht bewertbar";
                return result.Status == "screening-hit" ? "Treffer im Punkt-Screening" : "Kein Treffer an geprüften Punkten";
            }
        }
        public string SourceScreeningSummary { get { return Scene == null || Scene.SourceScreening == null ? "Kein Quellmodell zugeordnet." : Scene.SourceScreening.Summary; } }
        public bool HasSeparateSupport { get { return Scene != null && Scene.Surfaces.Any(s => s.Role == "Support"); } }
        public string SurfaceAvailabilityText { get {
            if (Scene == null) return string.Empty;
            var missing = new List<string>();
            if (!Scene.Surfaces.Any(s => s.Role == "External" && s.Mesh != null && s.Mesh.Vertices != null && s.Mesh.Vertices.Count > 0)) missing.Add("Körperoberfläche fehlt");
            if (!HasSeparateSupport) missing.Add("Körperbezogene Ansicht · kein separates Tischmodell");
            else if (!Scene.Surfaces.Any(s => s.Role == "Support" && s.Mesh != null && s.Mesh.Vertices != null && s.Mesh.Vertices.Count > 0)) missing.Add("Tischoberfläche nicht auswertbar");
            return string.Join(" · ", missing);
        } }
        public IEnumerable<SourceCollisionScreeningFinding> SelectedSourceFindings
        {
            get
            {
                var pose = SelectedPose;
                return Scene == null || Scene.SourceScreening == null ? new SourceCollisionScreeningFinding[0] :
                    Scene.SourceScreening.Findings.Where(f => pose != null && f.BeamId == pose.BeamId && f.ControlPointIndex == pose.ControlPointIndex);
            }
        }
        public string IllustrationStatus { get { return (IsRadialReview ? "Radial (erfasst) · " : HasSourceModel && !UsesProfileGeometry ? "Quellmodell · " : "Gesamtbefund: ") +
            CollisionSweepFrameViewModel.Label(CollisionSweepFrameViewModel.Aggregate(BeamFrames.Select(f => f.StatusCode))) +
            " · Aktiv: " + (SelectedFrame == null ? "Ungewiss" : SelectedFrame.StatusText) + " · " + BeamFrames.Count + " Positionen / 10°"; } }
        public string IllustrationSummary { get { return (IsRadialReview ? "Radial (erfasst) · " : HasSourceModel && !UsesProfileGeometry ? "Quellmodell · " : "Feld: ") + CollisionSweepFrameViewModel.Label(CollisionSweepFrameViewModel.Aggregate(BeamFrames.Select(f=>f.StatusCode))); } }
        public string IllustrationSummaryColor { get { var status=CollisionSweepFrameViewModel.Aggregate(BeamFrames.Select(f=>f.StatusCode)); return CollisionSweepFrameViewModel.IsFullyClear(status) ? "#69D391" : status=="hit"||status=="model-hit" ? "#FF7B72" : "#F2C45C"; } }
        public string SamplingText { get { return " · "+BeamFrames.Count+" Positionen / 10°"; } }
        public string RoomAngles { get { return SelectedPose==null ? "" : "Gantry "+SelectedPose.GantryDegrees.ToString("0.#")+"°\nTisch "+SelectedPose.PatientSupportAngleDegrees.ToString("0.#")+"°"; } }
        public const string Limitations = "Modell-Stichproben in 10°-Schritten, keine kontinuierliche Kollisionsfreigabe. Offener Körper / CT-Rand und fehlende Lagerungs- oder Modellnachweise bleiben ungewiss. Einfahrt, Zubehör und behandlungstägliche 6D-Korrekturen sind nicht freigegeben.";
        public string LimitationsText { get { return Limitations; } }
        public string[] BeamIds { get { return beamIds; } }
        public string SelectedBeamId
        {
            get { return beamId; }
            set { if (beamId == value) return; if (!advancingPlayback) StopPlayback(); beamId = value; posePosition = 0; SelectBeamFrames(); Notify(); ChangeScene(); CommandManager.InvalidateRequerySuggested(); }
        }
        private void SelectBeamFrames() { beamFrames = frames.Where(f => f.Frame.Pose.BeamId == beamId).ToList(); }
        public int MaximumPosePosition { get { return Math.Max(0, BeamFrames.Count - 1); } }
        public int PosePosition
        {
            get { return posePosition; }
            set { int next = Math.Max(0, Math.Min(MaximumPosePosition, value)); if (posePosition == next) return; if (!advancingPlayback) StopPlayback(); posePosition = next; Notify(); ChangeScene(); CommandManager.InvalidateRequerySuggested(); }
        }
        public CollisionPose SelectedPose { get { return SelectedFrame == null ? null : SelectedFrame.Frame.Pose; } }
        public bool HasSelectedPose { get { return SelectedPose != null; } }
        public ReviewBeamAnalysis MlcBeam { get {
            if (SelectedPose == null || snapshot.PlanAnalysis == null || snapshot.PlanAnalysis.Beams == null) return null;
            var matches = snapshot.PlanAnalysis.Beams.Where(b => b != null && b.BeamId == SelectedPose.BeamId).Take(2).ToArray();
            if (matches.Length != 1) return null;
            var beam = matches[0];
            if (!IsSynthetic)
            {
                string stamp;
                if (Scene == null || string.IsNullOrEmpty(Scene.NativePlanFingerprint) ||
                    Scene.NativePlanFingerprint != snapshot.PlanAnalysis.NativePlanFingerprint ||
                    Scene.NativeBeamFingerprints == null || !Scene.NativeBeamFingerprints.TryGetValue(beam.BeamId,out stamp) ||
                    string.IsNullOrEmpty(stamp) || stamp != beam.NativeGeometryFingerprint) return null;
            }
            return beam;
        } }
        public ReviewControlPointSample MlcControlPoint { get {
            var beam = MlcBeam; var frame = SelectedFrame == null ? null : SelectedFrame.Frame;
            if (beam == null || beam.ControlPoints == null || frame == null || !frame.CapturedControlPointIndex.HasValue) return null;
            var matches = beam.ControlPoints.Where(p => p != null && p.Index == frame.CapturedControlPointIndex.Value).Take(2).ToArray();
            var cp = matches.Length == 1 ? matches[0] : null;
            if (cp == null || !SameAngle(cp.GantryAngleDegrees, frame.CapturedGantryDegrees) ||
                !SameAngle(cp.PatientSupportAngleDegrees, frame.CapturedCouchDegrees)) return null;
            var iso = cp.IsocenterMm; var expected = frame.Pose.Isocenter;
            if ((!IsSynthetic || iso != null) && (iso == null || iso.Length != 3 || expected == null ||
                iso.Any(v => double.IsNaN(v) || double.IsInfinity(v)) ||
                Math.Abs(iso[0]-expected.X) > 0.01 || Math.Abs(iso[1]-expected.Y) > 0.01 || Math.Abs(iso[2]-expected.Z) > 0.01)) return null;
            return cp;
        } }
        private static bool SameAngle(double a, double b) { return !double.IsNaN(a) && !double.IsInfinity(a) && !double.IsNaN(b) && !double.IsInfinity(b) && Math.Abs(Math.IEEERemainder(a-b,360)) < 0.01; }
        public string MlcCaption { get {
            var cp = MlcControlPoint;
            return cp == null ? "Keine passende erfasste MLC-Geometrie" : string.Format(CultureInfo.InvariantCulture,
                "CP {0} · G {1:0.#}° · C {2:0.#}°", cp.Index, cp.GantryAngleDegrees, cp.CollimatorAngleDegrees);
        } }
        public string MlcScope { get { return MlcControlPoint == null ? "Geometrie fehlt oder Kontrollpunkt stimmt nicht überein." :
            SelectedFrame.Frame.Interpolated ? "Erfasster Referenz-CP zum 10°-Schritt · BLD · kein DRR" : "Erfasster Kontrollpunkt · BLD · kein DRR"; } }
        public string PoseText
        {
            get
            {
                var pose = SelectedPose;
                return pose == null ? "Keine Feldposition" : string.Format(CultureInfo.InvariantCulture,
                    "{0} · Gantry {1:0.#}° · Tisch {3:0.#}° · {2}", pose.BeamId, pose.GantryDegrees,
                    SelectedFrame.Frame.Interpolated ? "interpolierte Position" : "erfasste Position", pose.PatientSupportAngleDegrees);
            }
        }
        public IList<CollisionFinding> Findings { get { return Result == null ? new List<CollisionFinding>() : Result.Findings; } }
        public IEnumerable<CollisionFinding> SelectedFindings
        {
            get { var pose = SelectedPose; return Findings.Where(f => pose != null && f.BeamId == pose.BeamId && f.ControlPointIndex == pose.ControlPointIndex); }
        }
        public bool BodyVisible { get { return bodyVisible; } set { bodyVisible = value; Notify(); ChangeScene(); } }
        public bool TableVisible { get { return tableVisible; } set { tableVisible = value; Notify(); ChangeScene(); } }
        public bool EnvelopeVisible { get { return envelopeVisible; } set { envelopeVisible = value; Notify(); ChangeScene(); } }
        public bool UsesProfileGeometry
        {
            get
            {
                if (Scene == null || !Scene.SupportedCoordinates || Scene.Profile == null) return false;
                try { CollisionProfileCatalog.Validate(Scene.Profile, IsSynthetic); return true; }
                catch (ArgumentException) { return false; }
            }
        }
        public bool DeviceEnvelopeDrawable { get { return EnvelopeVisible && SelectedPose != null && UsesProfileGeometry; } }
        public string BoundaryNote { get { return HasSourceModel && !UsesProfileGeometry && Scene.SourceModel.Kind == "HalcyonBoreSource"
            ? "Offene Röhre · radialer Abstand" : ""; } }
        public bool SourceEnvelopeDrawable { get { return EnvelopeVisible && !DeviceEnvelopeDrawable && HasSourceModel &&
            !IsSynthetic && (Scene.SupportedCoordinates || Scene.NominalCoordinatesSupported) && SelectedPose != null &&
            (Scene.SourceModel.Kind == "TrueBeamHeadSource" || Scene.SourceModel.Kind == "HalcyonBoreSource"); } }
        public double Yaw { get { return yaw; } set { yaw = value; Notify(); ChangeScene(); } }
        public double Elevation { get { return elevation; } set { elevation = value; Notify(); ChangeScene(); } }
        public double Zoom { get { return zoom; } set { zoom = value; Notify(); ChangeScene(); } }
        public bool IncludeInReport
        {
            get { return snapshot.IncludeCollisionPreview; }
            set { snapshot.IncludeCollisionPreview = value; Notify(); ChangeScene(); }
        }
        public void SetLoading()
        {
            StopPlayback(); frames.Clear(); beamFrames = new List<CollisionSweepFrameViewModel>(); Sweep = null;
            loading = true; message = "Oberflächen und Feldpositionen werden read-only gelesen …";
            beamIds = new string[0]; beamId = null; posePosition = 0;
            snapshot.CollisionScene = null; Result = null; snapshot.CollisionPreviewPng = null;
            Notify(); ChangeScene(); CommandManager.InvalidateRequerySuggested();
        }
        public void SetFailure(string reason)
        {
            StopPlayback(); frames.Clear(); beamFrames = new List<CollisionSweepFrameViewModel>(); Sweep = null;
            loading = false; snapshot.CollisionScene = null; snapshot.CollisionPreviewPng = null;
            beamIds = new string[0]; beamId = null; posePosition = 0;
            Result = null; message = reason; Notify(); ChangeScene(); CommandManager.InvalidateRequerySuggested();
        }
        public void SetScene(CollisionScene scene, CollisionResult result, CollisionSweepResult sweep = null)
        {
            if (scene == null || scene.PlanKey != snapshot.ActivePlanKey || scene.Synthetic != snapshot.Synthetic)
                throw new ArgumentException("Collision geometry must match the active plan and source mode.");
            snapshot.CollisionScene = scene;
            StopPlayback();
            Result = result ?? CollisionEvaluator.Evaluate(scene, CancellationToken.None);
            Sweep = sweep ?? CollisionSweep.Evaluate(scene, CancellationToken.None);
            frames = Sweep.Frames.Select(f => new CollisionSweepFrameViewModel(f, HasSeparateSupport)).ToList();
            allFrames = frames.AsReadOnly();
            beamIds = scene.Poses.Select(p => p.BeamId).Distinct().ToArray();
            loading = false; beamId = BeamIds.FirstOrDefault(); posePosition = 0;
            SelectBeamFrames();
            var minimum = frames.Where(f => f.MinimumDisplayedDistance.HasValue).OrderBy(f => f.MinimumDisplayedDistance.Value).FirstOrDefault();
            if (minimum != null) { beamId = minimum.Frame.Pose.BeamId; SelectBeamFrames(); posePosition = beamFrames.IndexOf(minimum); }
            message = Sweep.Summary;
            Notify(); ChangeScene(); CommandManager.InvalidateRequerySuggested();
        }
        public void StoreIllustration(byte[] png)
        {
            if (loading || !HasScene) { snapshot.CollisionPreviewPng = null; return; }
            snapshot.CollisionPreviewPng = png;
            snapshot.CollisionPreviewCaption = (snapshot.Synthetic ? "SYNTHETIC GEOMETRY. " : "Read-only geometry snapshot. ") +
                PoseText + ". " + ProfileText + ". " + StatusLabel + ". " +
                "Evidence: " + (HasSourceModel && !UsesProfileGeometry
                    ? Scene.SourceModel.EvidenceLevel + "; source snapshot " + Scene.SourceModel.SourceSnapshotDate +
                      "; model " + Scene.SourceModel.ModelId + "; not commissioned"
                    : ProfileEvidence) + ". " + Scene.GeometryReason + ". " +
                "Visible: " + (BodyVisible ? "body; " : "") + (TableVisible ? "support; " : "") +
                (DeviceEnvelopeDrawable ? "device envelope; " : SourceEnvelopeDrawable ? "source-model envelope, nominal pose; " : "device envelope not displayed; ") +
                (SourceEnvelopeDrawable ? SourceScreeningLabel + ". Point screening is not surface clearance. Nominal illustration does not confirm pitch/roll. " +
                    (Scene.SourceModel.Kind == "HalcyonBoreSource" ? "Displayed bore length follows captured surface extent, not a measured housing length. " : "Head extent clipped by source-model reach envelope, not a measured rear face. ") : "") +
                "ISO is a projected overlay (not depth-tested). Geometry checks include all captured surfaces, independent of display visibility. " + Limitations;
        }
        public byte[] IllustrationPng { get { return snapshot.CollisionPreviewPng; } }
        public void StoreReportSweep(List<CollisionReportBeam> beams)
        {
            snapshot.CollisionBeams = !loading && HasScene && IncludeInReport ? beams : null;
        }
        private void Notify() { var handler = PropertyChanged; if (handler != null) handler(this, new PropertyChangedEventArgs(string.Empty)); }
        private void ChangeScene()
        {
            // Render is dispatched asynchronously. Never let an intervening export
            // retain an image from the previous control point/camera/visibility state.
            snapshot.CollisionPreviewPng = null; snapshot.CollisionPreviewCaption = null;
            snapshot.CollisionBeams = null;
            var handler = SceneChanged; if (handler != null) handler(this, EventArgs.Empty);
        }
        private static string Label(string status)
        {
            switch (status)
            {
                case "overlap": return "Modellüberschneidung";
                case "near": return "Innerhalb des Modellabstands";
                case "sampled-clear": return "Keine Überschneidung an geprüften Positionen";
                default: return "Nicht bewertbar";
            }
        }
    }

    public sealed class CollisionSweepFrameViewModel
    {
        private readonly bool hasSeparateSupport;
        public CollisionSweepFrameViewModel(CollisionSweepFrame frame, bool hasSeparateSupport = true) { Frame = frame; this.hasSeparateSupport = hasSeparateSupport; }
        public CollisionSweepFrame Frame { get; private set; }
        public string AngleText { get { return Frame.Pose.GantryDegrees.ToString("0.#", CultureInfo.InvariantCulture) + "°"; } }
        public string CouchText { get { return Frame.Pose.PatientSupportAngleDegrees.ToString("0.#", CultureInfo.InvariantCulture) + "°"; } }
        public string BodyStatusCode { get { return Frame.RadialReview != null ? Frame.RadialReview.BodyStatus : Frame.ModelBodyStatus ?? Frame.BodyStatus; } }
        public string TableStatusCode { get { return Frame.RadialReview != null ? Frame.RadialReview.TableStatus : Frame.ModelTableStatus ?? Frame.TableStatus; } }
        public string StatusCode { get {
            if (Frame.RadialReview != null)
            {
                if (Frame.RadialReview.Status == "hit" || BodyStatusCode == "hit" || TableStatusCode == "hit") return "hit";
                if (Frame.RadialReview.Status == "model-hit" || BodyStatusCode == "model-hit" || TableStatusCode == "model-hit") return "model-hit";
                var gap = Frame.RadialReview.BodyClearanceMm;
                if (!hasSeparateSupport && BodyStatusCode == "clear" && gap.HasValue && !double.IsInfinity(gap.Value) && gap.Value > 0)
                    return "body-clear";
                return Frame.RadialReview.Status;
            }
            if (Frame.ModelBodyStatus == null && Frame.ModelTableStatus == null) return Frame.Status;
            if (BodyStatusCode == "model-hit" || TableStatusCode == "model-hit") return "model-hit";
            if (BodyStatusCode == "model-clear" && TableStatusCode == "model-clear") return "model-clear";
            if (BodyStatusCode == "model-clear" && !hasSeparateSupport) return "body-clear";
            return BodyStatusCode == "model-clear" || TableStatusCode == "model-clear" ? "partial-clear" : "uncertain";
        } }
        public double? MinimumDisplayedDistance { get { var values = Frame.RadialReview == null
            ? new[] { Frame.BodyLowerBoundMm, Frame.TableLowerBoundMm }
            : new[] { Frame.RadialReview.BodyClearanceMm, Frame.RadialReview.TableClearanceMm };
            return values.Where(v => v.HasValue && !double.IsNaN(v.Value) && !double.IsInfinity(v.Value)).DefaultIfEmpty().Min(); } }
        public string StatusText { get { return Label(StatusCode); } }
        public string EvidenceText { get { return "Körper: " + (BodyDistanceText == "—" ? "Abstand nicht verfügbar" : BodyDistanceText + " mm") +
            " · Tisch: " + (TableDistanceText == "—" ? "Abstand nicht verfügbar" : TableDistanceText + " mm"); } }
        public string Reason { get { return Frame.RadialReview == null ? Frame.Reason : Frame.RadialReview.Scope; } }
        public string BodyResultText { get { return Label(BodyStatusCode); } }
        public string TableResultText { get { return !hasSeparateSupport ? "Kein separates Modell" : Label(TableStatusCode); } }
        public string StatusColor { get { return IsFullyClear(StatusCode) ? "#69D391" : StatusCode == "hit" || StatusCode == "model-hit" ? "#FF7B72" : "#F2C45C"; } }
        public string StatusSurface { get { return IsFullyClear(StatusCode) ? "#EAF6EE" : StatusCode == "hit" || StatusCode == "model-hit" ? "#FDEDEC" : "#FFF5DD"; } }
        public string BodyDistanceText { get { return Frame.RadialReview == null ? Distance(Frame.BodyLowerBoundMm, Frame.BodyClearanceMm)
            : Distance(Frame.RadialReview.BodyClearanceMm, Frame.RadialReview.BodyClearanceMm); } }
        public string TableDistanceText { get { return Frame.RadialReview == null ? Distance(Frame.TableLowerBoundMm, Frame.TableClearanceMm)
            : Distance(Frame.RadialReview.TableClearanceMm, Frame.RadialReview.TableClearanceMm); } }
        public static bool IsClear(string status) { return status == "pass" || status == "clear" || status == "model-clear" || status == "body-clear"; }
        public static bool IsFullyClear(string status) { return status == "pass" || status == "clear" || status == "model-clear"; }
        public static string Label(string status) { return status == "body-clear" || status == "partial-clear" ? "Still clear" : IsFullyClear(status) ? "Full clear" : status == "model-hit" ? "Modelltreffer" : status == "hit" ? "Treffer" : "Ungewiss"; }
        public static string Aggregate(IEnumerable<string> codes)
        {
            var values = codes.ToArray();
            return values.Contains("hit") ? "hit" : values.Contains("model-hit") ? "model-hit" : values.Length == 0 ? "uncertain" :
                values.All(v => v == "clear") ? "clear" : values.All(v => v == "pass") ? "pass" :
                values.All(IsClear) ? (values.Contains("body-clear") ? "body-clear" : "model-clear") : values.All(v => v == "partial-clear" || IsClear(v)) ? "partial-clear" : "uncertain";
        }
        public static string Distance(double? lower, double? upper)
        {
            if (!upper.HasValue || double.IsNaN(upper.Value) || double.IsInfinity(upper.Value)) return "—";
            if (lower.HasValue && Math.Abs(lower.Value - upper.Value) > 0.1)
                return lower.Value.ToString("0.0", CultureInfo.InvariantCulture) + " … " + upper.Value.ToString("0.0", CultureInfo.InvariantCulture);
            return (!lower.HasValue ? "≤ " : "") + upper.Value.ToString("0.0", CultureInfo.InvariantCulture);
        }
    }
}
