using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using ClearPlan;
using ClearPlan.Helpers;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Collision;
using ClearPlan.Core.Review;
using ClearPlan.Core.Localization;
using ClearPlan.Presentation.ViewModels;
using ClearPlan.Presentation.Views;
using ClearPlan.Reporting;
using ClearPlan.Reporting.MigraDoc;
using ClearPlan.Review;
using EsapiEssentials.Plugin;
using Newtonsoft.Json;
using VMS.TPS.Common.Model.API;

[assembly: ESAPIScript(IsWriteable = false)]

namespace ClearPlan.Headless
{
    // Explicit patient-scoped engineering inventory, protected destination only.
    // No patient identifiers or anatomy, no automatic plan choice or TPS changes.
    public sealed class PatientPlanInventory
    {
        public void Run(PluginScriptContext context)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA || context == null || context.Patient == null)
                throw new InvalidOperationException("An explicit patient on the owner STA is required.");
            var rows = context.Patient.Courses.SelectMany(c => c.PlanSetups.Select(p => new {
                Course = c.Id, Plan = p.Id, Fractions = p.NumberOfFractions, Dose = p.TotalDose.ToString(),
                ValidDose = p.IsDoseValid, Status = p.ApprovalStatus.ToString(),
                Machines = p.Beams.Where(b => !b.IsSetupField).Select(b => b.TreatmentUnit.Id).Distinct().ToArray(),
                Fields = p.Beams.Count(b => !b.IsSetupField) })).ToArray();
            string path = Path.Combine(ClearPlanSettings.Load().GetReportsDirectory(), "plan-inventory-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + ".json");
            using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
            using (var writer = new StreamWriter(file, new System.Text.UTF8Encoding(false))) writer.Write(JsonConvert.SerializeObject(rows, Formatting.Indented));
        }
    }

    // Explicit isolated Runner entry: no window, Save dialog, patient mutation or simulator fallback.
    // The host owns the patient lifetime; its STA is pumped until all detached work has completed.
    public sealed class ReportScript
    {
        public void Run(PluginScriptContext context)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("The ESAPI owner must be STA.");
            if (context == null || context.Patient == null || context.PlanSetup == null)
                throw new InvalidOperationException("An explicit patient and plan context is required.");
            LoadReviewLanguage();
            var dispatcher = Dispatcher.CurrentDispatcher;
            var previousContext = SynchronizationContext.Current;
            SynchronizationContext.SetSynchronizationContext(new DispatcherSynchronizationContext(dispatcher));
            try
            {
                Task work = ExportAsync(context);
                if (!work.IsCompleted)
                {
                    var frame = new DispatcherFrame();
                    work.ContinueWith(t => dispatcher.BeginInvoke(new Action(() => frame.Continue = false)), TaskScheduler.Default);
                    Dispatcher.PushFrame(frame);
                }
                work.GetAwaiter().GetResult();
            }
            finally { SynchronizationContext.SetSynchronizationContext(previousContext); }
        }

        private static void LoadReviewLanguage()
        {
            ReviewLanguage.LoadPreference();
            // Initialize the XAML language singleton before the candidate override;
            // its initial preference load must not reset the requested capture language.
            ReviewLanguage.Code = ClearPlan.Presentation.Localization.DisplayLanguage.Instance.Code;
            string path = Path.Combine(Path.GetDirectoryName(typeof(ReportScript).Assembly.Location), "review-language.txt");
            if (!File.Exists(path)) return;
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (stream.Length > 16) throw new InvalidDataException("Candidate review language must be de or en.");
                using (var reader = new StreamReader(stream, new System.Text.UTF8Encoding(false, true), true))
                    ReviewLanguage.Code = ReviewLanguage.Normalize(reader.ReadToEnd().Trim());
            }
            // No preference write or environment mutation: override is confined to this process.
        }

        private static async Task ExportAsync(PluginScriptContext context)
        {
            var plan = context.PlanSetup;
            var reviewLogStartedUtc = DateTimeOffset.UtcNow;
            var settings = ClearPlanSettings.Load();
            // The isolated candidate's settings.ini defines the protected destination.
            string directory = settings.GetReportsDirectory();
            string token = new string(plan.Id.Select(c => char.IsLetterOrDigit(c) || c == '-' || c == '_' ? c : '_').ToArray());
            string basename = "ClearPlan-" + token + "-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
            string trace = Path.Combine(directory, basename + ".technical.log");
            Action<string> stage = value => File.AppendAllText(trace, DateTime.UtcNow.ToString("O") + " " + value + Environment.NewLine);
            stage("begin");
            try
            {
                bool hasDose = plan.Dose != null;
                bool? doseValid = null;
                try { doseValid = plan.IsDoseValid; } catch (Exception) { }
                double initialNormalization = plan.PlanNormalizationValue;
                var item = new PlanningItemViewModel(plan);
                var vm = new MainViewModel(context.CurrentUser, context.Patient, "3.1.0 development PDF preview",
                    new ObservableCollection<PlanningItemViewModel> { item }, item, plan);
                ClearPlan.PhantomValidation.Selection.Apply(vm, context.Patient, plan);
                stage("review-ready");
                File.WriteAllText(Path.Combine(directory, basename + ".review-diagnostics.json"),
                    JsonConvert.SerializeObject(CaptureReviewDiagnostics(plan), Formatting.Indented));
                vm.PqmSummaries = new ObservableCollection<PQMSummaryViewModel>();
                // Preserve unresolved objectives; do not silently approve a suggested table.
                if (vm.ActiveConstraintPath != null &&
                    vm.ConstraintCatalogStatus != null && vm.ConstraintCatalogStatus.IsUsable)
                {
                    vm.GetPQMSummaries(vm.ActiveConstraintPath, item, context.Patient);
                }
                vm.ApplyDefaultDvhSelections();
                var snapshot = new EsapiReviewSnapshotBuilder().Build(vm, settings);
                stage("snapshot-ready");
                if (vm.ConstraintSelectionRequiresConfirmation || vm.PqmSummaries.Count == 0)
                {
                    snapshot.PlanCheckRows.Insert(0, new ReviewCheckRow {
                        CheckCode="constraint-selection-unconfirmed", Category="PQM", Status=ReviewStatusCodes.Variation,
                        Severity=ReviewSeverityCodes.Warning, Unit=ReviewUnitCodes.Text,
                        Message="RefDB selection needs review: " + (vm.ActiveConstraintPath == null ? "No constraint table available." : vm.ActiveConstraintPath.SelectionReason) + " Native Eclipse Clinical Goals and the user target rule are listed independently."
                    });
                    snapshot.Report.Notes.Add("For unconfirmed RefDB objectives, values are measured on the active plan dose without scaling; goal assessment awaits confirmation of the table and treatment scope. Native Eclipse Clinical Goals retain their own source and evaluation.");
                }
                NativeMlcProfileCatalog profiles = null;
                try { profiles = await NativeMlcProfileFileSource.LoadAsync(settings.ResolvePath(settings.Paths.MlcGeometryProfilesJsonPath), CancellationToken.None); }
                catch (Exception) { snapshot.Report.Notes.Add("Native MLC geometry profile unavailable; no physical leaf layout was inferred."); }
                DoseRateEstimationProfileCatalog rateProfiles = null;
                try { rateProfiles = await DoseRateProfileFileSource.LoadAsync(string.IsNullOrWhiteSpace(settings.Paths.DoseRateProfilesJsonPath) ? null : settings.ResolvePath(settings.Paths.DoseRateProfilesJsonPath), CancellationToken.None); }
                catch (Exception) { snapshot.Report.Notes.Add("Dose-rate estimate profile unavailable; no machine timing was inferred."); }
                var analysisBuilder = new EsapiPlanAnalysisBuilder(profiles, rateProfiles);
                var analysis = analysisBuilder.Build(plan);
                stage("native-analysis-ready");
                File.WriteAllText(Path.Combine(directory, basename + ".frame-diagnostics.json"),
                    JsonConvert.SerializeObject(CaptureFrameDiagnostics(plan, analysis.TargetStructureId), Formatting.Indented));
                analysis = await analysisBuilder.EnrichTargetProjectionsAsync(plan, analysis, CancellationToken.None);
                stage("target-projections-ready");
                // The DRR snapshot guard deliberately requires finite dose metadata.
                // A dose-less diagnostic report can still show geometry, but must
                // not relax that freshness guard or fail for this optional image.
                if (hasDose && doseValid == true)
                    analysis = await new EsapiBevBuilder().BuildStartsAsync(plan, analysis, CancellationToken.None);
                else
                {
                    foreach (var beam in analysis.Beams)
                    {
                        var start = beam.ControlPoints.FirstOrDefault();
                        if (start == null) continue;
                        start.BevImage = new BeamEyeViewImage { SourceStatus = "unavailable", Synthetic = false,
                            UnavailableReason = "DRR omitted: no valid native dose; dose-bound snapshot freshness cannot be established.",
                            ControlPointIndex = start.Index, GantryAngleDegrees = start.GantryAngleDegrees,
                            CollimatorAngleDegrees = start.CollimatorAngleDegrees, PatientSupportAngleDegrees = start.PatientSupportAngleDegrees };
                    }
                    stage("bev-omitted-no-valid-native-dose");
                }
                stage("bev-capture-finished");
                snapshot.PlanAnalysis = analysis;
                var dosePresentationBeforeImages = plan.DoseValuePresentation;
                snapshot.PlanImages = EsapiPlanImageBuilder.Build(plan, snapshot.ActivePlanKey);
                if (plan.DoseValuePresentation != dosePresentationBeforeImages)
                    throw new InvalidOperationException("CT export did not restore the native dose display context.");
                stage("ct-overview-finished");
                CollisionResult collisionResult = null;
                CollisionSweepResult collisionSweep = null;
                using (var collisionTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(40)))
                {
                    try
                    {
                        string collisionPath = settings.Paths.CollisionProfilesJsonPath;
                        CollisionProfileCatalog collisionProfiles = null;
                        try
                        {
                            collisionProfiles = await CollisionProfileFileSource.LoadAsync(
                                string.IsNullOrWhiteSpace(collisionPath) ? null : settings.ResolvePath(collisionPath), collisionTimeout.Token);
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception) { stage("collision-profile-unavailable"); }
                        SourceCollisionModelCatalog sourceModels = null;
                        string sourceModelPath = settings.Paths.SourceCollisionModelsJsonPath;
                        try
                        {
                            sourceModels = await CollisionProfileFileSource.LoadSourceModelsAsync(
                                string.IsNullOrWhiteSpace(sourceModelPath) ? null : settings.ResolvePath(sourceModelPath), collisionTimeout.Token);
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception) { stage("collision-source-model-unavailable"); }
                        // This entry has no operator confirmation dialog. Never manufacture
                        // the per-plan zero pitch/roll assurance required by native ESAPI.
                        // Actual surfaces remain inspectable even when clearance is unavailable.
                        var collisionScene = EsapiCollisionBuilder.Capture(plan, snapshot.ActivePlanKey,
                            collisionProfiles, false, collisionTimeout.Token, stage);
                        var collisionEvaluation = await Task.Run(() => {
                            SourceCollisionSceneAdapter.Attach(collisionScene, sourceModels, collisionTimeout.Token);
                            return new { Result = CollisionEvaluator.Evaluate(collisionScene, collisionTimeout.Token),
                                Sweep = CollisionSweep.Evaluate(collisionScene, collisionTimeout.Token) };
                        }, collisionTimeout.Token);
                        collisionResult = collisionEvaluation.Result;
                        collisionSweep = collisionEvaluation.Sweep;
                        collisionTimeout.Token.ThrowIfCancellationRequested();
                        var collisionModel = new CollisionViewModel(snapshot);
                        collisionModel.SetScene(collisionScene, collisionResult, collisionEvaluation.Sweep);
                        collisionModel.IncludeInReport = true;
                        if (collisionModel.WorstPoseCommand.CanExecute(null)) collisionModel.WorstPoseCommand.Execute(null);
                        CollisionView.CaptureReportSweep(collisionModel);
                        // Synchronous WPF calls cannot be preempted. Honor the cooperative
                        // deadline after rendering too; never accept a late illustration.
                        collisionTimeout.Token.ThrowIfCancellationRequested();
                        if (snapshot.CollisionPreviewPng == null)
                            snapshot.Report.Notes.Add("Collision / 3D requested: no usable evaluated sweep illustration. " + collisionEvaluation.Sweep.Summary);
                        stage(snapshot.CollisionPreviewPng == null ? "collision-illustration-unavailable" : "collision-illustration-ready");
                    }
                    catch (OperationCanceledException)
                    {
                        snapshot.CollisionPreviewPng = null; snapshot.CollisionBeams = null;
                        snapshot.Report.Notes.Add("Collision / 3D unavailable: capture or evaluation exceeded its time limit. No collision clearance is asserted.");
                        stage("collision-timeout");
                    }
                    catch (Exception error)
                    {
                        snapshot.CollisionPreviewPng = null; snapshot.CollisionBeams = null;
                        snapshot.Report.Notes.Add("Collision / 3D unavailable: native surfaces or illustration could not be captured safely. No collision clearance is asserted.");
                        stage("collision-capture-unavailable");
                        // Technical type/method only: exception messages may contain native identifiers.
                        stage("collision-error-type=" + error.GetType().Name + " method=" + (error.TargetSite == null ? "unknown" : error.TargetSite.Name));
                    }
                }
                if (!plan.PlanNormalizationValue.Equals(initialNormalization) || plan.IsDoseValid != doseValid)
                    throw new InvalidOperationException("Plan normalization or dose validity changed during export.");
                var document = new ReviewSnapshotReportMapper().Map(snapshot);
                document.Title = ReviewReportDocument.ReportTitle;
                document.Subtitle = "Read-only development PDF preview";
                document.PatientDisplayLabel = "Native Eclipse patient (identity omitted from diagnostic JSON)";
                document.Notes.Add("Native read-only ESAPI export through the published Citrix Runner; no DICOM export or synthetic anatomy.");
                document.Notes.Add("MLC geometry and DRR orientation require workstation acceptance. Unavailable physical layers are not replaced by guessed leaf widths or index mappings.");
                ReportDoseSafety.Apply(document, hasDose, doseValid);
                ReportDvhSelection.ApplyIfEmpty(document);
                string pdfPath = Path.Combine(directory, basename + ".pdf");
                if (File.Exists(pdfPath)) throw new IOException("The selected output already exists.");
                // Protected detached intermediate for local report layout verification; never publish clinical anatomy.
                string sourceJsonPath=Path.Combine(directory, basename + ".report.json");
                byte[] sourceJson=new System.Text.UTF8Encoding(false).GetBytes(JsonConvert.SerializeObject(document));
                File.WriteAllBytes(sourceJsonPath,sourceJson);
                ProtectedReportSidecar.Export(sourceJsonPath,sourceJson,document,context.Patient.Id);
                // Requested identity belongs only in protected PDF/HTML reports, never diagnostic JSON/logs.
                document.PatientDisplayLabel = context.Patient.LastName + ", " + context.Patient.FirstName + " | ID: " + context.Patient.Id;
                new ReportPdf().Export(pdfPath, document);
                string htmlPath = new HtmlReviewReportRenderer().ExportCompanion(pdfPath, document);
                File.WriteAllText(Path.Combine(directory, basename + ".verification.json"), JsonConvert.SerializeObject(new {
                    LanguageCode = ReviewLanguage.Code,
                    ReportFile=Path.GetFileName(pdfPath), HtmlReportFile=Path.GetFileName(htmlPath), NativeDosePresent=hasDose, NativeDoseValid=doseValid,
                    PqmCount=document.PqmRows.Count, PlanCheckCount=document.PlanCheckRows.Count,
                    SelectedDvhCount=document.DvhSeries.Count(s=>s.Selected),
                    Collision = new { Requested=true, IllustrationPresent=document.CollisionPreviewPng != null,
                        Status=collisionResult == null ? "unavailable" : collisionResult.Status,
                        SweepStatus=collisionSweep == null ? "unavailable" : collisionSweep.Status,
                        SweepSummary=collisionSweep == null ? null : collisionSweep.Summary,
                        SweepFrameCount=collisionSweep == null ? 0 : collisionSweep.Frames.Count,
                        ZeroPitchRollOperatorConfirmed=false,
                        SourceModel=snapshot.CollisionScene == null || snapshot.CollisionScene.SourceModel == null ? null : snapshot.CollisionScene.SourceModel.ModelId,
                        SourceScreeningStatus=snapshot.CollisionScene == null || snapshot.CollisionScene.SourceScreening == null ? "unavailable" : snapshot.CollisionScene.SourceScreening.Status,
                        ClinicalClearanceSupported=false,
                        SurfaceCount=snapshot.CollisionScene == null ? 0 : snapshot.CollisionScene.Surfaces.Count,
                        Surfaces=snapshot.CollisionScene == null ? null : snapshot.CollisionScene.Surfaces.Select(s=>new {
                            s.Role, Vertices=s.Mesh.Vertices.Count, Triangles=s.Mesh.TriangleIndices.Length / 3 }),
                        PoseCount=snapshot.CollisionScene == null ? 0 : snapshot.CollisionScene.Poses.Count },
                    SelectedDvhStructures=document.DvhSeries.Where(s=>s.Selected).Select(s=>s.StructureId),
                    PlanAnalysis=new {analysis.Pam,analysis.PamStatus,analysis.PamReason,analysis.TargetStructureId,
                        analysis.PamWeightingMode,analysis.TargetSelectionProvenance,
                        Beams=analysis.Beams.Select(b=>new {b.BeamNumber,b.Pam,b.PamReason,b.PamBeamWeightFactor,
                            Samples=b.ControlPoints.Select(c=>new {c.Index,c.BlockedTargetFraction,c.NativeProjectionDifferenceFraction,c.TargetProjectionReason})})},
                    ConstraintTable=vm.ActiveConstraintPath == null ? null : vm.ActiveConstraintPath.ConstraintName,
                    ConstraintSelectionReason=vm.ActiveConstraintPath == null ? "No constraint table available." : vm.ActiveConstraintPath.SelectionReason,
                    ConstraintSelectionRequiresConfirmation=vm.ConstraintSelectionRequiresConfirmation,
                    NativeGoalSource=snapshot.Sources.FirstOrDefault(s=>s.SourceCode=="eclipse-clinical-goals"),
                    NativeWarningsSource=snapshot.Sources.FirstOrDefault(s=>s.SourceCode=="eclipse-validation-messages"),
                    UserTargetSource=snapshot.Sources.FirstOrDefault(s=>s.SourceCode=="user-target-coverage"),
                    DvhTargets=document.DvhSeries.Where(s=>s.Role==ReviewDvhRoleCodes.Target).Select(s=>new {s.StructureId,s.TargetKind,s.Selected,s.RequiredForTargetReview,s.TargetSelectionReason}),
                    PqmRows=document.PqmRows.Select(p=>new {p.TemplateStructure,p.ResolvedStructureId,p.Objective,p.Comparator,p.Goal,p.Unit,p.AchievedValue,p.Status,p.SourceLabel,p.MappingDescription}),
                    MlcCapture=NativeMlcDiagnostics.Capture(plan),
                    DosePresentationRestored=plan.DoseValuePresentation == dosePresentationBeforeImages,
                    CtImages=document.PlanImages.Select(i=>new {i.Kind,i.SourceStatus,i.UnavailableReason,i.OverlaySummary,
                        i.DoseFocusRegion, Viewport=i.SourceStatus==ReviewStatusCodes.Available ? PlanImageViewport.Calculate(i).Description : null,
                        Overlays=i.Overlays.Select(o=>new {o.Kind,o.Label,o.SourceStatus,o.UnavailableReason,PathCount=o.Paths.Count})}),
                    Beams=analysis.Beams.Select(b=>new {b.BeamNumber,b.GeometryStatus,b.GeometryReason,
                        DrrStatus=b.ControlPoints.Count==0 || b.ControlPoints[0].BevImage==null ? "missing" : b.ControlPoints[0].BevImage.SourceStatus,
                        DrrReason=b.ControlPoints.Count==0 || b.ControlPoints[0].BevImage==null ? null : b.ControlPoints[0].BevImage.UnavailableReason,
                        PhysicalLayers=b.ControlPoints.Count==0 || b.ControlPoints[0].Aperture==null ? null : b.ControlPoints[0].Aperture.Layers,
                        IsocenterMm=b.ControlPoints.Count==0 ? null : b.ControlPoints[0].IsocenterMm}),
                    NativeMlc=plan.Beams.Where(b=>!b.IsSetupField).Select(b=>new {b.BeamNumber,
                        Model=b.MLC==null ? "none" : b.MLC.Model,
                        Rows=b.ControlPoints.First().LeafPositions==null ? 0 : b.ControlPoints.First().LeafPositions.GetLength(0),
                        Columns=b.ControlPoints.First().LeafPositions==null ? 0 : b.ControlPoints.First().LeafPositions.GetLength(1)})
                }, Formatting.Indented));
                // Structured review log (mode = headless); guarded, never fails the export.
                string reviewLogFailure;
                var reviewLog = EsapiReviewLog.CreateContext(ClearPlan.Core.ReviewLog.ReviewLogModes.Headless, context.CurrentUser,
                    context.Patient, plan, EsapiReviewLog.GoalTableId(vm), reviewLogStartedUtc);
                reviewLog.ReportWritten = true;
                reviewLog.ReportSha256 = EsapiReviewLog.FileSha256OrNull(pdfPath);
                bool reviewLogged = EsapiReviewLog.TryWrite(snapshot, reviewLog, settings, EsapiReviewLog.ActiveConstraintSource(vm), 10000, out reviewLogFailure);
                stage(reviewLogged ? "review-log-written" : "review-log-not-written=" + reviewLogFailure);
                snapshot.PlanImages.Clear();
                stage("completed");
            }
            catch (Exception error)
            {
                stage("failed: " + error.GetType().FullName + Environment.NewLine + error.StackTrace);
                throw;
            }
        }

        private static object CaptureReviewDiagnostics(PlanSetup plan)
        {
            var result = new Dictionary<string, object>();
            try { var rx = plan.RTPrescription; result["Prescription"] = rx == null ? null : new { rx.Id, rx.Name, rx.Site, rx.NumberOfFractions }; }
            catch (Exception error) { result["PrescriptionReadError"] = error.GetType().Name; }
            try
            {
                var image = plan.StructureSet.Image;
                var iso = plan.Beams.First(b => !b.IsSetupField).IsocenterPosition;
                double x0 = Math.Min(image.Origin.x, image.Origin.x + (image.XSize - 1) * image.XRes * image.XDirection.x);
                double x1 = Math.Max(image.Origin.x, image.Origin.x + (image.XSize - 1) * image.XRes * image.XDirection.x);
                double y0 = Math.Min(image.Origin.y, image.Origin.y + (image.YSize - 1) * image.YRes * image.YDirection.y);
                var samples = new List<object>();
                foreach (double y in new[] { y0, iso.y })
                {
                    var profile = plan.Dose.GetDoseProfile(new VMS.TPS.Common.Model.Types.VVector(x0, y, iso.z),
                        new VMS.TPS.Common.Model.Types.VVector(x1, y, iso.z), new double[192]);
                    samples.Add(profile == null ? null : new {
                        profile.Unit, profile.Count, ExpectedStart = new[] { x0, y, iso.z }, ExpectedEnd = new[] { x1, y, iso.z },
                        Samples = Enumerable.Range(0, Math.Min(3, profile.Count)).Select(i => new { profile[i].Position, profile[i].Value }).ToArray()
                    });
                }
                result["DoseProfiles"] = samples;
            }
            catch (Exception error) { result["DoseProfileReadError"] = error.GetType().Name; }
            return result;
        }

        // Protected, local diagnostic evidence. Never include these native coordinates in public artifacts.
        private static object CaptureFrameDiagnostics(PlanSetup plan, string targetId)
        {
            var rows = new List<object>();
            var structure = plan.StructureSet.Structures.FirstOrDefault(s => s.Id == targetId && !s.IsEmpty && s.HasSegment)
                ?? plan.StructureSet.Structures.SingleOrDefault(s => s.DicomType == "EXTERNAL" && !s.IsEmpty && s.HasSegment);
            foreach (var beam in plan.Beams.Where(b => !b.IsSetupField))
            {
                var row = new Dictionary<string, object>();
                rows.Add(row);
                row["BeamNumber"] = beam.BeamNumber;
                string step = "calibration-structure";
                try
                {
                    if (structure == null) throw new InvalidOperationException("No calibration structure.");
                    step = "native-outline-false";
                    var beamOutline = beam.GetStructureOutlines(structure, false);
                    row["BeamOutline"] = beamOutline;
                    step = "native-outline-true";
                    var bevOutline = beam.GetStructureOutlines(structure, true);
                    row["BevOutline"] = bevOutline;
                    step = "rotation-fit";
                    double rotation = MeshTargetProjector.NativeCoordinateRotation(
                        bevOutline.Select(l => l.Select(p => new BeamPoint(p.X, p.Y)).ToList()).ToList(),
                        beamOutline.Select(l => l.Select(p => new BeamPoint(p.X, p.Y)).ToList()).ToList(),beam.ControlPoints.First().CollimatorAngle);
                    row["RotationRadians"] = rotation;
                    step = "native-source-locations";
                    Func<VMS.TPS.Common.Model.Types.VVector, BeamPoint3D> point = p => new BeamPoint3D(p.x, p.y, p.z);
                    var iso = point(beam.IsocenterPosition);
                    var source = point(beam.GetSourceLocation(beam.ControlPoints.First().GantryAngle));
                    var zero = point(beam.GetSourceLocation(0));
                    var ninety = point(beam.GetSourceLocation(90));
                    row["Iso"] = iso; row["Source"] = source; row["SourceZero"] = zero; row["SourceNinety"] = ninety;
                    step = "create-frame";
                    row["Frame"] = MeshTargetProjector.CreateFrame(iso, source, zero, ninety, rotation);
                    row["Status"] = "available";
                }
                catch (Exception error)
                {
                    row["Status"] = "failed"; row["Step"] = step;
                    row["ErrorType"] = error.GetType().FullName;
                    row["ErrorMessage"] = error.Message;
                }
            }
            return rows;
        }
    }
}
