using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Review
{
    /// <summary>Read-only extraction. Call synchronously on the ESAPI owner thread.</summary>
    public sealed class EsapiPlanAnalysisBuilder
    {
        private readonly NativeMlcProfileCatalog nativeProfiles;
        private readonly DoseRateEstimationProfileCatalog doseRateProfiles;
        public EsapiPlanAnalysisBuilder(NativeMlcProfileCatalog profiles = null, DoseRateEstimationProfileCatalog rateProfiles = null)
        { nativeProfiles = profiles; doseRateProfiles = rateProfiles; }
        /// <summary>
        /// Explicit, potentially expensive PAM workflow. Capture is synchronous on the owner
        /// thread. The worker closes over detached snapshots only, never ESAPI objects.
        /// The returned snapshot is a copy; a canceled/stale operation cannot mutate the UI.
        /// </summary>
        public Task<ReviewPlanAnalysis> EnrichTargetProjectionsAsync(PlanSetup plan,
            ReviewPlanAnalysis analysis, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (System.Windows.Application.Current != null)
                System.Windows.Application.Current.Dispatcher.VerifyAccess();
            if (plan == null) throw new ArgumentNullException("plan");
            cancellationToken.ThrowIfCancellationRequested();
            var copy = PlanAnalysisSnapshot.Copy(analysis);
            if (copy.TargetSelectionMode == "AutomaticLowestPam")
            {
                var candidates = CaptureAutomaticCandidates(plan, copy, cancellationToken);
                if (candidates.Count > 0)
                {
                    // Only detached analyses, primitive meshes and frames cross this boundary.
                    return Task.Run(() => CalculateAutomaticCandidates(candidates, copy, cancellationToken));
                }
            }
            List<ProjectionJob> jobs = CaptureProjectionJobs(plan,copy,cancellationToken);
            // Do not capture plan, beam, target, native mesh, or other ESAPI objects here.
            return Task.Run(() => CalculateDetachedProjections(copy,jobs,cancellationToken));
        }

        public ReviewPlanAnalysis Build(PlanSetup plan) { return Build(plan, null); }

        public ReviewPlanAnalysis Build(PlanSetup plan, string targetStructureId)
        {
            if (System.Windows.Application.Current != null)
                System.Windows.Application.Current.Dispatcher.VerifyAccess();
            if (plan == null) throw new ArgumentNullException("plan");
            var result = new ReviewPlanAnalysis
            {
                TargetSelectionMode = string.IsNullOrWhiteSpace(targetStructureId) ? "AutomaticLowestPam" : "Explicit",
                PamWeightingMode = "BeamWeightFactor",
                GeometryProvenance = "Native ESAPI treatment beams excluding setup and imaging-treatment fields, matching PlanCheck beam scope. Control-point positions with explicitly configured physical leaf boundaries/layer-index mapping. Unrecognized models or incomplete layers remain unavailable; no physical layout is inferred. " + PlanAnalysisCalculator.SamplingDefinition,
                DoseRateProvenance = "Estimated plan trajectory: PlanCheck-style segment averages from control-point MU, gantry angles, a configured speed assumption and the nominal beam MU/min cap. Not ESAPI-supplied timing or measured delivery. Acceleration, MLC/jaw motion, dose ramping and beam holds are not modeled."
            };
            if (!(plan is ExternalPlanSetup))
            {
                result.PamReason = "Only external photon plan analysis is supported.";
                return result;
            }
            try { result.DosePerFractionGy = DoseGy(plan.DosePerFraction); }
            catch (Exception) { result.Warnings.Add("Prescribed dose per fraction could not be read."); }
            try { if (Finite(plan.PlanNormalizationValue)) result.PlanNormalizationPercent = plan.PlanNormalizationValue; }
            catch (Exception) { result.Warnings.Add("TPS plan normalization is unavailable."); }
            try { result.NativePlanFingerprint=EsapiNativeGeometryFingerprint.CapturePlan(plan); }
            catch(Exception) { result.NativePlanFingerprint=null;result.Warnings.Add("Native plan freshness could not be captured; report DRRs require a refreshed valid snapshot."); }
            try
            {
                if (plan.StructureSet != null)
                {
                    result.AvailableTargetStructureIds = plan.StructureSet.Structures
                        .Where(s => !s.IsEmpty && s.HasSegment).Select(s => s.Id).OrderBy(id => id,StringComparer.OrdinalIgnoreCase).ToList();
                    if (result.TargetSelectionMode == "AutomaticLowestPam")
                        result.PamTargetCandidateCount = AutomaticPtvIds(plan).Count;
                }
            }
            catch (Exception) { result.Warnings.Add("Selectable target structure list could not be read."); }
            Structure target = null;
            string targetReason, targetProvenance = null;
            try { target = SelectTarget(plan, targetStructureId, out targetProvenance, out targetReason); }
            catch (Exception) { targetReason = "Target structure selection could not be read."; }
            result.TargetSelectionProvenance = targetProvenance;
            if (target != null) result.TargetStructureId = target.Id;
            else result.Warnings.Add(targetReason);
            foreach (var beam in plan.Beams.Where(b => !b.IsSetupField && !b.IsImagingTreatmentField))
            {
                var row = new ReviewBeamAnalysis();
                result.Beams.Add(row);
                // Do not propagate exception messages: vendor messages can include patient data.
                try
                {
                    // Stamp independently of physical-model recognition: unknown models remain representable as unavailable apertures.
                    try { row.NativeGeometryFingerprint=EsapiNativeGeometryFingerprint.CaptureBeam(beam); }
                    catch(Exception) { row.NativeGeometryFingerprint=null; }
                    ExtractBeam(beam, target, targetReason, row);
                    EsapiNativeMlcAdapter.ApplyToBeam(beam, row, nativeProfiles);
                }
                catch (Exception)
                {
                    row.NativeGeometryFingerprint=null;
                    row.GeometryStatus = "unavailable";
                    row.GeometryReason = "Beam analysis extraction failed; incomplete values are not used for plan metrics.";
                    row.ControlPoints.Clear();
                    row.Pam = null;
                }
            }
            PlanAnalysisCalculator.Calculate(result);
            foreach (var row in result.Beams) DoseRateEstimator.Apply(row, doseRateProfiles);
            try { EsapiTargetQualityBuilder.Populate(plan, result); }
            catch (Exception) { result.TargetQuality.Clear(); result.TargetQualityNote = "Target quality unavailable: native DVH inputs could not be read. No fallback values used."; }
            if (target == null) result.PamReason = targetReason;
            return result;
        }

        private static List<string> AutomaticPtvIds(PlanSetup plan)
        {
            if (plan.StructureSet == null) return new List<string>();
            var eligible = plan.StructureSet.Structures.Where(s => !s.IsEmpty && s.HasSegment).ToList();
            return PamTargetSelection.AutomaticCandidateIds(eligible.Select(s => s.Id),
                eligible.Where(s => DvhSelectionPolicy.ClassifyTarget(s.Id, s.DicomType) == "PTV").Select(s => s.Id));
        }

        // This entire method runs synchronously on the owning STA. Never rebuild beam
        // apertures here: the supplied snapshot may contain matched imported dual-layer MLC data.
        private static List<AutomaticCandidate> CaptureAutomaticCandidates(PlanSetup plan,
            ReviewPlanAnalysis analysis, CancellationToken cancellationToken)
        {
            var candidates = new List<AutomaticCandidate>();
            foreach (string id in AutomaticPtvIds(plan))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var copy = PlanAnalysisSnapshot.Copy(analysis);
                copy.TargetStructureId = id;
                copy.TargetSelectionProvenance = null;
                copy.PamTargetCandidates.Clear();
                string provenance, reason;
                Structure target = SelectTarget(plan, id, out provenance, out reason);
                var nativeBeams = plan.Beams.Where(b => !b.IsSetupField && !b.IsImagingTreatmentField).ToList();
                foreach (var row in copy.Beams)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    foreach (var cp in row.ControlPoints)
                    {
                        cp.TargetOutlines.Clear();
                        cp.TargetProjectionStrips.Clear();
                        cp.TargetProjectionProvenance = null;
                        cp.TargetProjectionResolutionMm = null;
                        cp.NativeProjectionDifferenceFraction = null;
                        cp.TargetProjectionReason = "Target projection has not been captured for this PTV candidate.";
                        cp.BevImage = null;
                    }
                    try
                    {
                        var matches = nativeBeams.Where(b => row.BeamNumber.HasValue && b.BeamNumber == row.BeamNumber.Value).ToList();
                        if (target == null || matches.Count != 1) continue;
                        var beam = matches[0];
                        var points = beam.ControlPoints.ToList();
                        if (!MatchesSnapshot(beam, points, row)) continue;
                        var outlines = CopyOutline(beam.GetStructureOutlines(target, false));
                        bool fixedProjection = points.All(p => SameProjection(points[0], p));
                        foreach (var cp in row.ControlPoints)
                            if (fixedProjection || cp.Index == 0)
                            {
                                cp.TargetOutlines = outlines;
                                cp.TargetProjectionReason = outlines.Count > 0 ? null : "Native target BEV projection is empty.";
                                cp.TargetProjectionProvenance = "Native Beam.GetStructureOutlines(target, false), candidate-specific beam-start geometry.";
                            }
                    }
                    catch (Exception)
                    {
                        foreach (var cp in row.ControlPoints)
                            cp.TargetProjectionReason = "Native target outline could not be captured for this PTV candidate.";
                    }
                }
                candidates.Add(new AutomaticCandidate { Analysis = copy,
                    Jobs = CaptureProjectionJobs(plan, copy, cancellationToken) });
            }
            return candidates;
        }

        private static ReviewPlanAnalysis CalculateAutomaticCandidates(List<AutomaticCandidate> candidates,
            ReviewPlanAnalysis fallback, CancellationToken cancellationToken)
        {
            using (var bounded = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
            {
                // One budget for the complete comparison, not 30 seconds per PTV.
                bounded.CancelAfter(TimeSpan.FromSeconds(30));
                foreach (var candidate in candidates)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    CalculateDetachedProjections(candidate.Analysis, candidate.Jobs, bounded.Token);
                }
            }
            cancellationToken.ThrowIfCancellationRequested();
            return PamTargetSelection.SelectLowestAvailable(candidates.Select(c => c.Analysis), fallback);
        }

        private sealed class AutomaticCandidate
        {
            internal ReviewPlanAnalysis Analysis;
            internal List<ProjectionJob> Jobs;
        }

        private static List<ProjectionJob> CaptureProjectionJobs(PlanSetup plan,
            ReviewPlanAnalysis copy,CancellationToken cancellationToken)
        {
            var jobs=copy.Beams.Select(b=>new ProjectionJob { Row=b }).ToList();
            // Build owns target resolution. Do not replace a failed explicit selection
            // with the automatic PTV policy while enriching its snapshot.
            if (string.IsNullOrWhiteSpace(copy.TargetStructureId))
            {
                foreach (var job in jobs) job.Error=copy.PamReason ?? "No resolved PAM target in the analysis snapshot.";
                return jobs;
            }
            string targetReason, targetProvenance;
            Structure target=SelectTarget(plan,copy.TargetStructureId,out targetProvenance,out targetReason);
            if(target==null) { foreach(var job in jobs) job.Error=targetReason; return jobs; }
            TargetMeshGeometry mesh;
            try
            {
                var nativeMesh=target.MeshGeometry;
                if(nativeMesh==null || nativeMesh.TriangleIndices.Count==0 || nativeMesh.TriangleIndices.Count>750000)
                    throw new ArgumentException("Target mesh exceeds the bounded geometry budget.");
                // Primitive copies only; native WPF/ESAPI objects remain on the owner thread.
                mesh=new TargetMeshGeometry {
                    Vertices=nativeMesh.Positions.Select(p=>new BeamPoint3D(p.X,p.Y,p.Z)).ToList(),
                    TriangleIndices=nativeMesh.TriangleIndices.ToArray() };
            }
            catch(Exception)
            {
                foreach(var job in jobs) job.Error="Target surface mesh is unavailable or exceeds the 250000-triangle analysis budget.";
                return jobs;
            }
            var beams=plan.Beams.Where(b => !b.IsSetupField && !b.IsImagingTreatmentField).ToList();
            foreach(var job in jobs)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var matches=beams.Where(b=>job.Row.BeamNumber.HasValue && b.BeamNumber==job.Row.BeamNumber.Value).ToList();
                    if(matches.Count!=1) throw new ArgumentException("Beam snapshot mismatch.");
                    var beam=matches[0];
                    var cps=beam.ControlPoints.ToList();
                    if(!MatchesSnapshot(beam,cps,job.Row))
                    { job.Error="Active beam geometry, MU or PAM beam weight changed since the analysis snapshot. Refresh before computing PAM.";job.DiscardNative=true;continue; }
                    if(plan.TreatmentOrientation!=PatientOrientation.HeadFirstSupine ||
                        cps.Any(cp=>!SameAngle(cp.PatientSupportAngle,cps[0].PatientSupportAngle)) ||
                        cps.Any(cp=>!SameAngle(cp.CollimatorAngle,cps[0].CollimatorAngle) ||
                            !SameOptional(cp.TableTopLateralPosition,cps[0].TableTopLateralPosition) ||
                            !SameOptional(cp.TableTopLongitudinalPosition,cps[0].TableTopLongitudinalPosition) ||
                            !SameOptional(cp.TableTopVerticalPosition,cps[0].TableTopVerticalPosition)))
                    { job.Error="Moving target projection requires HFS orientation, a fixed couch angle within each field, fixed collimator and fixed table translation. Other geometries are not inferred.";continue; }
                    var nativeBeam=beam.GetStructureOutlines(target,false);
                    var nativeBev=beam.GetStructureOutlines(target,true);
                    job.NativeOutline=CopyOutline(nativeBeam);
                    double rotation=MeshTargetProjector.NativeCoordinateRotation(CopyOutline(nativeBev),job.NativeOutline,cps[0].CollimatorAngle);
                    var iso=Point(beam.IsocenterPosition);
                    var sourceZero=Point(beam.GetSourceLocation(0));
                    var sourceNinety=Point(beam.GetSourceLocation(90));
                    job.Frames=cps.Select(cp=>MeshTargetProjector.CreateFrame(iso,Point(beam.GetSourceLocation(cp.GantryAngle)),
                        sourceZero,sourceNinety,rotation)).ToList();
                    job.Mesh=mesh;
                }
                catch(Exception)
                { job.Error="Vendor source coordinates, native outline correspondence, or beam snapshot could not be validated. Moving PAM remains unavailable."; }
            }
            return jobs;
        }

        private static ReviewPlanAnalysis CalculateDetachedProjections(ReviewPlanAnalysis copy,
            List<ProjectionJob> jobs,CancellationToken cancellationToken)
        {
            const double resolution=MeshTargetProjector.DefaultResolutionMm;
            const double maximumNativeDifference=0.03; // Numerical consistency guard, not a clinical acceptance limit.
            using(var bounded=CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
            {
                bounded.CancelAfter(TimeSpan.FromSeconds(30));
                foreach(var job in jobs)
                {
                    if(job.Error!=null) { InvalidateProjection(job,job.Error);continue; }
                    try
                    {
                        bounded.Token.ThrowIfCancellationRequested();
                        var first=MeshTargetProjector.Project(job.Mesh,job.Frames[0],resolution,bounded.Token);
                        var native=MeshTargetProjector.RasterizeOutlines(job.NativeOutline,resolution);
                        double difference=MeshTargetProjector.SymmetricDifferenceFraction(first,native);
                        if(difference>maximumNativeDifference)
                        {
                            InvalidateProjection(job,"Projected mesh differs from native beam-start BEV by more than the 3% numerical union-area guard. No moving-gantry PAM is reported.");
                            foreach(var cp in job.Row.ControlPoints) cp.NativeProjectionDifferenceFraction=difference;
                            continue;
                        }
                        var projections=new List<List<ApertureRectangle>> { first };
                        for(int i=1;i<job.Frames.Count;i++)
                        {
                            bounded.Token.ThrowIfCancellationRequested();
                            projections.Add(MeshTargetProjector.Project(job.Mesh,job.Frames[i],resolution,bounded.Token));
                        }
                        // Commit a complete beam only; never subsample or silently keep partial CPs.
                        for(int i=0;i<job.Row.ControlPoints.Count;i++)
                        {
                            var cp=job.Row.ControlPoints[i];
                            cp.TargetProjectionStrips=projections[i];
                            cp.TargetProjectionResolutionMm=resolution;
                            cp.NativeProjectionDifferenceFraction=difference;
                            cp.TargetProjectionReason=null;
                            cp.TargetProjectionProvenance="Divergent ESAPI surface-mesh projection; union of projected triangles at 0.625 mm. Vendor source positions and native true/false outline rotation; beam-start native silhouette difference <=3% numerical guard. HFS/zero couch/fixed collimator only. Research calculation, not clinical commissioning.";
                        }
                    }
                    catch(OperationCanceledException)
                    { InvalidateProjection(job,"Target projection was canceled or exceeded the 30-second computation budget. No partial-beam PAM is reported."); }
                    catch(ArgumentException)
                    { InvalidateProjection(job,"Target projection geometry failed validation or exceeded the bounded raster/triangle-span budget. No partial-beam PAM is reported."); }
                }
            }
            return PlanAnalysisCalculator.Calculate(copy);
        }

        private static void InvalidateProjection(ProjectionJob job,string reason)
        {
            foreach(var cp in job.Row.ControlPoints)
            {
                cp.TargetProjectionStrips.Clear();
                cp.TargetProjectionResolutionMm=null;
                cp.TargetProjectionReason=reason;
                if(job.DiscardNative) cp.TargetOutlines.Clear();
            }
            job.Row.PamReason=reason;
        }

        private static bool MatchesSnapshot(Beam beam,List<ControlPoint> cps,ReviewBeamAnalysis row)
        {
            if (row.PamBeamWeightFactor.HasValue)
            {
                double factor=beam.WeightFactor;
                if (!Finite(factor) || Math.Abs(factor-row.PamBeamWeightFactor.Value)>1e-8) return false;
            }
            var meterset=beam.Meterset;
            var nativeIso=beam.IsocenterPosition;
            if(cps.Count<2 || row.ControlPoints.Count!=cps.Count || !row.MetersetMu.HasValue ||
                meterset.Unit!=DosimeterUnit.MU || !Finite(meterset.Value) || !Finite(row.MetersetMu.Value) ||
                Math.Abs(meterset.Value-row.MetersetMu.Value)>0.001) return false;
            for(int i=0;i<cps.Count;i++)
            {
                var cp=cps[i];var saved=row.ControlPoints[i];var iso=saved.IsocenterMm;
                if(!SameAngle(cp.GantryAngle,saved.GantryAngleDegrees) || !SameAngle(cp.CollimatorAngle,saved.CollimatorAngleDegrees) ||
                    !SameAngle(cp.PatientSupportAngle,saved.PatientSupportAngleDegrees) ||
                    !Finite(cp.MetersetWeight) || !Finite(saved.CumulativeMetersetWeight) ||
                    Math.Abs(cp.MetersetWeight-saved.CumulativeMetersetWeight)>1e-8 || iso==null || iso.Length!=3 || !iso.All(Finite) ||
                    Math.Abs(iso[0]-nativeIso.x)>0.001 || Math.Abs(iso[1]-nativeIso.y)>0.001 ||
                    Math.Abs(iso[2]-nativeIso.z)>0.001) return false;
            }
            return true;
        }

        private static List<List<BeamPoint>> CopyOutline(System.Windows.Point[][] outline)
        {
            if(outline==null) throw new ArgumentException("Native target outline is unavailable.");
            return outline.Where(l=>l!=null && l.Length>=3).Select(l=>l.Select(p=>new BeamPoint(p.X,p.Y)).ToList()).ToList();
        }
        private static BeamPoint3D Point(VVector p) { return new BeamPoint3D(p.x,p.y,p.z); }

        // Deliberately contains no vendor API objects. Safe to access on a worker thread.
        private sealed class ProjectionJob
        {
            internal ReviewBeamAnalysis Row;
            internal TargetMeshGeometry Mesh;
            internal List<BeamProjectionFrame> Frames;
            internal List<List<BeamPoint>> NativeOutline;
            internal string Error;
            internal bool DiscardNative;
        }

        private static void ExtractBeam(Beam beam, Structure target, string targetReason, ReviewBeamAnalysis row)
        {
            row.BeamId = beam.Id;
            row.BeamNumber = beam.BeamNumber;
            row.MachineId = beam.TreatmentUnit == null ? null : beam.TreatmentUnit.Id;
            row.MachineModelName = beam.TreatmentUnit == null ? null : beam.TreatmentUnit.MachineModelName;
            row.MachineModel = beam.TreatmentUnit == null ? null : beam.TreatmentUnit.MachineModel;
            row.GantryDirection = beam.GantryDirection.ToString();
            row.EnergyDisplay = beam.EnergyModeDisplayName;
            row.Technique = beam.Technique == null ? null : beam.Technique.Id;
            row.MlcModel = beam.MLC == null ? null : beam.MLC.Model;
            var meterset = beam.Meterset;
            if (meterset.Unit == DosimeterUnit.MU && Finite(meterset.Value) && meterset.Value >= 0)
                row.MetersetMu = meterset.Value;
            else row.MetersetReason = "Beam meterset is unavailable, nonfinite, or is not in MU.";
            try
            {
                double factor = beam.WeightFactor;
                if (Finite(factor) && factor >= 0) row.PamBeamWeightFactor = factor;
            }
            catch (Exception) { row.PamBeamWeightFactor = null; }
            if (beam.DoseRate > 0) row.NominalDoseRateMuPerMin = beam.DoseRate;
            bool halcyon = ((beam.TreatmentUnit == null ? "" : beam.TreatmentUnit.MachineModelName) + " " +
                (beam.TreatmentUnit == null ? "" : beam.TreatmentUnit.MachineModel) + " " + row.MlcModel)
                .IndexOf("Halcyon",StringComparison.OrdinalIgnoreCase) >= 0;
            if (halcyon) { row.MlcLayerCount = 2; row.HasJaws = false; }
            else if (beam.MLC != null) row.MlcLayerCount = 1;
            var points = beam.ControlPoints.ToList();
            if (points.Count == 0) return;
            bool fixedProjection = points.All(p => SameProjection(points[0],p));
            var initialOutline = new List<List<BeamPoint>>();
            string projectionReason = targetReason;
            if (target != null)
            {
                try
                {
                    // false is the collimator/beam-limiting-device coordinate system.
                    // true would ignore collimator rotation and misalign with leaf/jaw positions.
                    var native = beam.GetStructureOutlines(target, false);
                    if (native != null) initialOutline = native.Where(p => p != null && p.Length >= 3)
                        .Select(p => p.Select(v => new BeamPoint(v.X,v.Y)).ToList()).ToList();
                    projectionReason = initialOutline.Count > 0 ? null : "Native target BEV projection is empty.";
                }
                catch (Exception) { projectionReason = "Native target BEV projection is unavailable."; }
            }
            bool unsupportedAccessory = beam.Applicator != null || beam.Blocks.Any();
            foreach (var cp in points)
            {
                var sample = new ReviewControlPointSample
                {
                    Index = row.ControlPoints.Count,
                    NativeIndex = cp.Index,
                    GantryAngleDegrees = cp.GantryAngle,
                    CollimatorAngleDegrees = cp.CollimatorAngle,
                    PatientSupportAngleDegrees = cp.PatientSupportAngle,
                    CumulativeMetersetWeight = cp.MetersetWeight,
                    IsocenterMm = new[] { beam.IsocenterPosition.x, beam.IsocenterPosition.y, beam.IsocenterPosition.z },
                    NominalDoseRateMuPerMin = row.NominalDoseRateMuPerMin,
                    PlannedDoseRateMuPerMin = null,
                    GeometryReason = halcyon
                        ? "Halcyon requires both native physical MLC layers and explicitly verified staggered boundaries/index mapping. Unknown native layouts remain unavailable."
                        : "Physical MLC leaf boundaries require an exact configured native geometry profile."
                };
                if (fixedProjection || cp.Index == points[0].Index)
                {
                    sample.TargetOutlines = initialOutline;
                    sample.TargetProjectionReason = projectionReason;
                    sample.TargetProjectionProvenance = "Native Beam.GetStructureOutlines(target, false), beam-start geometry in isocenter-plane IEC beam-limiting-device mm.";
                }
                else sample.TargetProjectionReason = "This ESAPI target-outline method supplies only the beam-start projection, not a control-point projection. It is not reused across a moving gantry/collimator/couch.";
                // Jaw-only photon beams can be calculated from the exact API rectangle.
                // MLC beams are never promoted to jaw-only beams when leaf data are missing.
                if (!halcyon && beam.MLC == null && !unsupportedAccessory)
                {
                    var jaws = cp.JawPositions;
                    if (Finite(jaws.X1) && Finite(jaws.Y1) && Finite(jaws.X2) && Finite(jaws.Y2) &&
                        jaws.X2 > jaws.X1 && jaws.Y2 > jaws.Y1)
                    {
                        sample.Aperture = new ApertureGeometry { Jaws = new ApertureRectangle(jaws.X1,jaws.Y1,jaws.X2,jaws.Y2) };
                        sample.GeometryReason = null;
                    }
                }
                if (unsupportedAccessory) sample.GeometryReason = "Electron applicator or blocking accessory geometry is not supported.";
                row.ControlPoints.Add(sample);
            }
        }

        private static Structure SelectTarget(PlanSetup plan, string requestedId, out string provenance, out string reason)
        {
            provenance = reason = null;
            if (plan.StructureSet == null)
            { reason = "The plan has no structure set."; return null; }
            var eligible = plan.StructureSet.Structures.Where(s => !s.IsEmpty && s.HasSegment).ToList();
            string id = PamTargetSelection.Resolve(requestedId, plan.TargetVolumeID,
                eligible.Select(s => s.Id),
                eligible.Where(s => DvhSelectionPolicy.ClassifyTarget(s.Id, s.DicomType) == "PTV").Select(s => s.Id),
                out provenance, out reason);
            return id == null ? null : eligible.Single(s => string.Equals(s.Id,id,StringComparison.OrdinalIgnoreCase));
        }

        private static bool SameProjection(ControlPoint a, ControlPoint b)
        {
            return SameAngle(a.GantryAngle,b.GantryAngle) && SameAngle(a.CollimatorAngle,b.CollimatorAngle) &&
                SameAngle(a.PatientSupportAngle,b.PatientSupportAngle) &&
                SameOptional(a.TableTopLateralPosition,b.TableTopLateralPosition) &&
                SameOptional(a.TableTopLongitudinalPosition,b.TableTopLongitudinalPosition) &&
                SameOptional(a.TableTopVerticalPosition,b.TableTopVerticalPosition);
        }
        private static bool SameAngle(double a, double b)
        { return Finite(a) && Finite(b) && Math.Abs(Math.IEEERemainder(a-b,360)) < 1e-6; }
        private static bool SameOptional(double a, double b)
        { return double.IsNaN(a) && double.IsNaN(b) || Finite(a) && Finite(b) && Math.Abs(a-b)<1e-6; }
        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static double? DoseGy(DoseValue dose)
        {
            if (!Finite(dose.Dose) || dose.Dose <= 0) return null;
            if (dose.Unit == DoseValue.DoseUnit.Gy) return dose.Dose;
            if (dose.Unit == DoseValue.DoseUnit.cGy) return dose.Dose/100;
            return null;
        }
    }
}
