using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using ClearPlan.Core.PlanAnalysis;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Review
{
    /// <summary>
    /// Read-only ESAPI planning-CT capture. Call on the ESAPI owner's STA dispatcher while the
    /// host holds the patient context open. Only detached numeric DTOs enter worker tasks.
    /// </summary>
    public sealed class EsapiBevBuilder
    {
        private const string OwnerReason="ESAPI_BEV_OWNER_THREAD: The ESAPI STA dispatcher context is unavailable.";
        private const string ContextReason="ESAPI_BEV_CONTEXT: A single external plan with planning CT is required.";
        private const string SnapshotReason="ESAPI_BEV_SNAPSHOT_CHANGED: Native beam, plan normalization or captured image metadata differs from the review snapshot, or its freshness cannot be verified. Refresh before requesting a DRR or report.";
        private const string FrameReason="ESAPI_BEV_NATIVE_FRAME: Native source coordinates and paired structure outlines could not establish a matching beam-coordinate frame. No replacement frame was assumed.";
        private const string MotionReason="ESAPI_BEV_UNSUPPORTED_MOTION: Native frame calibration requires HFS, fixed couch angle within each field, fixed collimator and fixed table translation.";
        private const string CtReason="ESAPI_BEV_CT_GRID: Planning CT HU conversion, directions, dimensions or sampling exceeded the supported bounds.";
        private const string ReadReason="ESAPI_BEV_READ_FAILED: Read-only planning CT or beam capture failed. No substitute anatomy was generated.";
        private const string ComputeReason="ESAPI_BEV_PROJECTION: Detached CT projection failed validation or exceeded its numerical sampling budget.";
        private const string TimeoutReason="ESAPI_BEV_TIMEOUT: The cooperative 45-second operation budget expired; no partial image series is published.";

        public async Task<BeamEyeViewImage> BuildAsync(PlanSetup plan,ReviewPlanAnalysis analysis,
            int beamNumber,int cpIndex,CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var owner=OwnerDispatcher();
            if(owner==null) return Unavailable(OwnerReason,null,cpIndex);
            ReviewControlPointSample selected=null;
            using(var bounded=CancellationTokenSource.CreateLinkedTokenSource(token))
            {
                bounded.CancelAfter(TimeSpan.FromSeconds(45));
                string phase="snapshot";
                try
                {
                    owner.VerifyAccess();
                    RequireContext(plan,analysis);
                    var copy=PlanAnalysisSnapshot.Copy(analysis);
                    if(!MatchesPlanSnapshot(plan,copy)) throw new CaptureFailure(SnapshotReason);
                    var rows=copy.Beams.Where(b=>b.BeamNumber==beamNumber).ToList();
                    if(rows.Count!=1) throw new CaptureFailure(SnapshotReason);
                    var points=rows[0].ControlPoints.Where(cp=>cp.Index==cpIndex).ToList();
                    if(points.Count!=1) throw new CaptureFailure(SnapshotReason);
                    selected=points[0];
                    var job=CaptureJob(plan,copy,rows[0],selected,bounded.Token);
                    phase="ct-capture";
                    job.Volume=await CaptureVolumeAsync(plan.StructureSet.Image,owner,bounded.Token);
                    bounded.Token.ThrowIfCancellationRequested();
                    owner.VerifyAccess();
                    if(!MatchesPlanSnapshot(plan,copy)) throw new CaptureFailure(SnapshotReason);
                    bounded.Token.ThrowIfCancellationRequested();
                    phase="projection";
                    var completed=await RunDetached(job,bounded.Token);
                    bounded.Token.ThrowIfCancellationRequested();
                    owner.VerifyAccess();
                    if(!MatchesPlanSnapshot(plan,copy)) throw new CaptureFailure(SnapshotReason);
                    bounded.Token.ThrowIfCancellationRequested();
                    return completed;
                }
                catch(OperationCanceledException)
                {
                    token.ThrowIfCancellationRequested();
                    return Unavailable(TimeoutReason,selected,cpIndex);
                }
                catch(CaptureFailure failure) { return Unavailable(failure.Code,selected,cpIndex); }
                catch(ArgumentException) { return Unavailable(ComputeReason,selected,cpIndex); }
                catch(Exception error) { return Unavailable(ReadFailure(phase,error),selected,cpIndex); }
            }
        }

        /// <summary>Detached report copy: one planning-CT capture shared by all treatment beam-start DRRs.</summary>
        public async Task<ReviewPlanAnalysis> BuildStartsAsync(PlanSetup plan,ReviewPlanAnalysis analysis,CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if(analysis==null) throw new ArgumentNullException("analysis");
            var copy=PlanAnalysisSnapshot.Copy(analysis);
            var owner=OwnerDispatcher();
            if(owner==null) return MarkStartsUnavailable(copy,OwnerReason);
            using(var bounded=CancellationTokenSource.CreateLinkedTokenSource(token))
            {
                bounded.CancelAfter(TimeSpan.FromSeconds(45));
                string phase="snapshot";
                try
                {
                    owner.VerifyAccess();
                    if(!MatchesPlanSnapshot(plan,copy)) throw new CaptureFailure(SnapshotReason);
                    RequireContext(plan,copy);
                    var jobs=new List<DetachedJob>();
                    foreach(var row in copy.Beams)
                    {
                        bounded.Token.ThrowIfCancellationRequested();
                        var start=row.ControlPoints[0];
                        try { jobs.Add(CaptureJob(plan,copy,row,start,bounded.Token)); }
                        catch(OperationCanceledException) { throw; }
                        catch(CaptureFailure failure) { start.BevImage=Unavailable(failure.Code,start,start.Index); }
                        catch(Exception) { start.BevImage=Unavailable(FrameReason,start,start.Index); }
                    }
                    if(jobs.Count==0)
                    {
                        bounded.Token.ThrowIfCancellationRequested();
                        if(!MatchesPlanSnapshot(plan,copy)) throw new CaptureFailure(SnapshotReason);
                        return copy;
                    }
                    phase="ct-capture";
                    var volume=await CaptureVolumeAsync(plan.StructureSet.Image,owner,bounded.Token);
                    bounded.Token.ThrowIfCancellationRequested();
                    owner.VerifyAccess();
                    if(!MatchesPlanSnapshot(plan,copy)) throw new CaptureFailure(SnapshotReason);
                    foreach(var job in jobs) job.Volume=volume;
                    bounded.Token.ThrowIfCancellationRequested();
                    phase="projection";
                    var completed=await RunStartsDetached(copy,jobs,bounded.Token);
                    bounded.Token.ThrowIfCancellationRequested();
                    owner.VerifyAccess();
                    if(!MatchesPlanSnapshot(plan,copy)) throw new CaptureFailure(SnapshotReason);
                    bounded.Token.ThrowIfCancellationRequested();
                    return completed;
                }
                catch(OperationCanceledException)
                {
                    token.ThrowIfCancellationRequested();
                    return MarkStartsUnavailable(copy,TimeoutReason);
                }
                catch(CaptureFailure failure)
                {
                    if(failure.Code==SnapshotReason) throw new InvalidOperationException(SnapshotReason);
                    return MarkStartsUnavailable(copy,failure.Code);
                }
                catch(Exception error) { return MarkStartsUnavailable(copy,ReadFailure(phase,error)); }
            }
        }

        private static Dispatcher OwnerDispatcher()
        {
            var owner=Dispatcher.FromThread(Thread.CurrentThread);
            if(Thread.CurrentThread.GetApartmentState()!=ApartmentState.STA || owner==null ||
                !(SynchronizationContext.Current is DispatcherSynchronizationContext) ||
                (System.Windows.Application.Current!=null && !System.Windows.Application.Current.Dispatcher.CheckAccess())) return null;
            owner.VerifyAccess();
            return owner;
        }

        private static void RequireContext(PlanSetup plan,ReviewPlanAnalysis analysis)
        {
            if(!(plan is ExternalPlanSetup) || plan.StructureSet==null || plan.StructureSet.Image==null ||
                analysis==null || analysis.Beams==null || analysis.Beams.Count<1 || analysis.Beams.Count>128)
                throw new CaptureFailure(ContextReason);
        }

        private static DetachedJob CaptureJob(PlanSetup plan,ReviewPlanAnalysis analysis,ReviewBeamAnalysis row,
            ReviewControlPointSample selected,CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var beam=plan.Beams.Single(b=>!b.IsSetupField && !b.IsImagingTreatmentField && b.BeamNumber==row.BeamNumber.Value);
            var cps=beam.ControlPoints.ToList();
            if(plan.TreatmentOrientation!=PatientOrientation.HeadFirstSupine ||
                cps.Any(cp=>!SameAngle(cp.PatientSupportAngle,cps[0].PatientSupportAngle)) ||
                cps.Any(cp=>!SameAngle(cp.CollimatorAngle,cps[0].CollimatorAngle) ||
                    !SameOptional(cp.TableTopLateralPosition,cps[0].TableTopLateralPosition) ||
                    !SameOptional(cp.TableTopLongitudinalPosition,cps[0].TableTopLongitudinalPosition) ||
                    !SameOptional(cp.TableTopVerticalPosition,cps[0].TableTopVerticalPosition)))
                throw new CaptureFailure(MotionReason);
            var calibration=SelectCalibrationStructure(plan,analysis.TargetStructureId);
            if(calibration==null) throw new CaptureFailure(FrameReason);
            try
            {
                // These paired native outlines establish the actual BLD rotation; no IEC sign is guessed.
                var beamOutline=CopyOutline(beam.GetStructureOutlines(calibration,false));
                var bevOutline=CopyOutline(beam.GetStructureOutlines(calibration,true));
                double rotation=MeshTargetProjector.NativeCoordinateRotation(bevOutline,beamOutline,cps[0].CollimatorAngle);
                var sourceZero=Point(beam.GetSourceLocation(0));
                var sourceNinety=Point(beam.GetSourceLocation(90));
                var frame=MeshTargetProjector.CreateFrame(Point(beam.IsocenterPosition),
                    Point(beam.GetSourceLocation(selected.GantryAngleDegrees)),sourceZero,sourceNinety,rotation);
                return new DetachedJob { Frame=frame,BeamNumber=row.BeamNumber.Value,ControlPointIndex=selected.Index,
                    Gantry=selected.GantryAngleDegrees,Collimator=selected.CollimatorAngleDegrees,
                    BldToDisplayRotation=rotation*180/Math.PI,
                    Couch=selected.PatientSupportAngleDegrees,Extent=BeamExtent(row,cps) };
            }
            catch(OperationCanceledException) { throw; }
            catch(Exception) { throw new CaptureFailure(FrameReason); }
        }

        // The worker launch is isolated in a static helper so its closure cannot retain a live plan or image.
        private static Task<BeamEyeViewImage> RunDetached(DetachedJob job,CancellationToken token)
        { return Task.Run(() => CalculateDetached(job,token),token); }

        private static BeamEyeViewImage CalculateDetached(DetachedJob job,CancellationToken token)
        {
            var image=CtDrrProjector.Project(job.Volume,job.Frame,224,job.Extent,token);
            image.ControlPointIndex=job.ControlPointIndex; image.GantryAngleDegrees=job.Gantry;
            image.CollimatorAngleDegrees=job.Collimator; image.PatientSupportAngleDegrees=job.Couch;
            image.BldToDisplayRotationDegrees=job.BldToDisplayRotation;
            image.Synthetic=false;
            image.ProjectionDescription="Read-only ESAPI planning CT, not a stored beam reference image. " +
                "Native source locations and paired beam/BEV structure outlines define the beam-coordinate frame. " +
                "HFS, zero couch and fixed collimator/table translation. " +
                "Uniform nearest-neighbor integer-stride overview; the final native border shorter than one stride may be omitted. " +
                "Downsampled edge-voxel support approximates the original CT boundary; field of view is capped at +/-300 mm. " +
                string.Format(CultureInfo.InvariantCulture,"Overview voxel spacing: {0:0.###}/{1:0.###}/{2:0.###} mm. ",
                    job.Volume.SpacingX,job.Volume.SpacingY,job.Volume.SpacingZ) + CtDrrProjector.MethodDescription;
            return image;
        }

        private static Task<ReviewPlanAnalysis> RunStartsDetached(ReviewPlanAnalysis copy,List<DetachedJob> jobs,CancellationToken token)
        { return Task.Run(() => CalculateStartsDetached(copy,jobs,token),token); }

        private static ReviewPlanAnalysis CalculateStartsDetached(ReviewPlanAnalysis copy,List<DetachedJob> jobs,CancellationToken token)
        {
            foreach(var job in jobs)
            {
                token.ThrowIfCancellationRequested();
                var cp=copy.Beams.Single(b=>b.BeamNumber==job.BeamNumber).ControlPoints.Single(p=>p.Index==job.ControlPointIndex);
                try { cp.BevImage=CalculateDetached(job,token); }
                catch(ArgumentException) { cp.BevImage=Unavailable(ComputeReason,cp,cp.Index); }
            }
            token.ThrowIfCancellationRequested();
            return copy;
        }

        private static async Task<CtVolume> CaptureVolumeAsync(VMS.TPS.Common.Model.API.Image image,Dispatcher owner,CancellationToken token)
        {
            owner.VerifyAccess();
            if(image==null || !string.Equals(image.DisplayUnit,"HU",StringComparison.OrdinalIgnoreCase)) throw new CaptureFailure(CtReason);
            int nx=image.XSize,ny=image.YSize,nz=image.ZSize;
            if(nx<2 || ny<2 || nz<2 || nx>2048 || ny>2048 || nz>4096) throw new CaptureFailure(CtReason);
            int strideX=Stride(nx),strideY=Stride(ny),strideZ=Stride(nz);
            var volume=new CtVolume { SizeX=(nx-1)/strideX+1,SizeY=(ny-1)/strideY+1,SizeZ=(nz-1)/strideZ+1,
                SpacingX=image.XRes*strideX,SpacingY=image.YRes*strideY,SpacingZ=image.ZRes*strideZ,
                Origin=Point(image.Origin),XAxis=Point(image.XDirection),YAxis=Point(image.YDirection),ZAxis=Point(image.ZDirection) };
            int count=checked(volume.SizeX*volume.SizeY*volume.SizeZ);
            if(count>CtDrrProjector.MaximumVoxelCount || !ValidSpacing(volume.SpacingX) ||
                !ValidSpacing(volume.SpacingY) || !ValidSpacing(volume.SpacingZ)) throw new CaptureFailure(CtReason);
            volume.HounsfieldUnits=new float[count];
            var buffer=new int[nx,ny];
            var conversion=new Dictionary<int,float>();
            for(int z=0;z<volume.SizeZ;z++)
            {
                owner.VerifyAccess();
                token.ThrowIfCancellationRequested();
                image.GetVoxels(z*strideZ,buffer);
                for(int y=0;y<volume.SizeY;y++)
                {
                    token.ThrowIfCancellationRequested();
                    for(int x=0;x<volume.SizeX;x++)
                    {
                        int raw=buffer[x*strideX,y*strideY];
                        float hu;
                        if(!conversion.TryGetValue(raw,out hu))
                        {
                            double display=image.VoxelToDisplayValue(raw);
                            if(!Finite(display) || Math.Abs(display)>float.MaxValue || conversion.Count>=65536) throw new CaptureFailure(CtReason);
                            hu=(float)display; conversion.Add(raw,hu);
                        }
                        volume.HounsfieldUnits[(z*volume.SizeY+y)*volume.SizeX+x]=hu;
                    }
                }
                if((z+1)%8==0)
                {
                    await Dispatcher.Yield(DispatcherPriority.Background);
                    token.ThrowIfCancellationRequested();
                    owner.VerifyAccess();
                }
            }
            return volume;
        }

        private static int Stride(int size) { return Math.Max(1,(int)Math.Ceiling((size-1)/255.0)); }
        // Static phase and exception type only: vendor messages can contain patient context.
        private static string ReadFailure(string phase,Exception error)
        { return ReadReason+" Stage: "+phase+"; type: "+error.GetType().Name+
            "; member: "+(error.TargetSite==null ? "unknown" : error.TargetSite.Name)+"."; }
        private static bool ValidSpacing(double spacing) { return Finite(spacing) && spacing>=0.05 && spacing<=100; }

        private static bool MatchesPlanSnapshot(PlanSetup plan,ReviewPlanAnalysis analysis)
        {
            try
            {
                // A missing stamp is not evidence of equality. Native read failures reject the report as stale/unverifiable.
                if(string.IsNullOrWhiteSpace(analysis.NativePlanFingerprint) ||
                    !string.Equals(analysis.NativePlanFingerprint,EsapiNativeGeometryFingerprint.CapturePlan(plan),StringComparison.Ordinal) ||
                    !analysis.PlanNormalizationPercent.HasValue || !Finite(analysis.PlanNormalizationPercent.Value) ||
                    !Finite(plan.PlanNormalizationValue) || Math.Abs(plan.PlanNormalizationValue-analysis.PlanNormalizationPercent.Value)>1e-8)
                    return false;
                var fraction=plan.DosePerFraction;
                double? fractionGy=fraction.Unit==DoseValue.DoseUnit.Gy ? (double?)fraction.Dose :
                    fraction.Unit==DoseValue.DoseUnit.cGy ? fraction.Dose/100 : (double?)null;
                if(!fractionGy.HasValue || !analysis.DosePerFractionGy.HasValue || !Finite(fractionGy.Value) ||
                    !Finite(analysis.DosePerFractionGy.Value) || Math.Abs(fractionGy.Value-analysis.DosePerFractionGy.Value)>1e-8) return false;
                var beams=plan.Beams.Where(b=>!b.IsSetupField && !b.IsImagingTreatmentField).ToList();
                if(beams.Count!=analysis.Beams.Count || analysis.Beams.Any(b=>b==null || !b.BeamNumber.HasValue) ||
                    analysis.Beams.Select(b=>b.BeamNumber.Value).Distinct().Count()!=analysis.Beams.Count) return false;
                foreach(var row in analysis.Beams)
                {
                    var matches=beams.Where(b=>b.BeamNumber==row.BeamNumber.Value).ToList();
                    if(matches.Count!=1 || !MatchesBeamSnapshot(matches[0],row)) return false;
                }
                return true;
            }
            catch(Exception) { return false; }
        }

        private static bool MatchesBeamSnapshot(Beam beam,ReviewBeamAnalysis row)
        {
            if(string.IsNullOrWhiteSpace(row.NativeGeometryFingerprint) ||
                !string.Equals(row.NativeGeometryFingerprint,EsapiNativeGeometryFingerprint.CaptureBeam(beam),StringComparison.Ordinal)) return false;
            var cps=beam.ControlPoints.ToList();
            var meterset=beam.Meterset;
            var iso=beam.IsocenterPosition;
            if(row.ControlPoints==null || cps.Count<2 || cps.Count>4096 || cps.Count!=row.ControlPoints.Count ||
                !string.Equals(beam.Id,row.BeamId,StringComparison.Ordinal) || !row.MetersetMu.HasValue ||
                meterset.Unit!=DosimeterUnit.MU || !Finite(meterset.Value) || !Finite(row.MetersetMu.Value) ||
                Math.Abs(meterset.Value-row.MetersetMu.Value)>0.001 || !Finite(iso.x) || !Finite(iso.y) || !Finite(iso.z)) return false;
            for(int i=0;i<cps.Count;i++)
            {
                var cp=cps[i]; var saved=row.ControlPoints[i];
                if(saved==null) return false;
                var position=saved.IsocenterMm;
                if(cp.Index!=(saved.NativeIndex ?? saved.Index) || !SameAngle(cp.GantryAngle,saved.GantryAngleDegrees) ||
                    !SameAngle(cp.CollimatorAngle,saved.CollimatorAngleDegrees) || !SameAngle(cp.PatientSupportAngle,saved.PatientSupportAngleDegrees) ||
                    !Finite(cp.MetersetWeight) || !Finite(saved.CumulativeMetersetWeight) ||
                    Math.Abs(cp.MetersetWeight-saved.CumulativeMetersetWeight)>1e-8 ||
                    position==null || position.Length!=3 || !position.All(Finite) ||
                    Math.Abs(position[0]-iso.x)>0.001 || Math.Abs(position[1]-iso.y)>0.001 || Math.Abs(position[2]-iso.z)>0.001) return false;
            }
            return true;
        }

        private static Structure SelectCalibrationStructure(PlanSetup plan,string requestedTarget)
        {
            var structures=plan.StructureSet.Structures.ToList();
            if(!string.IsNullOrWhiteSpace(requestedTarget))
            {
                var targets=structures.Where(s=>string.Equals(s.Id,requestedTarget,StringComparison.OrdinalIgnoreCase) && !s.IsEmpty && s.HasSegment).ToList();
                if(targets.Count==1) return targets[0];
            }
            var external=structures.Where(s=>string.Equals(s.DicomType,"EXTERNAL",StringComparison.OrdinalIgnoreCase) && !s.IsEmpty && s.HasSegment).ToList();
            return external.Count==1 ? external[0] : null;
        }

        private static List<List<BeamPoint>> CopyOutline(System.Windows.Point[][] outline)
        {
            if(outline==null || outline.Length==0 || outline.Length>512 || outline.Any(l=>l==null || l.Length<3) ||
                outline.Sum(l=>(long)l.Length)>100000) throw new CaptureFailure(FrameReason);
            return outline.Select(l=>l.Select(p=>new BeamPoint(p.X,p.Y)).ToList()).ToList();
        }

        // One common extent for the whole beam, including closed leaves and every complete physical boundary.
        private static double BeamExtent(ReviewBeamAnalysis row,List<ControlPoint> native)
        {
            double extent=100;
            foreach(var cp in row.ControlPoints)
            {
                if(cp.Aperture==null) continue;
                extent=RectangleExtent(extent,cp.Aperture.Jaws);
                extent=RectangleExtent(extent,cp.Aperture.FixedBoundingBox);
                foreach(var layer in cp.Aperture.Layers ?? new List<ApertureLayer>())
                {
                    extent=ArrayExtent(extent,layer.LeafBoundariesMm);
                    extent=ArrayExtent(extent,layer.Bank1PositionsMm);
                    extent=ArrayExtent(extent,layer.Bank2PositionsMm);
                }
            }
            foreach(var cp in native)
            {
                var jaws=cp.JawPositions;
                if(Finite(jaws.X1) && Finite(jaws.X2) && Finite(jaws.Y1) && Finite(jaws.Y2) && jaws.X2>jaws.X1 && jaws.Y2>jaws.Y1)
                    extent=RectangleExtent(extent,new ApertureRectangle(jaws.X1,jaws.Y1,jaws.X2,jaws.Y2));
            }
            return Math.Max(100,Math.Min(300,extent*1.2));
        }

        private static double RectangleExtent(double extent,ApertureRectangle rectangle)
        { return rectangle==null ? extent : ArrayExtent(extent,new[] { rectangle.X1,rectangle.Y1,rectangle.X2,rectangle.Y2 }); }
        private static double ArrayExtent(double extent,double[] values)
        {
            if(values==null) return extent;
            foreach(double value in values) { if(!Finite(value)) throw new CaptureFailure(FrameReason);extent=Math.Max(extent,Math.Abs(value)); }
            return extent;
        }

        private static BeamEyeViewImage Unavailable(string reason,ReviewControlPointSample cp,int index)
        {
            return new BeamEyeViewImage { SourceStatus="unavailable",UnavailableReason=reason,Synthetic=false,ControlPointIndex=index,
                GantryAngleDegrees=cp==null ? 0 : cp.GantryAngleDegrees,CollimatorAngleDegrees=cp==null ? 0 : cp.CollimatorAngleDegrees,
                PatientSupportAngleDegrees=cp==null ? 0 : cp.PatientSupportAngleDegrees,ProjectionDescription=CtDrrProjector.MethodDescription };
        }

        private static ReviewPlanAnalysis MarkStartsUnavailable(ReviewPlanAnalysis copy,string reason)
        {
            foreach(var row in copy.Beams) if(row.ControlPoints.Count>0)
            { var cp=row.ControlPoints[0]; cp.BevImage=Unavailable(reason,cp,cp.Index); }
            return copy;
        }

        private static BeamPoint3D Point(VVector point) { return new BeamPoint3D(point.x,point.y,point.z); }
        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static bool SameAngle(double a,double b)
        { return Finite(a) && Finite(b) && Math.Abs(Math.IEEERemainder(a-b,360))<1e-6; }
        private static bool SameOptional(double a,double b)
        { return double.IsNaN(a) && double.IsNaN(b) || Finite(a) && Finite(b) && Math.Abs(a-b)<1e-6; }

        // Numeric DTOs only; never add a live vendor image, beam, plan or structure to this type.
        private sealed class DetachedJob
        {
            internal CtVolume Volume;
            internal BeamProjectionFrame Frame;
            internal int BeamNumber,ControlPointIndex;
            internal double Gantry,Collimator,Couch,Extent,BldToDisplayRotation;
        }

        private sealed class CaptureFailure : Exception
        {
            internal CaptureFailure(string code) { Code=code; }
            internal string Code { get; private set; }
        }
    }
}
