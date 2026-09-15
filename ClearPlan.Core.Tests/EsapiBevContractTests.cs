using System;
using System.IO;

namespace ClearPlan.Core.Tests
{
    internal static class EsapiBevContractTests
    {
        public static void RunAll()
        {
            OwnerThreadCaptureAndDetachedWorkerAreSeparated();
            NativeFrameAndFullSnapshotAreRequired();
            CtSamplingIsUniformBoundedAndReadOnly();
            ReportBeamStartsShareOneDetachedCtCapture();
            TreatmentBeamScopeMatchesPamAnalysis();
            StaticNativeIndicesRemainDistinctInTheSnapshot();
        }

        public static void OwnerThreadCaptureAndDetachedWorkerAreSeparated()
        {
            string code=Source();
            Require(code,"ApartmentState.STA");
            Require(code,"Dispatcher.FromThread(Thread.CurrentThread)");
            Require(code,"SynchronizationContext.Current is DispatcherSynchronizationContext");
            Require(code,"owner.VerifyAccess();");
            Require(code,"await Dispatcher.Yield(DispatcherPriority.Background);");
            Require(code,"var completed=await RunDetached(job,bounded.Token);");
            Require(code,"var completed=await RunStartsDetached(copy,jobs,bounded.Token);");
            Require(code,"bounded.Token.ThrowIfCancellationRequested();\n                    return completed;");
            Require(code,"private static Task<BeamEyeViewImage> RunDetached(DetachedJob job,CancellationToken token)");
            string worker=Between(code,"private static Task<BeamEyeViewImage> RunDetached", "private static async Task<CtVolume> CaptureVolumeAsync");
            Require(worker,"Task.Run(() => CalculateDetached(job,token),token)");
            Require(worker,"CtDrrProjector.Project(job.Volume,job.Frame");
            foreach(string forbidden in new[] { "PlanSetup", "StructureSet", "GetVoxels", "VoxelToDisplayValue", "VVector", "GetStructureOutlines" })
                TestAssert.False(worker.Contains(forbidden),"Worker boundary must not reference ESAPI: "+forbidden);
            TestAssert.False(code.Contains("ConfigureAwait(false)"),"Owner capture continuation must not lose its dispatcher context.");
        }

        public static void NativeFrameAndFullSnapshotAreRequired()
        {
            string code=Source();
            Require(code,"GetStructureOutlines(calibration,false)");
            Require(code,"GetStructureOutlines(calibration,true)");
            Require(code,"MeshTargetProjector.NativeCoordinateRotation");
            Require(code,"MeshTargetProjector.CreateFrame");
            Require(code,"GetSourceLocation(0)");
            Require(code,"GetSourceLocation(90)");
            Require(code,"PatientOrientation.HeadFirstSupine");
            Require(code,"SameAngle(cp.PatientSupportAngle,cps[0].PatientSupportAngle)");
            Require(code,"TableTopLateralPosition");
            Require(code,"TableTopLongitudinalPosition");
            Require(code,"TableTopVerticalPosition");
            Require(code,"beams.Count!=analysis.Beams.Count");
            Require(code,"cp.Index!=(saved.NativeIndex ?? saved.Index)");
            Require(code,"meterset.Unit!=DosimeterUnit.MU");
            Require(code,"saved.IsocenterMm");
            Require(code,"saved.CumulativeMetersetWeight");
            TestAssert.True(Occurrences(code,"MatchesPlanSnapshot(plan,copy)")>=2,"Recheck active beam state after yielding during CT capture.");
            TestAssert.False(code.Contains("SyntheticDrrFactory"),"A live ESAPI DRR must not use a synthetic coordinate fallback.");
        }

        public static void CtSamplingIsUniformBoundedAndReadOnly()
        {
            string code=Source();
            Require(code,"DisplayUnit"); Require(code,"\"HU\"");
            Require(code,"image.GetVoxels(z*strideZ,buffer)");
            Require(code,"buffer[x*strideX,y*strideY]");
            Require(code,"image.VoxelToDisplayValue(raw)");
            Require(code,"(int)Math.Ceiling((size-1)/255.0)");
            Require(code,"(nx-1)/strideX+1");
            Require(code,"SpacingX=image.XRes*strideX");
            Require(code,"XAxis=Point(image.XDirection)");
            Require(code,"TimeSpan.FromSeconds(45)");
            Require(code,"CtDrrProjector.MaximumVoxelCount");
            Require(code,"SourceStatus=\"unavailable\"");
            Require(code,"nearest-neighbor integer-stride");
            foreach(string forbidden in new[] { "BeginModifications", "SaveModifications", "SetVoxels", "DicomFile", "File.Write", "error.Message", "exception.Message" })
                TestAssert.False(code.Contains(forbidden),"Live CT workflow must remain read-only and PHI-safe: "+forbidden);
        }

        public static void ReportBeamStartsShareOneDetachedCtCapture()
        {
            string code=Source();
            Require(code,"public async Task<ReviewPlanAnalysis> BuildStartsAsync");
            string batch=Between(code,"public async Task<ReviewPlanAnalysis> BuildStartsAsync", "private static Dispatcher OwnerDispatcher");
            TestAssert.Equal(1,Occurrences(batch,"CaptureVolumeAsync("),"Report must capture one shared CT volume, not one per beam.");
            Require(batch,"job.Volume=volume");
            Require(batch,"PlanAnalysisSnapshot.Copy(analysis)");
            Require(code,"CalculateStartsDetached(copy,jobs,token)");
            Require(code,"MarkStartsUnavailable(copy");
        }

        private static void TreatmentBeamScopeMatchesPamAnalysis()
        {
            string code=Source();
            string capture=Between(code,"private static DetachedJob CaptureJob", "private static Task<BeamEyeViewImage> RunDetached");
            Require(capture,"!b.IsSetupField && !b.IsImagingTreatmentField");
            string snapshot=Between(code,"private static bool MatchesPlanSnapshot", "private static bool MatchesBeamSnapshot");
            Require(snapshot,"!b.IsSetupField && !b.IsImagingTreatmentField");
        }

        private static void StaticNativeIndicesRemainDistinctInTheSnapshot()
        {
            var source=new ClearPlan.Core.PlanAnalysis.ReviewPlanAnalysis();
            var beam=new ClearPlan.Core.PlanAnalysis.ReviewBeamAnalysis { BeamNumber=1 };
            beam.ControlPoints.Add(new ClearPlan.Core.PlanAnalysis.ReviewControlPointSample { Index=0,NativeIndex=-1 });
            beam.ControlPoints.Add(new ClearPlan.Core.PlanAnalysis.ReviewControlPointSample { Index=1,NativeIndex=-1 });
            source.Beams.Add(beam);
            var copy=ClearPlan.Core.PlanAnalysis.PlanAnalysisSnapshot.Copy(source);
            TestAssert.Equal(0,copy.Beams[0].ControlPoints[0].Index,"Static field start has an unambiguous internal ordinal.");
            TestAssert.Equal(1,copy.Beams[0].ControlPoints[1].Index,"Static field end is a distinct endpoint.");
            TestAssert.Equal(-1,copy.Beams[0].ControlPoints[0].NativeIndex.Value,"Preserve the vendor sentinel for freshness validation.");
            TestAssert.Equal(-1,copy.Beams[0].ControlPoints[1].NativeIndex.Value,"Do not falsify the second vendor index.");
            var root=new DirectoryInfo(Directory.GetCurrentDirectory());
            while(root!=null && !File.Exists(Path.Combine(root.FullName,"ClearPlan.sln"))) root=root.Parent;
            TestAssert.NotNull(root,"Repository root required.");
            string builder=File.ReadAllText(Path.Combine(root.FullName,"ClearPlan.Script","Review","EsapiPlanAnalysisBuilder.cs"));
            Require(builder,"Index = row.ControlPoints.Count");
            Require(builder,"NativeIndex = cp.Index");
            string collision=File.ReadAllText(Path.Combine(root.FullName,"ClearPlan.Script","Review","EsapiCollisionBuilder.cs"));
            Require(collision,"for (int pointIndex = 0; pointIndex < points.Count; pointIndex++)");
            Require(collision,"var cp = points[pointIndex]");
            Require(collision,"ControlPointIndex = pointIndex");
            TestAssert.False(collision.Contains("ControlPointIndex = cp.Index"),
                "Static ESAPI sentinel indices must not invalidate or merge collision poses.");
        }

        private static string Source()
        {
            var root=new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while(root!=null && !File.Exists(Path.Combine(root.FullName,"ClearPlan.sln"))) root=root.Parent;
            if(root==null)
            {
                root=new DirectoryInfo(Directory.GetCurrentDirectory());
                while(root!=null && !File.Exists(Path.Combine(root.FullName,"ClearPlan.sln"))) root=root.Parent;
            }
            TestAssert.NotNull(root,"Repository root is required for ESAPI source contract checks.");
            string path=Path.Combine(root.FullName,"ClearPlan.Script","Review","EsapiBevBuilder.cs");
            TestAssert.True(File.Exists(path),"ESAPI DRR capture is not implemented.");
            return File.ReadAllText(path).Replace("\r\n","\n");
        }

        private static string Between(string text,string first,string after)
        {
            int start=text.IndexOf(first,StringComparison.Ordinal),end=text.IndexOf(after,StringComparison.Ordinal);
            TestAssert.True(start>=0 && end>start,"Detached worker section is missing.");
            return text.Substring(start,end-start);
        }
        private static int Occurrences(string text,string value)
        { return (text.Length-text.Replace(value,"").Length)/value.Length; }
        private static void Require(string text,string value)
        { TestAssert.True(text.Contains(value),"Missing ESAPI safety contract: "+value); }
    }
}
