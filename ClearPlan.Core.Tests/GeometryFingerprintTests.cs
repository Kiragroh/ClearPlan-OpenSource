using System;
using System.IO;
using ClearPlan.Core.PlanAnalysis;
using Newtonsoft.Json;

namespace ClearPlan.Core.Tests
{
    internal static class GeometryFingerprintTests
    {
        public static void RunAll()
        {
            StableTypedHashDetectsNativeValueChanges();
            InvalidOrOversizedInputsAreRejected();
            NativeStampsArePrivateButPreservedBySnapshotCopy();
            NativeCaptureAndStaleReportContracts();
        }

        public static void StableTypedHashDetectsNativeValueChanges()
        {
            string original=Fingerprint(10,-50,double.NaN,"Unknown model");
            TestAssert.NotNull(original,"Geometry fingerprint is not implemented.");
            TestAssert.Equal(64,original.Length);
            TestAssert.Equal(original,Fingerprint(10,-50,double.NaN,"Unknown model"));
            TestAssert.False(original==Fingerprint(10.001f,-50,double.NaN,"Unknown model"),"Raw native leaf edits must change the fingerprint.");
            TestAssert.False(original==Fingerprint(10,-49.99,double.NaN,"Unknown model"),"Jaw edits must change the fingerprint.");
            TestAssert.False(original==Fingerprint(10,-50,0,"Unknown model"),"Missing table position is not zero.");
            TestAssert.False(original==Fingerprint(10,-50,double.NaN,"Different model"));
            using(var integer=new GeometryFingerprintWriter()) using(var scalar=new GeometryFingerprintWriter())
            {
                integer.Add(1);scalar.Add(1.0);
                TestAssert.False(integer.Complete()==scalar.Complete(),"Typed values need unambiguous encoding.");
            }
        }

        public static void InvalidOrOversizedInputsAreRejected()
        {
            using(var writer=new GeometryFingerprintWriter())
            {
                TestAssert.Throws<ArgumentException>(()=>writer.Add(double.NaN));
                TestAssert.Throws<ArgumentException>(()=>writer.Add(double.PositiveInfinity));
                TestAssert.Throws<ArgumentException>(()=>writer.AddOptional(double.NegativeInfinity));
                TestAssert.Throws<ArgumentException>(()=>writer.Add(new string('x',257)));
                TestAssert.Throws<ArgumentException>(()=>writer.Add(new float[9,2]));
                TestAssert.Throws<ArgumentException>(()=>writer.Add(new float[,] {{float.NaN}}));
            }
        }

        public static void NativeStampsArePrivateButPreservedBySnapshotCopy()
        {
            var plan=SyntheticPlanAnalysisFactory.Create();
            var planProperty=typeof(ReviewPlanAnalysis).GetProperty("NativePlanFingerprint");
            var beamProperty=typeof(ReviewBeamAnalysis).GetProperty("NativeGeometryFingerprint");
            TestAssert.NotNull(planProperty,"Native plan stamp is missing.");
            TestAssert.NotNull(beamProperty,"Native beam stamp is missing.");
            planProperty.SetValue(plan,"plan-stamp",null);
            beamProperty.SetValue(plan.Beams[0],"beam-stamp",null);
            string json=JsonConvert.SerializeObject(plan);
            TestAssert.False(json.Contains("Fingerprint"));
            TestAssert.False(json.Contains("plan-stamp") || json.Contains("beam-stamp"));
            var copy=PlanAnalysisSnapshot.Copy(plan);
            TestAssert.Equal("plan-stamp",(string)planProperty.GetValue(copy,null));
            TestAssert.Equal("beam-stamp",(string)beamProperty.GetValue(copy.Beams[0],null));
            // Beams with zero CPs must still retain a read-only freshness stamp in detached copies.
            plan.Beams[0].ControlPoints.Clear();
            TestAssert.Equal("beam-stamp",(string)beamProperty.GetValue(PlanAnalysisSnapshot.Copy(plan).Beams[0],null));
        }

        public static void NativeCaptureAndStaleReportContracts()
        {
            string root=Root();
            string path=Path.Combine(root,"ClearPlan.Script","Review","EsapiNativeGeometryFingerprint.cs");
            TestAssert.True(File.Exists(path),"Native ESAPI fingerprint capture is missing.");
            string native=File.ReadAllText(path);
            foreach(string required in new[] {"LeafPositions","JawPositions","IsocenterPosition","MetersetWeight","DoseRate","MLC",
                "GantryAngle","CollimatorAngle","PatientSupportAngle","TableTopLateralPosition","TableTopLongitudinalPosition",
                "TableTopVerticalPosition","PlanNormalizationValue","DosePerFraction","TreatmentOrientation","XDirection","XSize",
                "beam.Applicator","beam.Blocks.Any()","unit.Id"})
                TestAssert.True(native.Contains(required),"Native stamp is missing "+required);
            string builder=File.ReadAllText(Path.Combine(root,"ClearPlan.Script","Review","EsapiPlanAnalysisBuilder.cs"));
            TestAssert.True(builder.Contains("NativeGeometryFingerprint=EsapiNativeGeometryFingerprint.CaptureBeam"));
            TestAssert.True(builder.Contains("NativePlanFingerprint=EsapiNativeGeometryFingerprint.CapturePlan"));
            string bev=File.ReadAllText(Path.Combine(root,"ClearPlan.Script","Review","EsapiBevBuilder.cs")).Replace("\r\n","\n");
            TestAssert.True(bev.Contains("EsapiNativeGeometryFingerprint.CaptureBeam(beam"));
            TestAssert.True(bev.Contains("EsapiNativeGeometryFingerprint.CapturePlan(plan"));
            TestAssert.True(bev.Contains("if(failure.Code==SnapshotReason) throw new InvalidOperationException(SnapshotReason);"),
                "A stale report snapshot must escape as an error, never a normal report with unavailable images.");
            TestAssert.True(bev.Contains("var completed=await RunDetached(job,bounded.Token);\n                    bounded.Token.ThrowIfCancellationRequested();"),
                "Cancellation must precede native reads on a queued worker continuation.");
            TestAssert.True(bev.Contains("var completed=await RunStartsDetached(copy,jobs,bounded.Token);\n                    bounded.Token.ThrowIfCancellationRequested();"));
        }

        private static string Fingerprint(float leaf,double jaw,double table,string model)
        {
            using(var writer=new GeometryFingerprintWriter())
            {
                writer.Add(model);writer.Add(200.0);writer.Add(600);writer.Add(jaw);writer.AddOptional(table);
                writer.Add(new float[,] {{-leaf,-11},{leaf,11}});
                return writer.Complete();
            }
        }

        private static string Root()
        {
            var dir=new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while(dir!=null && !File.Exists(Path.Combine(dir.FullName,"ClearPlan.sln"))) dir=dir.Parent;
            if(dir==null) dir=new DirectoryInfo(Directory.GetCurrentDirectory());
            TestAssert.True(File.Exists(Path.Combine(dir.FullName,"ClearPlan.sln")),"Repository root is required.");
            return dir.FullName;
        }
    }
}
