using System;
using System.Collections.Generic;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Tests
{
    internal static class PlanAnalysisTests
    {
        public static int Main()
        {
            CalculationContractExists();
            RunCalculations();
            SyntheticExamplesAreCalculatedAndVendorFree();
            MeshTargetProjectionTests.AnalyticFramesAndPerspective();
            MeshTargetProjectionTests.TriangleUnionNotXorAndNoConvexHull();
            MeshTargetProjectionTests.PlanPamUsesUnionRows();
            MeshTargetProjectionTests.NativeRotationUsesCoordinatesNotAngleGuess();
            MeshTargetProjectionTests.CancellationAndInvalidFramesAreExplicit();
            DetachedCopyPreservesPrivateGeometryWithoutSharing();
            EsapiAdapterIsReadOnlyAndExplicitAboutMissingData();
            return 0;
        }

        public static void DetachedCopyPreservesPrivateGeometryWithoutSharing()
        {
            var source=SyntheticPlanAnalysisFactory.Create(true,true);
            var copy=PlanAnalysisSnapshot.Copy(source);
            TestAssert.False(ReferenceEquals(source,copy),"Background work must not mutate the active snapshot.");
            TestAssert.True(copy.Beams[0].ControlPoints[0].Aperture.Layers.Count==2);
            double before=source.Beams[0].ControlPoints[0].Aperture.Layers[0].Bank1PositionsMm[0];
            copy.Beams[0].ControlPoints[0].Aperture.Layers[0].Bank1PositionsMm[0]+=1;
            Near(before,source.Beams[0].ControlPoints[0].Aperture.Layers[0].Bank1PositionsMm[0]);
            string json=Newtonsoft.Json.JsonConvert.SerializeObject(source);
            TestAssert.False(json.Contains("Bank1PositionsMm"));
            TestAssert.False(json.Contains("TargetOutlines"));
            TestAssert.False(json.Contains("IsocenterMm"));
        }

        public static void SyntheticExamplesAreCalculatedAndVendorFree()
        {
            var open = SyntheticPlanAnalysisFactory.Create(false, false);
            var dual = SyntheticPlanAnalysisFactory.Create(true, true);
            Near(0, open.Pam);
            TestAssert.True(dual.Pam > 0 && dual.Pam < 1);
            TestAssert.True(dual.Beams.All(b => b.MlcLayerCount == 2 && b.HasJaws == false));
            TestAssert.True(dual.Beams.SelectMany(b => b.ControlPoints).All(c => c.Aperture.EffectiveOpenings.Count > 0));
            Near(dual.Pam.Value, SyntheticPlanAnalysisFactory.Create(true, true).Pam);
        }

        public static void EsapiAdapterIsReadOnlyAndExplicitAboutMissingData()
        {
            var dir = new System.IO.DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (dir != null && !System.IO.File.Exists(System.IO.Path.Combine(dir.FullName, "ClearPlan.sln"))) dir = dir.Parent;
            TestAssert.NotNull(dir, "Repository root required for adapter contract test.");
            string path = System.IO.Path.Combine(dir.FullName, "ClearPlan.Script", "Review", "EsapiPlanAnalysisBuilder.cs");
            TestAssert.True(System.IO.File.Exists(path), "Missing clinical plan analysis adapter.");
            string code = System.IO.File.ReadAllText(path);
            TestAssert.False(code.Contains("BeginModifications("));
            TestAssert.False(code.Contains("SaveModifications("));
            TestAssert.False(code.Contains("ApplyParameters("));
            TestAssert.True(code.Contains("DosimeterUnit.MU"));
            TestAssert.True(code.Contains("GetStructureOutlines(target, false)"));
            TestAssert.True(code.Contains("TargetVolumeID"));
            TestAssert.True(code.Contains("PlannedDoseRateMuPerMin = null"));
            TestAssert.True(code.Contains("EnrichTargetProjectionsAsync"), "Missing explicit async projection workflow.");
            TestAssert.True(code.Contains("CaptureProjectionJobs"));
            TestAssert.True(code.Contains("string.IsNullOrWhiteSpace(copy.TargetStructureId)"),
                "Projection enrichment must not silently resolve a new target after explicit selection failed.");
            TestAssert.True(code.Contains("PamTargetSelection.Resolve"), "Native extraction must use the tested deterministic target policy.");
            TestAssert.True(code.Contains("s.DicomType"), "Automatic PTV eligibility must come from the DICOM type, not a name guess.");
            TestAssert.True(code.Contains("PamWeightingMode = \"BeamWeightFactor\""));
            TestAssert.True(code.Contains("beam.WeightFactor"));
            TestAssert.True(code.Contains("Math.Abs(factor-row.PamBeamWeightFactor.Value)"),
                "Projection enrichment must reject a changed PAM beam weight, even when MU and geometry are unchanged.");
            string fingerprint = System.IO.File.ReadAllText(System.IO.Path.Combine(dir.FullName,
                "ClearPlan.Script", "Review", "EsapiNativeGeometryFingerprint.cs"));
            TestAssert.True(fingerprint.Contains("stamp.Add(beam.WeightFactor)"),
                "Native report freshness must include the PAM beam-weight factor.");
            TestAssert.True(fingerprint.Contains("stamp.Add(beam.IsImagingTreatmentField"),
                "Native beam freshness must include imaging-treatment scope changes.");
            TestAssert.True(code.Split(new[] { "!b.IsSetupField && !b.IsImagingTreatmentField" },StringSplitOptions.None).Length >= 3,
                "Both beam extraction and projection capture must exclude setup and imaging-treatment fields like PlanCheck.");
        }

        public static void RunCalculations()
        {
            TargetSelectionIsDeterministicAndFailsClosed();
            NativePamUsesPlanCheckBeamWeights();
            FullHalcyonLayersRemainJawlessAndUseBothBanks();
            EndpointWeightsMatchPlanCheckDmsw();
            SingleLayerAndJaws();
            DualLayerIntersectionUsesStaggeredBoundaries();
            PamUsesTargetAreaAndMuWeights();
            ConcaveAndHoleTargets();
            InvalidAndMissingGeometryStaysUnavailable();
            InvalidMetersetCannotProducePartialPlanMetric();
            ZeroMuIntervalDoesNotBiasPam();
            NonUnitFinalMetersetAndMuPerGy();
            FixedJawlessBoundaryClipsWithoutCreatingJaws();
            NoNonFiniteMetricsAndNoStaleGeometryReason();
        }

        private static void TargetSelectionIsDeterministicAndFailsClosed()
        {
            string provenance, reason;
            TestAssert.Equal("PTV_ONLY", ResolveTarget(null,null,new[] {"PTV_ONLY","BODY"},new[] {"PTV_ONLY"},out provenance,out reason));
            TestAssert.True(provenance.Contains("unique"));
            TestAssert.True(reason==null);
            TestAssert.Equal("CTV", ResolveTarget(" ctv ","PTV_A",new[] {"PTV_A","PTV_B","CTV"},new[] {"PTV_A","PTV_B"},out provenance,out reason));
            TestAssert.True(provenance.Contains("Explicit"));
            TestAssert.Equal(null, ResolveTarget(null,"ptv_b",new[] {"PTV_A","PTV_B"},new[] {"PTV_A","PTV_B"},out provenance,out reason));
            TestAssert.True(reason.Contains("lowest"));
            TestAssert.Equal("PTV_ONLY", ResolveTarget(null,"STALE",new[] {"PTV_ONLY"},new[] {"PTV_ONLY"},out provenance,out reason));
            TestAssert.True(provenance.Contains("invalid"));
            TestAssert.True(ResolveTarget("MISSING",null,new[] {"PTV_ONLY"},new[] {"PTV_ONLY"},out provenance,out reason)==null);
            TestAssert.True(reason.Contains("explicit"));
            TestAssert.True(ResolveTarget(null,null,new[] {"PTV_A","PTV_B"},new[] {"PTV_B","PTV_A"},out provenance,out reason)==null);
            TestAssert.True(reason.Contains("Multiple"));
            TestAssert.True(ResolveTarget(null,null,new[] {"PTV_like_name","BODY"},new string[0],out provenance,out reason)==null);
            TestAssert.True(ResolveTarget(null,null,new[] {"DUP","dup"},new[] {"DUP"},out provenance,out reason)==null);
        }

        private static string ResolveTarget(string requested,string planTarget,string[] eligible,string[] ptvs,
            out string provenance,out string reason)
        {
            return PamTargetSelection.Resolve(requested,planTarget,eligible,ptvs,out provenance,out reason);
        }

        private static void NativePamUsesPlanCheckBeamWeights()
        {
            var first=Plan(Layer(new[] {-10d,10},new[] {-20d},new[] {0d}));
            var second=Plan(Layer(new[] {-10d,10},new[] {-20d},new[] {20d}));
            first.Beams[0].MetersetMu=100;second.Beams[0].MetersetMu=300;
            first.Beams.Add(second.Beams[0]);
            first.PamWeightingMode="BeamWeightFactor";
            first.Beams[0].PamBeamWeightFactor=3;first.Beams[1].PamBeamWeightFactor=1;
            first.TargetSelectionProvenance="Explicit target selection; plan data unchanged.";
            PlanAnalysisCalculator.Calculate(first);
            Near(0.375,first.Pam);
            Near(7,first.MeanApertureAreaCm2);
            Near(400,first.TotalMetersetMu);
            TestAssert.True(first.PamReason.Contains("Beam.WeightFactor"));
            TestAssert.True(first.PamReason.Contains(first.TargetSelectionProvenance));
            var copy=PlanAnalysisSnapshot.Copy(first);
            TestAssert.Equal(first.TargetSelectionProvenance,copy.TargetSelectionProvenance);
            Near(3,copy.Beams[0].PamBeamWeightFactor);
            TestAssert.Equal("BeamWeightFactor",copy.PamWeightingMode);
            first.Beams[1].PamBeamWeightFactor=null;
            PlanAnalysisCalculator.Calculate(first);
            TestAssert.False(first.Pam.HasValue,"Missing native weight must not silently fall back to MU.");
            first.Beams[0].PamBeamWeightFactor=0;first.Beams[1].PamBeamWeightFactor=0;
            PlanAnalysisCalculator.Calculate(first);
            TestAssert.False(first.Pam.HasValue,"All-zero beam weights cannot produce a plan PAM.");
            first.Beams[0].PamBeamWeightFactor=double.NaN;first.Beams[1].PamBeamWeightFactor=1;
            PlanAnalysisCalculator.Calculate(first);
            TestAssert.False(first.Pam.HasValue);
            first.Beams[0].PamBeamWeightFactor=-1;
            PlanAnalysisCalculator.Calculate(first);
            TestAssert.False(first.Pam.HasValue);
            first.Beams[0].PamBeamWeightFactor=0;
            first.Beams[0].ControlPoints[1].TargetOutlines.Clear();
            PlanAnalysisCalculator.Calculate(first);
            Near(0,first.Pam);
            first.Beams[0].PamBeamWeightFactor=1;
            PlanAnalysisCalculator.Calculate(first);
            TestAssert.False(first.Pam.HasValue,"A positive-weight beam with incomplete projections cannot be silently omitted.");
        }

        private static void FullHalcyonLayersRemainJawlessAndUseBothBanks()
        {
            var distal=Layer(Enumerable.Range(0,29).Select(i=>-140d+10*i).ToArray(),
                Enumerable.Repeat(-20d,28).ToArray(),Enumerable.Repeat(20d,28).ToArray());
            var proximal=Layer(Enumerable.Range(0,30).Select(i=>-145d+10*i).ToArray(),
                Enumerable.Repeat(-10d,29).ToArray(),Enumerable.Repeat(0d,29).ToArray());
            var plan=Plan(distal,proximal);
            foreach(var cp in plan.Beams[0].ControlPoints)
                cp.TargetOutlines=new List<List<BeamPoint>> { Outline(-10,-140,10,140) };
            PlanAnalysisCalculator.Calculate(plan);
            Near(28,plan.MeanApertureAreaCm2);Near(0.5,plan.Pam);
            TestAssert.Equal(2,plan.Beams[0].MlcLayerCount);
            TestAssert.Equal(false,plan.Beams[0].HasJaws);
            TestAssert.True(plan.Beams[0].ControlPoints.All(cp=>cp.Aperture.Jaws==null));
            foreach(var cp in plan.Beams[0].ControlPoints)
            {
                cp.Aperture.Layers[0].Bank1PositionsMm=Enumerable.Repeat(-5d,28).ToArray();
                cp.Aperture.Layers[0].Bank2PositionsMm=Enumerable.Repeat(5d,28).ToArray();
            }
            PlanAnalysisCalculator.Calculate(plan);
            Near(14,plan.MeanApertureAreaCm2);Near(0.75,plan.Pam);
        }

        private static void EndpointWeightsMatchPlanCheckDmsw()
        {
            var plan=Plan(Layer(new[] {-10d,10},new[] {-20d},new[] {20d}));
            var template=PlanAnalysisSnapshot.Copy(plan).Beams[0].ControlPoints;
            plan.Beams[0].ControlPoints.AddRange(template);
            double[] weights={0,0.2,0.8,1};
            for(int i=0;i<4;i++) plan.Beams[0].ControlPoints[i].CumulativeMetersetWeight=weights[i];
            PlanAnalysisCalculator.Calculate(plan);
            double[] expected={10,40,40,10};
            for(int i=0;i<4;i++) Near(expected[i],plan.Beams[0].ControlPoints[i].MetricMetersetWeightMu);
        }

        private static void NoNonFiniteMetricsAndNoStaleGeometryReason()
        {
            var plan=Plan(Layer(new[] {-10d,10},new[] {-20d},new[] {20d}));
            PlanAnalysisCalculator.Calculate(plan);
            plan.Beams[0].ControlPoints[0].Aperture.Jaws=new ApertureRectangle(double.NaN,-10,10,10);
            PlanAnalysisCalculator.Calculate(plan);
            TestAssert.True(plan.Beams[0].GeometryReason.IndexOf("invalid",StringComparison.OrdinalIgnoreCase)>=0,
                "A previous successful geometry provenance must not hide the new invalid geometry reason.");
            plan=Plan(Layer(new[] {-1e200,1e200},new[] {-1e200},new[] {1e200}));
            PlanAnalysisCalculator.Calculate(plan);
            TestAssert.False(plan.MeanApertureAreaCm2.HasValue,"Overflowed aperture area must not become a nonfinite scalar.");
            plan=Plan(Layer(new[] {-10d,10},new[] {-20d},new[] {20d}));
            plan.Beams[0].MetersetMu=double.MaxValue;
            plan.Beams.Add(Plan(Layer(new[] {-10d,10},new[] {-20d},new[] {20d})).Beams[0]);
            plan.Beams[1].MetersetMu=double.MaxValue;
            PlanAnalysisCalculator.Calculate(plan);
            TestAssert.False(plan.TotalMetersetMu.HasValue,"Overflowed total MU must remain unavailable.");
        }

        private static void FixedJawlessBoundaryClipsWithoutCreatingJaws()
        {
            var plan=Plan(Layer(new[] {-10d,10},new[] {-20d},new[] {20d}));
            foreach(var cp in plan.Beams[0].ControlPoints)
                cp.Aperture.FixedBoundingBox=new ApertureRectangle(-10,-10,10,10);
            PlanAnalysisCalculator.Calculate(plan);
            Near(4,plan.MeanApertureAreaCm2);
            TestAssert.Equal(false,plan.Beams[0].HasJaws);
            TestAssert.True(PlanAnalysisSnapshot.Copy(plan).Beams[0].ControlPoints[0].Aperture.FixedBoundingBox!=null);
        }

        private static void SingleLayerAndJaws()
        {
            var plan = Plan(Layer(new[] { -10d, 0, 10 }, new[] { -20d, -10 }, new[] { 20d, 10 }));
            foreach (var cp in plan.Beams[0].ControlPoints)
                cp.Aperture.Jaws = new ApertureRectangle(-15, -5, 15, 5);
            PlanAnalysisCalculator.Calculate(plan);
            Near(2.5, plan.MeanApertureAreaCm2);
            Near(0, plan.Pam);
            Near(1, plan.SmallApertureFraction);
        }

        private static void DualLayerIntersectionUsesStaggeredBoundaries()
        {
            var plan = Plan(Layer(new[] { -10d, 0, 10 }, new[] { -10d, -10 }, new[] { 10d, 10 }),
                Layer(new[] { -15d, -5, 5, 15 }, new[] { 0d, -5, 0 }, new[] { 5d, 5, 5 }));
            PlanAnalysisCalculator.Calculate(plan);
            Near(1.5, plan.MeanApertureAreaCm2);
            Near(0.5, plan.Pam);
            TestAssert.Equal(2, plan.Beams[0].MlcLayerCount);
            TestAssert.Equal(false, plan.Beams[0].HasJaws);
            TestAssert.Equal(4, plan.Beams[0].ControlPoints[0].Aperture.EffectiveOpenings.Count);
        }

        private static void PamUsesTargetAreaAndMuWeights()
        {
            var first = Plan(Layer(new[] { -10d, 10 }, new[] { -20d }, new[] { 0d }));
            first.Beams[0].MetersetMu = 100;
            var second = Plan(Layer(new[] { -10d, 10 }, new[] { -20d }, new[] { 20d }));
            second.Beams[0].MetersetMu = 300;
            first.Beams.Add(second.Beams[0]);
            PlanAnalysisCalculator.Calculate(first);
            Near(0.125, first.Pam);
            Near(7, first.MeanApertureAreaCm2);
            Near(400, first.TotalMetersetMu);
        }

        private static void ConcaveAndHoleTargets()
        {
            var plan = Plan(Layer(new[] { -20d, 20 }, new[] { -20d }, new[] { 0d }));
            var outer = Outline(-10, -10, 10, 10);
            var hole = Outline(-5, -5, 5, 5);
            foreach (var cp in plan.Beams[0].ControlPoints)
                cp.TargetOutlines = new List<List<BeamPoint>> { outer, hole };
            PlanAnalysisCalculator.Calculate(plan);
            Near(3, plan.Beams[0].ControlPoints[0].TargetAreaCm2);
            Near(0.5, plan.Pam);
            foreach (var cp in plan.Beams[0].ControlPoints)
                cp.TargetOutlines = new List<List<BeamPoint>> { new List<BeamPoint> {
                    new BeamPoint(-10,-10),new BeamPoint(10,-10),new BeamPoint(10,0),
                    new BeamPoint(0,0),new BeamPoint(0,10),new BeamPoint(-10,10) } };
            PlanAnalysisCalculator.Calculate(plan);
            Near(3, plan.Beams[0].ControlPoints[0].TargetAreaCm2);
            Near(1.0/3, plan.Pam);
        }

        private static void InvalidAndMissingGeometryStaysUnavailable()
        {
            var plan = Plan(Layer(new[] { -10d, 10 }, new[] { -20d }, new[] { 20d }));
            plan.Beams[0].ControlPoints[1].Aperture.Layers[0].LeafBoundariesMm = new[] { 10d, -10 };
            PlanAnalysisCalculator.Calculate(plan);
            TestAssert.False(plan.Pam.HasValue);
            TestAssert.False(plan.MeanApertureAreaCm2.HasValue);
            TestAssert.True(!string.IsNullOrEmpty(plan.PamReason));
            plan = Plan(Layer(new[] { -10d, 10 }, new[] { -20d }, new[] { 20d }));
            plan.Beams[0].ControlPoints[1].TargetOutlines.Clear();
            PlanAnalysisCalculator.Calculate(plan);
            TestAssert.False(plan.Pam.HasValue);
            Near(8, plan.MeanApertureAreaCm2);
            plan.Beams[0].ControlPoints[0].Aperture = new ApertureGeometry();
            PlanAnalysisCalculator.Calculate(plan);
            TestAssert.False(plan.MeanApertureAreaCm2.HasValue);
        }

        private static void InvalidMetersetCannotProducePartialPlanMetric()
        {
            var plan = Plan(Layer(new[] { -10d, 10 }, new[] { -20d }, new[] { 20d }));
            var bad = Plan(Layer(new[] { -10d, 10 }, new[] { -20d }, new[] { 20d })).Beams[0];
            bad.MetersetMu = null;
            plan.Beams.Add(bad);
            PlanAnalysisCalculator.Calculate(plan);
            TestAssert.False(plan.TotalMetersetMu.HasValue);
            TestAssert.False(plan.Pam.HasValue);
            plan.Beams.Remove(bad);
            plan.Beams[0].ControlPoints[0].CumulativeMetersetWeight = 0.2;
            PlanAnalysisCalculator.Calculate(plan);
            TestAssert.False(plan.Pam.HasValue);
            TestAssert.False(plan.Beams[0].ControlPoints[0].IncrementalMetersetMu.HasValue);
        }

        private static void ZeroMuIntervalDoesNotBiasPam()
        {
            var plan = Plan(Layer(new[] { -10d, 10 }, new[] { -20d }, new[] { 20d }));
            var cp = new ReviewControlPointSample { Index = 2, CumulativeMetersetWeight = 1,
                Aperture = null };
            plan.Beams[0].ControlPoints.Add(cp);
            PlanAnalysisCalculator.Calculate(plan);
            Near(0, plan.Pam);
            Near(0, cp.MetricMetersetWeightMu);
        }

        private static void NonUnitFinalMetersetAndMuPerGy()
        {
            var plan = Plan(Layer(new[] { -10d, 10 }, new[] { -20d }, new[] { 20d }));
            plan.Beams[0].ControlPoints[1].CumulativeMetersetWeight = 100;
            PlanAnalysisCalculator.Calculate(plan);
            Near(100, plan.Beams[0].ControlPoints[1].IncrementalMetersetMu);
            Near(50, plan.Beams[0].ControlPoints[0].MetricMetersetWeightMu);
            Near(50, plan.MuPerGy);
            plan.DosePerFractionGy = 0;
            PlanAnalysisCalculator.Calculate(plan);
            TestAssert.False(plan.MuPerGy.HasValue);
            TestAssert.True(plan.Beams[0].ControlPoints.All(p => !p.PlannedDoseRateMuPerMin.HasValue));
        }

        private static ReviewPlanAnalysis Plan(params ApertureLayer[] layers)
        {
            var beam = new ReviewBeamAnalysis { BeamId = "SYNTHETIC", MetersetMu = 100 };
            for (int i = 0; i < 2; i++) beam.ControlPoints.Add(new ReviewControlPointSample {
                Index = i, CumulativeMetersetWeight = i,
                Aperture = new ApertureGeometry { Layers = layers.Select(l => Layer(
                    (double[])l.LeafBoundariesMm.Clone(), (double[])l.Bank1PositionsMm.Clone(),
                    (double[])l.Bank2PositionsMm.Clone())).ToList() },
                TargetOutlines = new List<List<BeamPoint>> { Outline(-10, -5, 10, 5) } });
            return new ReviewPlanAnalysis { DosePerFractionGy = 2, TargetStructureId = "SYNTHETIC_TARGET",
                Beams = new List<ReviewBeamAnalysis> { beam } };
        }

        private static ApertureLayer Layer(double[] bounds, double[] left, double[] right)
        { return new ApertureLayer { LeafBoundariesMm = bounds, Bank1PositionsMm = left, Bank2PositionsMm = right }; }
        private static List<BeamPoint> Outline(double x1, double y1, double x2, double y2)
        { return new List<BeamPoint> { new BeamPoint(x1,y1),new BeamPoint(x2,y1),new BeamPoint(x2,y2),new BeamPoint(x1,y2) }; }
        private static void Near(double expected, double? actual)
        { TestAssert.True(actual.HasValue && Math.Abs(expected-actual.Value)<1e-8,
            "Expected " + expected + "; actual " + (actual.HasValue ? actual.Value.ToString() : "unavailable")); }

        public static void CalculationContractExists()
        {
            TestAssert.NotNull(typeof(ClearPlan.Core.PlanAnalysis.ReviewPlanAnalysis).Assembly.GetType(
                "ClearPlan.Core.PlanAnalysis.PlanAnalysisCalculator"),
                "Missing vendor-free plan analysis calculation contract.");
        }
    }
}
