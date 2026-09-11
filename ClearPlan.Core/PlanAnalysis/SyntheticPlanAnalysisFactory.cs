using System;
using System.Collections.Generic;

namespace ClearPlan.Core.PlanAnalysis
{
    public static class SyntheticPlanAnalysisFactory
    {
        public static ReviewPlanAnalysis Create(bool dualLayer = false, bool modulated = true)
        {
            var plan = new ReviewPlanAnalysis
            {
                DosePerFractionGy = 2,
                PlanNormalizationPercent = 100,
                TargetStructureId = "PTV_60",
                AvailableTargetStructureIds = new List<string> { "PTV_60" },
                GeometryProvenance = "Deterministic synthetic ellipse and aperture geometry; not a commissioned machine model. Synthetic DVHs are independent analytical examples, not dose calculated from these apertures. " + PlanAnalysisCalculator.SamplingDefinition,
                DoseRateProvenance = "Synthetic illustrative planned MU/min by control point. No measured or delivered dose-rate data."
            };
            for (int beamIndex = 0; beamIndex < 2; beamIndex++)
            {
                var beam = new ReviewBeamAnalysis
                {
                    BeamId = "SYNTHETIC_ARC_" + (beamIndex + 1),
                    BeamNumber = beamIndex + 1,
                    MachineId = dualLayer ? "SYNTHETIC_DUAL_LAYER" : "SYNTHETIC_C_ARM",
                    MlcModel = dualLayer ? "Synthetic staggered dual-layer" : "Synthetic single-layer",
                    EnergyDisplay = dualLayer ? "6 MV FFF (synthetic)" : "6 MV (synthetic)",
                    Technique = "VMAT (synthetic)",
                    MetersetMu = modulated ? 450 : 220,
                    NominalDoseRateMuPerMin = dualLayer ? 800 : 600
                };
                for (int i = 0; i <= 30; i++)
                {
                    double phase = (double)i / 30;
                    var aperture = new ApertureGeometry();
                    if (!dualLayer) aperture.Jaws = new ApertureRectangle(-90,-90,90,90);
                    aperture.Layers.Add(Layer(20,-100,phase,beamIndex,modulated,"Layer 1"));
                    if (dualLayer) aperture.Layers.Add(Layer(21,-105,phase+0.16,beamIndex,modulated,"Layer 2"));
                    var outline = new List<BeamPoint>();
                    for (int p = 0; p < 64; p++)
                    {
                        double angle = p * 2 * Math.PI / 64;
                        outline.Add(new BeamPoint(40*Math.Cos(angle),60*Math.Sin(angle)));
                    }
                    beam.ControlPoints.Add(new ReviewControlPointSample
                    {
                        Index = i,
                        GantryAngleDegrees = (beamIndex == 0 ? 180 + 180*phase : 180*phase) % 360,
                        CollimatorAngleDegrees = beamIndex == 0 ? 15 : 345,
                        PatientSupportAngleDegrees = 0,
                        CumulativeMetersetWeight = phase,
                        NominalDoseRateMuPerMin = beam.NominalDoseRateMuPerMin,
                        PlannedDoseRateMuPerMin = beam.NominalDoseRateMuPerMin.Value *
                            (modulated ? 0.55+0.4*Math.Pow(Math.Sin(phase*Math.PI+beamIndex),2) : 1),
                        Aperture = aperture,
                        TargetOutlines = new List<List<BeamPoint>> { outline },
                        TargetProjectionProvenance = "Synthetic ellipse in isocenter beam-limiting-device plane."
                    });
                }
                plan.Beams.Add(beam);
            }
            return PlanAnalysisCalculator.Calculate(plan);
        }

        private static ApertureLayer Layer(int count, double lower, double phase, int beamIndex,
            bool modulated, string label)
        {
            var layer = new ApertureLayer { Label = label, LeafBoundariesMm = new double[count+1],
                Bank1PositionsMm = new double[count], Bank2PositionsMm = new double[count] };
            for (int i = 0; i <= count; i++) layer.LeafBoundariesMm[i] = lower+i*10;
            for (int i = 0; i < count; i++)
            {
                double width = modulated ? 12+23*(0.5+0.5*Math.Sin(phase*2*Math.PI+i*0.33+beamIndex)) : 80;
                layer.Bank1PositionsMm[i] = -width;
                layer.Bank2PositionsMm[i] = width;
            }
            return layer;
        }
    }
}
