using System.Text.RegularExpressions;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Calculators
{
    class PQMGradientIndex
    {
        public static string GetGradientIndex(StructureSet structureSet, PlanningItemViewModel planningItem, Structure evalStructure, MatchCollection testMatch, Group evalunit)
        {
            // we have Gradient Index pattern
            DVHData dvh = planningItem.PlanningItemObject.GetDVHCumulativeData(evalStructure, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
            if ((dvh.SamplingCoverage < 0.9) || (dvh.Coverage < 0.9))
            {
                return "insufficient dose or sampling coverage";
            }
            Group eval = testMatch[0].Groups["evalpt"];
            Group unit = testMatch[0].Groups["unit"];
            DoseValue prescribedDose;
            double planDoseDouble = 0;
            DoseValue.DoseUnit du = (unit.Value.CompareTo("%") == 0) ? DoseValue.DoseUnit.Percent :
            (unit.Value.CompareTo("Gy") == 0) ? DoseValue.DoseUnit.Gy : DoseValue.DoseUnit.Unknown;
            if (planningItem.PlanningItemObject is PlanSum)
            {
                PlanSum planSum = (PlanSum)planningItem.PlanningItemObject;
                foreach (PlanSetup planSetup in planSum.PlanSetups)
                {
                    planDoseDouble += planSetup.TotalDose.Dose;
                }

            }
            if (planningItem.PlanningItemObject is PlanSetup)
            {
                PlanSetup planSetup = (PlanSetup)planningItem.PlanningItemObject;
                planDoseDouble = planSetup.TotalDose.Dose;
            }
            prescribedDose = new DoseValue(planDoseDouble, DoseValue.DoseUnit.Gy);
            //var body = structureSet.Structures.Where(x => x.Id.Contains("BODY")).First();
            VolumePresentation vpFinal = VolumePresentation.AbsoluteCm3;
            DoseValuePresentation dvpFinal = (evalunit.Value.CompareTo("%") == 0) ? DoseValuePresentation.Relative : DoseValuePresentation.Absolute;
            DoseValue dv = new DoseValue(double.Parse(eval.Value) / 100 * prescribedDose.Dose, DoseValue.DoseUnit.Gy);
            //DoseValue d2 = new DoseValue(2 / 100 * prescribedDose.Dose, DoseValue.DoseUnit.Gy);
            //double d2 = planningItem.PlanningItemObject.GetDoseAtVolume(evalStructure, 2, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
            //double d98 = planningItem.PlanningItemObject.GetDoseAtVolume(evalStructure, 98, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
            double bodyWithPrescribedDoseVolume = planningItem.PlanningItemObject.GetVolumeAtDose(evalStructure, prescribedDose, vpFinal);
            double bodyWithEvalDoseVolume = planningItem.PlanningItemObject.GetVolumeAtDose(evalStructure, dv, vpFinal);
            //double bodyWith2DoseVolume = planningItem.PlanningItemObject.GetVolumeAtDose(evalStructure, d2, vpFinal);
            var gi = bodyWithEvalDoseVolume / bodyWithPrescribedDoseVolume > 100? 100: bodyWithEvalDoseVolume / bodyWithPrescribedDoseVolume;
            if (gi == 100)
                return null;
            return string.Format("{0:0.00}", gi);
        }
    }
}
