using System.Text.RegularExpressions;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Calculators
{
    class PQMVolumeAtDose
    {
        public static string GetVolumeAtDose(StructureSet structureSet, PlanningItemViewModel planningItem, Structure evalStructure, MatchCollection testMatch, Group evalunit)
        {
            //check for sufficient sampling and dose coverage
            DVHData dvh = planningItem.PlanningItemObject.GetDVHCumulativeData(evalStructure, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
            //MessageBox.Show(evalStructure.Id + "- Eval unit: " + evalunit.Value.ToString() + "Achieved unit: " + dvAchieved.UnitAsString + " - Sampling coverage: " + dvh.SamplingCoverage.ToString() + " Coverage: " + dvh.Coverage.ToString());
            if ((dvh.SamplingCoverage< 0.9) || (dvh.Coverage< 0.9))
            {
                return "insufficient dose or sampling coverage";
            }
            Group eval = testMatch[0].Groups["evalpt"];
            Group unit = testMatch[0].Groups["unit"];
            DoseValue.DoseUnit du = (unit.Value.CompareTo("%") == 0) ? DoseValue.DoseUnit.Percent :
                    (unit.Value.CompareTo("Gy") == 0) ? DoseValue.DoseUnit.Gy : DoseValue.DoseUnit.Unknown;
            VolumePresentation vp = (unit.Value.CompareTo("%") == 0) ? VolumePresentation.Relative : VolumePresentation.AbsoluteCm3;
            DoseValue dv;
            dv = new DoseValue(0, DoseValue.DoseUnit.Gy);

            if (planningItem.PlanningItemObject is PlanSetup && evalStructure.Id == planningItem.PlanningItemTargetId)
            {
                dv = new DoseValue(double.Parse(eval.Value) * planningItem.PlanningItemTreatmentPercentage, du);
            }
            if (planningItem.PlanningItemObject is PlanSetup && evalStructure.Id != planningItem.PlanningItemTargetId)
            {
                dv = new DoseValue(double.Parse(eval.Value), du);
            }
            if (planningItem.PlanningItemObject is PlanSum && du != DoseValue.DoseUnit.Percent)
            {
                /*PlanSum planSum = (PlanSum)planningItem.PlanningItemObject;
                double planDoseDouble = 0;
                foreach (PlanSetup planSetup in planSum.PlanSetups)
                {
                    double planSetupRxDose = planSetup.TotalDose.Dose;
                    planDoseDouble += planSetupRxDose;
                }
                dv += new DoseValue(planDoseDouble, DoseValue.DoseUnit.Gy);*/
                dv = new DoseValue(double.Parse(eval.Value), du);
            }
            if (planningItem.PlanningItemObject is PlanSum && du == DoseValue.DoseUnit.Percent)
            {
                /*PlanSum planSum = (PlanSum)planningItem.PlanningItemObject;
                double planDoseDouble = 0;
                foreach (PlanSetup planSetup in planSum.PlanSetups)
                {
                    double planSetupRxDose = planSetup.TotalDose.Dose;
                    planDoseDouble += planSetupRxDose;
                }
                dv += new DoseValue(planDoseDouble, DoseValue.DoseUnit.Gy);*/
                //dv = new DoseValue(double.Parse(eval.Value)*55, DoseValue.DoseUnit.Gy);
                return "Unable to calculate";

            }


            double volume = double.Parse(eval.Value);
            VolumePresentation vpFinal = (evalunit.Value.CompareTo("%") == 0) ? VolumePresentation.Relative : VolumePresentation.AbsoluteCm3;
            
            //DoseValuePresentation dvpFinal = (evalunit.Value.CompareTo("%") == 0) ? DoseValuePresentation.Relative : DoseValuePresentation.Absolute;
            

            double volumeAchieved = planningItem.PlanningItemObject.GetVolumeAtDose(evalStructure, dv, vpFinal);
            return string.Format("{0:0.00} {1}", volumeAchieved, evalunit.Value);   // todo: better formatting based on VolumePresentation

        }
    }
}
