using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows;
using System.Windows.Controls;

namespace ClearPlan.Calculators
{
    class PQMConformationNumber
    {
        public static string GetConformationNumber(StructureSet structureSet, PlanningItemViewModel planningItem, Structure evalStructure, MatchCollection testMatch, Group evalunit)
        {
            DVHData dvh = planningItem.PlanningItemObject.GetDVHCumulativeData(evalStructure, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
            DoseValue prescribedDose;
            double planDoseDouble = 0;
            if ((dvh.SamplingCoverage < 0.9) || (dvh.Coverage < 0.9))
            {
                return "insufficient dose or sampling coverage";
            }
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
            Group eval = testMatch[0].Groups["evalpt"];
            Group unit = testMatch[0].Groups["unit"];
            DoseValue.DoseUnit du = (unit.Value.CompareTo("%") == 0) ? DoseValue.DoseUnit.Percent :
                (unit.Value.CompareTo("Gy") == 0) ? DoseValue.DoseUnit.Gy : DoseValue.DoseUnit.Unknown;
            var body = structureSet.Structures.Where(x => x.DicomType.ToUpper().Equals("EXTERNAL") || x.Id.ToUpper().Equals("BODY") || x.Id.ToUpper().Contains("OUTER CONTOUR") || x.Id.ToUpper().Equals("KÖRPER")).FirstOrDefault();
            if (body == null)
            {
                string msg2 = string.Format("No Body-Structure found in '{0}'.", structureSet.Id);
                //MessageBox.Show(msg2, "CI-Calculation-Problem", MessageBoxButton.OK, MessageBoxImage.Information); 
            }
            //VolumePresentation vpFinal = (evalunit.Value.CompareTo("%") == 0) ? VolumePresentation.Relative : VolumePresentation.AbsoluteCm3;
            VolumePresentation vpFinal = VolumePresentation.AbsoluteCm3;
            DoseValuePresentation dvpFinal = (evalunit.Value.CompareTo("%") == 0) ? DoseValuePresentation.Relative : DoseValuePresentation.Absolute;
            DoseValue dv = new DoseValue(double.Parse(eval.Value) / 100 * prescribedDose.Dose, DoseValue.DoseUnit.Gy);
            double bodyWithPrescribedDoseVolume = planningItem.PlanningItemObject.GetVolumeAtDose(body, dv, vpFinal);
            double targetWithPrescribedDoseVolume = planningItem.PlanningItemObject.GetVolumeAtDose(evalStructure, dv, vpFinal);
            double targetVolume = evalStructure.Volume;
            var cn = 1/((targetWithPrescribedDoseVolume / targetVolume) * (targetWithPrescribedDoseVolume / bodyWithPrescribedDoseVolume));
            return string.Format("{0:0.00}", cn);
        }
    }
}
