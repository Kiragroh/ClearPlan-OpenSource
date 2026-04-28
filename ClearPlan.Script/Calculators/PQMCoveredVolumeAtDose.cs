using System.Text.RegularExpressions;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Calculators
{
    class PQMCoveredVolumeAtDose
    {
        public static string GetCoveredVolumeAtDose(StructureSet structureSet, PlanningItemViewModel planningItem, Structure evalStructure, MatchCollection testMatch, Group evalunit, string variation)
        {
            //check for sufficient sampling and dose coverage
            DVHData dvh = planningItem.PlanningItemObject.GetDVHCumulativeData(evalStructure, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
            if ((dvh.SamplingCoverage < 0.9) || (dvh.Coverage < 0.9))
            {
               return "insufficient dose or sampling coverage";
            }
            Group eval = testMatch[0].Groups["evalpt"];
            Group unit = testMatch[0].Groups["unit"];
            DoseValue.DoseUnit du = (unit.Value.CompareTo("%") == 0) ? DoseValue.DoseUnit.Percent :
                    (unit.Value.CompareTo("Gy") == 0) ? DoseValue.DoseUnit.Gy : DoseValue.DoseUnit.Unknown;
            VolumePresentation vp = (unit.Value.CompareTo("%") == 0) ? VolumePresentation.Relative : VolumePresentation.AbsoluteCm3;
            DoseValue dv = new DoseValue(double.Parse(eval.Value), du);
            double volume = double.Parse(eval.Value);            
            VolumePresentation vpFinal = (evalunit.Value.CompareTo("%") == 0) ? VolumePresentation.Relative : VolumePresentation.AbsoluteCm3;
            DoseValuePresentation dvpFinal = (evalunit.Value.CompareTo("%") == 0) ? DoseValuePresentation.Relative : DoseValuePresentation.Absolute;
            double volumeAchieved = planningItem.PlanningItemObject.GetVolumeAtDose(evalStructure, dv, vpFinal);
            double organVolume = evalStructure.Volume;
            double coveredVolume = organVolume - volumeAchieved;
            if (organVolume < (double.TryParse(variation, out double value) ? value : 0.0)) return "Volume too small";
           return string.Format("{0:0.00} {1}", coveredVolume, evalunit.Value);   // todo: better formatting based on VolumePresentation
        }
    }
}
