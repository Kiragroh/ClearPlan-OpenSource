using System.Text.RegularExpressions;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Calculators
{
    class PQMDoseAtVolume
    {
        public static string GetDoseAtVolume(StructureSet structureSet, PlanningItemViewModel planningItem, Structure evalStructure, MatchCollection testMatch, Group evalunit)
        {
            //check for sufficient dose and sampling coverage
            DVHData dvh = planningItem.PlanningItemObject.GetDVHCumulativeData(evalStructure, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
            if (dvh != null && planningItem.PlanningItemType.ToString() != "PlanSum")
            {
                if ((dvh.SamplingCoverage < 0.9) || (dvh.Coverage < 0.9))
                    return "insufficient dose or sampling coverage";
                Group eval = testMatch[0].Groups["evalpt"];
                Group unit = testMatch[0].Groups["unit"];
                DoseValue.DoseUnit du = (unit.Value.CompareTo("%") == 0) ? DoseValue.DoseUnit.Percent :
                        (unit.Value.CompareTo("Gy") == 0) ? DoseValue.DoseUnit.Gy : DoseValue.DoseUnit.Unknown;
                VolumePresentation vp = (unit.Value.CompareTo("%") == 0) ? VolumePresentation.Relative : VolumePresentation.AbsoluteCm3;
                DoseValue dv = new DoseValue(double.Parse(eval.Value), du);
                double volume = double.Parse(eval.Value);
                VolumePresentation vpFinal = (evalunit.Value.CompareTo("%") == 0) ? VolumePresentation.Relative : VolumePresentation.AbsoluteCm3;
                DoseValuePresentation dvpFinal = (evalunit.Value.CompareTo("%") == 0) ? DoseValuePresentation.Relative : DoseValuePresentation.Absolute;
                DoseValue dvAchieved = planningItem.PlanningItemObject.GetDoseAtVolume(evalStructure, volume, vp, dvpFinal);
                //checking dose output unit and adapting to template
                if (dvAchieved.UnitAsString.CompareTo(evalunit.Value.ToString()) != 0)
                {
                    if ((evalunit.Value.CompareTo("Gy") == 0) && (dvAchieved.Unit.CompareTo(DoseValue.DoseUnit.Gy) == 0))
                        dvAchieved = new DoseValue(dvAchieved.Dose / 100, DoseValue.DoseUnit.Gy);
                    else
                        return "Unable to calculate";
                }

                return dvAchieved.ToString();
            }
            else if (dvh != null && planningItem.PlanningItemType.ToString() == "PlanSum")
            {
                if ((dvh.SamplingCoverage < 0.9) || (dvh.Coverage < 0.9))
                    return "insufficient dose or sampling coverage";
                Group eval = testMatch[0].Groups["evalpt"];
                Group unit = testMatch[0].Groups["unit"];
                DoseValue.DoseUnit du = (unit.Value.CompareTo("%") == 0) ? DoseValue.DoseUnit.Percent :
                        (unit.Value.CompareTo("Gy") == 0) ? DoseValue.DoseUnit.Gy : DoseValue.DoseUnit.Unknown;
                VolumePresentation vp = (unit.Value.CompareTo("%") == 0) ? VolumePresentation.Relative : VolumePresentation.AbsoluteCm3;
                DoseValue dv = new DoseValue(double.Parse(eval.Value), du);
                double volume = double.Parse(eval.Value);
                //VolumePresentation vpFinal = (evalunit.Value.CompareTo("%") == 0) ? VolumePresentation.Relative : VolumePresentation.AbsoluteCm3;
                DoseValuePresentation dvpFinal = DoseValuePresentation.Absolute;
                DoseValue dvAchieved = planningItem.PlanningItemObject.GetDoseAtVolume(evalStructure, volume, vp, dvpFinal);
                //checking dose output unit and adapting to template
                if (dvAchieved.UnitAsString.CompareTo(evalunit.Value.ToString()) != 0)
                {
                    if ((evalunit.Value.CompareTo("Gy") == 0) && (dvAchieved.Unit.CompareTo(DoseValue.DoseUnit.Gy) == 0))
                        dvAchieved = new DoseValue(dvAchieved.Dose / 100, DoseValue.DoseUnit.Gy);
                    else
                        return "Unable to calculate";
                }

                return dvAchieved.ToString();
            }
            else
            {
                return "Unable to calculate - structure is empty";
            }
        }
    }
}
