using System;
using System.Globalization;
using System.Text.RegularExpressions;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Calculators
{
    class PQMMinMaxMean
    {
        public static string GetMinMaxMean(StructureSet structureSet, PlanningItemViewModel planningItem, Structure evalStructure, MatchCollection testMatch, Group evalunit, Group type)
        {
            if (type == null || evalunit == null || !type.Success || !evalunit.Success ||
                evalStructure == null || evalStructure.IsEmpty)
                return "Not evaluated - structure or metric is unavailable";

            if (type.Value == "Volume")
            {
                if (evalunit.Value != "cc") return "Not evaluated - structure volume requires cc";
                double volume = evalStructure.Volume;
                if (double.IsNaN(volume) || double.IsInfinity(volume) || volume <= 0)
                    return "Not evaluated - structure volume is invalid";
                return string.Format(CultureInfo.InvariantCulture, "{0:0.00} cc", volume);
            }

            if (type.Value != "Min" && type.Value != "Max" && type.Value != "Mean")
                return "Not evaluated - unsupported dose statistic";
            if (evalunit.Value != "Gy" && evalunit.Value != "cGy" && evalunit.Value != "%")
                return "Not evaluated - unsupported dose unit";
            if (planningItem == null || planningItem.PlanningItemObject == null)
                return "Not evaluated - planning item is unavailable";
            if (planningItem.PlanningItemObject is PlanSum && evalunit.Value == "%")
                return "Not evaluated - relative dose for plan sum is unsupported";
            if (!(planningItem.PlanningItemObject is PlanSetup) && !(planningItem.PlanningItemObject is PlanSum))
                return "Not evaluated - planning item is unsupported";

            DoseValuePresentation presentation = evalunit.Value == "%"
                ? DoseValuePresentation.Relative : DoseValuePresentation.Absolute;
            DVHData dvh = planningItem.PlanningItemObject.GetDVHCumulativeData(
                evalStructure, presentation, VolumePresentation.Relative, 0.1);
            if (dvh == null) return "Not evaluated - DVH is unavailable";
            if (double.IsNaN(dvh.SamplingCoverage) || double.IsInfinity(dvh.SamplingCoverage) ||
                double.IsNaN(dvh.Coverage) || double.IsInfinity(dvh.Coverage) ||
                dvh.SamplingCoverage < 0.9 || dvh.Coverage < 0.9)
                return "Not evaluated - insufficient dose or sampling coverage";

            DoseValue achieved = type.Value == "Min" ? dvh.MinDose :
                type.Value == "Max" ? dvh.MaxDose : dvh.MeanDose;
            return PQMDoseAtVolume.FormatDoseValue(achieved, evalunit.Value);
        }
    }
}
