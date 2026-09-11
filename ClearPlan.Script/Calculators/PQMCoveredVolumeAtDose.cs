using System.Globalization;
using System.Text.RegularExpressions;
using ClearPlan.Core.Constraints;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Calculators
{
    class PQMCoveredVolumeAtDose
    {
        public static string GetCoveredVolumeAtDose(StructureSet structureSet, PlanningItemViewModel planningItem, Structure evalStructure, MatchCollection testMatch, Group evalunit, string variation)
        {
            Group eval = testMatch[0].Groups["evalpt"];
            Group unit = testMatch[0].Groups["unit"];
            double dose, variationValue = 0;
            if (!PqmNumericEvaluator.TryParseNumber(eval.Value, out dose) || dose < 0 ||
                (evalunit.Value != "%" && evalunit.Value != "cc") ||
                (!string.IsNullOrWhiteSpace(variation) && (!PqmNumericEvaluator.TryParseNumber(variation, out variationValue) || variationValue < 0)))
                return "Unable to calculate - invalid parameter or unit";
            DoseValue.DoseUnit doseUnit;
            switch (unit.Value)
            {
                case "%": doseUnit = DoseValue.DoseUnit.Percent; break;
                case "Gy": doseUnit = DoseValue.DoseUnit.Gy; break;
                case "cGy": doseUnit = DoseValue.DoseUnit.cGy; break;
                default: return "Unable to calculate - unknown dose unit";
            }
            if (planningItem.PlanningItemObject is PlanSum && doseUnit == DoseValue.DoseUnit.Percent)
                return "Unable to calculate - relative dose for plan sum is unsupported";
            DVHData dvh = planningItem.PlanningItemObject.GetDVHCumulativeData(evalStructure, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
            if (dvh == null) return "Unable to calculate - structure is empty";
            if (dvh.SamplingCoverage < 0.9 || dvh.Coverage < 0.9)
                return "insufficient dose or sampling coverage";
            DoseValue threshold = new DoseValue(dose, doseUnit);
            VolumePresentation volumePresentation = evalunit.Value == "%" ? VolumePresentation.Relative : VolumePresentation.AbsoluteCm3;
            double volumeAchieved = planningItem.PlanningItemObject.GetVolumeAtDose(evalStructure, threshold, volumePresentation);
            double coveredVolume;
            if (!PqmNumericEvaluator.TryComplementVolume(volumeAchieved, evalStructure.Volume, evalunit.Value,
                string.IsNullOrWhiteSpace(variation) ? (double?)null : variationValue, out coveredVolume))
                return "Unable to calculate - invalid volume or volume too small";
            return string.Format(CultureInfo.InvariantCulture, "{0:0.00} {1}", coveredVolume, evalunit.Value);
        }
    }
}
