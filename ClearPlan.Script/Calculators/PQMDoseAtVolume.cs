using System.Globalization;
using System.Text.RegularExpressions;
using ClearPlan.Core.Constraints;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Calculators
{
    class PQMDoseAtVolume
    {
        public static string GetDoseAtVolume(StructureSet structureSet, PlanningItemViewModel planningItem, Structure evalStructure, MatchCollection testMatch, Group evalunit)
        {
            Group eval = testMatch[0].Groups["evalpt"];
            Group unit = testMatch[0].Groups["unit"];
            double volume;
            if (!PqmNumericEvaluator.TryParseNumber(eval.Value, out volume) || volume < 0 ||
                (unit.Value == "%" && volume > 100) ||
                (unit.Value != "%" && unit.Value != "cc") ||
                (evalunit.Value != "%" && evalunit.Value != "Gy" && evalunit.Value != "cGy"))
                return "Unable to calculate - invalid parameter or unit";
            if (planningItem.PlanningItemObject is PlanSum && evalunit.Value == "%")
                return "Unable to calculate - relative dose for plan sum is unsupported";

            DVHData dvh = planningItem.PlanningItemObject.GetDVHCumulativeData(evalStructure, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
            if (dvh == null) return "Unable to calculate - structure is empty";
            if (dvh.SamplingCoverage < 0.9 || dvh.Coverage < 0.9)
                return "insufficient dose or sampling coverage";
            VolumePresentation volumePresentation = unit.Value == "%" ? VolumePresentation.Relative : VolumePresentation.AbsoluteCm3;
            DoseValuePresentation dosePresentation = evalunit.Value == "%" ? DoseValuePresentation.Relative : DoseValuePresentation.Absolute;
            DoseValue achieved = planningItem.PlanningItemObject.GetDoseAtVolume(evalStructure, volume, volumePresentation, dosePresentation);
            return FormatDoseValue(achieved, evalunit.Value);
        }

        // Never interpret Unknown or convert relative dose without an explicit prescription.
        public static string FormatDoseValue(DoseValue achieved, string requestedUnit)
        {
            string actualUnit;
            switch (achieved.Unit)
            {
                case DoseValue.DoseUnit.Gy: actualUnit = "Gy"; break;
                case DoseValue.DoseUnit.cGy: actualUnit = "cGy"; break;
                case DoseValue.DoseUnit.Percent: actualUnit = "%"; break;
                default: return "Unable to calculate - unknown dose unit";
            }
            double value;
            if (!PqmNumericEvaluator.TryConvertDose(achieved.Dose, actualUnit, requestedUnit, out value))
                return "Unable to calculate - incompatible dose units";
            return value.ToString("0.00", CultureInfo.InvariantCulture) + " " + requestedUnit;
        }
    }
}
