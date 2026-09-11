using System.Globalization;
using System.Text.RegularExpressions;
using ClearPlan.Core.Constraints;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Calculators
{
    class PQMVolumeAtDose
    {
        public static string GetVolumeAtDose(StructureSet structureSet, PlanningItemViewModel planningItem, Structure evalStructure, MatchCollection testMatch, Group evalunit)
        {
            Group eval = testMatch[0].Groups["evalpt"];
            Group unit = testMatch[0].Groups["unit"];
            double dose;
            if (!PqmNumericEvaluator.TryParseNumber(eval.Value, out dose) || dose < 0 ||
                (evalunit.Value != "%" && evalunit.Value != "cc"))
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
            if (!(planningItem.PlanningItemObject is PlanSetup) && !(planningItem.PlanningItemObject is PlanSum))
                return "Unable to calculate - planning item is unsupported";

            DVHData dvh = planningItem.PlanningItemObject.GetDVHCumulativeData(evalStructure, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
            if (dvh == null) return "Unable to calculate - structure is empty";
            if (dvh.SamplingCoverage < 0.9 || dvh.Coverage < 0.9)
                return "insufficient dose or sampling coverage";

            // Preserve the existing target treatment-percentage behavior; no new Rx scaling.
            if (planningItem.PlanningItemObject is PlanSetup && evalStructure.Id == planningItem.PlanningItemTargetId)
                dose *= planningItem.PlanningItemTreatmentPercentage;
            if (double.IsNaN(dose) || double.IsInfinity(dose))
                return "Unable to calculate - non-finite dose";
            DoseValue threshold = new DoseValue(dose, doseUnit);
            VolumePresentation volumePresentation = evalunit.Value == "%" ? VolumePresentation.Relative : VolumePresentation.AbsoluteCm3;
            double achieved = planningItem.PlanningItemObject.GetVolumeAtDose(evalStructure, threshold, volumePresentation);
            if (double.IsNaN(achieved) || double.IsInfinity(achieved) || achieved < 0)
                return "Unable to calculate - non-finite volume";
            return string.Format(CultureInfo.InvariantCulture, "{0:0.00} {1}", achieved, evalunit.Value);
        }
    }
}
