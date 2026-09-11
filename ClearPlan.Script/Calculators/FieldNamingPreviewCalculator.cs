using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
using ClearPlan.Core.Fields;
using VMS.TPS.Common.Model.API;

namespace ClearPlan.Calculators
{
    public static class FieldNamingPreviewCalculator
    {
        public static ObservableCollection<FieldNamePreviewViewModel> Calculate(
            PlanningItemViewModel planningItem,
            FieldNamingRuleConfiguration configuration)
        {
            var result = new ObservableCollection<FieldNamePreviewViewModel>();
            PlanSetup plan = planningItem == null
                ? null
                : planningItem.PlanningItemObject as PlanSetup;
            if (plan == null)
            {
                return result;
            }

            bool treatmentOrderAvailable;
            IList<Beam> treatmentBeams = GetTreatmentBeams(
                plan,
                out treatmentOrderAvailable);
            var inputs = new List<BeamNamingInput>();
            for (int index = 0; index < treatmentBeams.Count; index++)
            {
                Beam beam = treatmentBeams[index];
                ControlPoint first = beam.ControlPoints.First();
                ControlPoint last = beam.ControlPoints.Last();
                inputs.Add(new BeamNamingInput
                {
                    CurrentId = beam.Id,
                    CurrentName = beam.Name,
                    BeamNumber = beam.BeamNumber,
                    TreatmentOrderIndex = treatmentOrderAvailable
                        ? (int?)index
                        : null,
                    GantryStartAngle = first.GantryAngle,
                    GantryStopAngle = last.GantryAngle,
                    PatientSupportAngle = first.PatientSupportAngle,
                    GantryDirection = beam.GantryDirection.ToString()
                });
            }

            foreach (FieldNameSuggestion suggestion in
                FieldNameSuggester.Suggest(plan.Id, inputs, configuration == null ? null : configuration.Rules))
            {
                result.Add(new FieldNamePreviewViewModel
                {
                    Order = suggestion.DisplayOrder,
                    BeamNumber = suggestion.BeamNumber,
                    CurrentId = suggestion.CurrentId,
                    ExpectedId = suggestion.ExpectedId,
                    CurrentName = suggestion.CurrentName,
                    SuggestedName = suggestion.SuggestedName,
                    IdWouldChange = suggestion.IdWouldChange,
                    NameWouldChange = suggestion.NameWouldChange,
                    WouldChange = suggestion.WouldChange,
                    IsEvaluated = suggestion.IsEvaluated,
                    EvaluationMessage = suggestion.IsEvaluated ? configuration.Message :
                        configuration == null ? suggestion.EvaluationMessage : configuration.Message
                });
            }

            return result;
        }

        private static IList<Beam> GetTreatmentBeams(
            PlanSetup plan,
            out bool treatmentOrderAvailable)
        {
            try
            {
                PropertyInfo property = plan.GetType().GetProperty(
                    "BeamsInTreatmentOrder",
                    BindingFlags.Instance | BindingFlags.Public);
                IEnumerable value = property == null
                    ? null
                    : property.GetValue(plan, null) as IEnumerable;
                if (value == null)
                {
                    throw new MissingMemberException(
                        plan.GetType().FullName,
                        "BeamsInTreatmentOrder");
                }

                IList<Beam> ordered = value.Cast<object>()
                    .OfType<Beam>()
                    .Where(beam => !beam.IsSetupField)
                    .ToList();
                treatmentOrderAvailable = true;
                return ordered;
            }
            catch
            {
                treatmentOrderAvailable = false;
                return plan.Beams
                    .Where(beam => !beam.IsSetupField)
                    .OrderBy(beam => beam.BeamNumber)
                    .ThenBy(beam => beam.Id)
                    .ToList();
            }
        }
    }
}
