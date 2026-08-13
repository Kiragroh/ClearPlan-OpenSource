using System;
using ClearPlan.Core.Constraints;

namespace ClearPlan.Mappers
{
    internal static class PqmObjectiveViewModelMapper
    {
        public static PQMSummaryViewModel Map(ConstraintDefinition constraint)
        {
            PqmObjectiveDefinition definition = PqmObjectiveMapper.Map(constraint);
            return new PQMSummaryViewModel
            {
                ConstraintId = definition.ConstraintId,
                TemplateId = definition.TemplateId,
                TemplateAliases = ToArray(definition.Aliases),
                TemplateCodes = ToArray(definition.Codes),
                TemplateType = ToArray(definition.DicomTypes),
                DVHObjective = definition.DvhObjective,
                Goal = definition.Goal,
                Variation = definition.Variation,
                Priority = definition.Priority,
                Source = definition.Source,
                Comment = definition.Comment,
                Met = string.Empty,
                Achieved = string.Empty
            };
        }

        private static string[] ToArray(System.Collections.Generic.IList<string> values)
        {
            if (values == null || values.Count == 0)
            {
                return new string[0];
            }

            var result = new string[values.Count];
            values.CopyTo(result, 0);
            return result;
        }
    }
}
