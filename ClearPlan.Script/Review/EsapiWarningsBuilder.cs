using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClearPlan.Core.Review;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Review
{
    /// <summary>Read-only native approval-validation messages, not a plan-approval operation.</summary>
    public static class EsapiWarningsBuilder
    {
        public static List<ReviewCheckRow> Build(PlanningItem item, out ReviewSourceStatus source)
        {
            if (System.Windows.Application.Current != null) System.Windows.Application.Current.Dispatcher.VerifyAccess();
            source = new ReviewSourceStatus {
                StableId = "source-eclipse-warnings", SourceCode = "eclipse-validation-messages",
                SourceType = "native-plan-validation", Status = ReviewStatusCodes.Unavailable,
                Optional = true, PathDisplayLabel = "Eclipse native plan-validation messages",
                Message = "Native validation messages unavailable. This is not an empty warning list or plan approval."
            };
            var rows = new List<ReviewCheckRow>();
            var plan = item as PlanSetup;
            if (plan == null) return rows;
            try
            {
                List<PlanValidationResultEsapiDetail> details;
                bool validationResult = plan.IsValidForPlanApproval(out details);
                if (details == null) return rows;
                for (int index = 0; index < details.Count; index++)
                {
                    var detail = details[index];
                    rows.Add(new ReviewCheckRow {
                        CheckCode = "eclipse-warning-" + (index + 1).ToString(CultureInfo.InvariantCulture),
                        Category = "Eclipse warnings", Unit = ReviewUnitCodes.Text,
                        Status = detail.IsError ? ReviewStatusCodes.Fail : ReviewStatusCodes.Variation,
                        Severity = detail.IsError ? ReviewSeverityCodes.Error : ReviewSeverityCodes.Warning,
                        ObservedValue = ClinicalReviewValueMapper.SanitizeClinicalLabel(detail.Code, "Native validation"),
                        Message = ClinicalReviewValueMapper.SanitizeClinicalLabel(detail.MessageForUser, "Native message unavailable")
                    });
                }
                source.Status = ReviewStatusCodes.Available;
                source.Message = string.Format(CultureInfo.InvariantCulture,
                    "{0} native validation message(s), including {1} error(s). API: IsValidForPlanApproval; result={2}. " +
                    "Read-only query, not plan approval. Configured approval extensions are not executed by ESAPI 18; this is not the complete Eclipse warning interface.",
                    rows.Count, details.Count(detail => detail.IsError), validationResult);
            }
            catch (Exception) { /* Preserve explicit unavailable status without exposing vendor exception text. */ }
            return rows;
        }
    }
}
