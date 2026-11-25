using System;
using ClearPlan.Core.Review;
using ClearPlan.Reporting;

namespace ClearPlan.Headless
{
    public static class ReportDoseSafety
    {
        public static void Apply(ReviewReportDocument report, bool hasDose, bool? doseValid)
        {
            if (report == null) throw new ArgumentNullException("report");
            if (report.Synthetic) throw new ArgumentException("Clinical dose safety requires a clinical report.");
            bool valid = hasDose && doseValid == true;
            report.ModeLabel = valid ? "READ-ONLY DEVELOPMENT REVIEW - NOT AN APPROVAL"
                : !hasDose ? "MISSING TPS DOSE - REVIEW ONLY"
                : doseValid == false ? "INVALID TPS DOSE - REVIEW ONLY"
                : "UNVERIFIED TPS DOSE - REVIEW ONLY";
            report.Watermark = valid ? "Development preview; not a clinical release or approval." : report.ModeLabel;
            report.PlanCheckRows.Insert(0, new ReviewReportCheckRow {
                CheckCode = "native-dose-validity", Category = "Dose validity",
                Status = valid ? ReviewStatusCodes.Info : ReviewStatusCodes.Fail,
                Severity = valid ? ReviewSeverityCodes.Info : ReviewSeverityCodes.Error,
                ObservedValue = !hasDose ? "Dose missing" : doseValid.HasValue ? doseValid.Value.ToString() : "Unknown",
                ExpectedValue = "Dose present and native PlanSetup.IsDoseValid=True",
                Unit = ReviewUnitCodes.Text,
                Message = valid ? "Native dose is valid; this does not approve the plan or these development checks."
                    : "Native dose is missing, invalid or unverified. Available cached DVHs are shown for troubleshooting only. PQM is not evaluated."
            });
            if (!valid)
            {
                report.Notes.Insert(0, "DOSE WARNING: available cached dose/DVH values are not evidence of a valid current calculation. No clinical decision or delivery approval may be based on this preview.");
                foreach (var row in report.PqmRows)
                {
                    row.Status = ReviewStatusCodes.NotEvaluated;
                    row.Severity = ReviewSeverityCodes.Warning;
                    row.Explanation = "Not evaluated: native dose is not valid. Any displayed value is a cached troubleshooting value. " + row.Explanation;
                }
                foreach (var plan in report.Plans) plan.Status = ReviewStatusCodes.Fail;
            }
        }
    }
}
