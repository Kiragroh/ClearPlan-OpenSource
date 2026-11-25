using ClearPlan.Reporting;
using System.Linq;
using System;

namespace ClearPlan.Headless
{
    public static class ReportDvhSelection
    {
        public static void ApplyIfEmpty(ReviewReportDocument report)
        {
            if(report==null) throw new ArgumentNullException("report");
            // Constraint assignment is sufficient for display; assessment/achieved values
            // and curve availability never decide which structures should be shown.
            var requested = report.PqmRows.Where(row => row != null)
                .Select(row => row.ResolvedStructureId).ToList();
            foreach(var series in report.DvhSeries.Where(row => row != null))
                series.Selected = series.Selected || series.RequiredForTargetReview || series.TargetKind == "PTV" ||
                    ClearPlan.Core.Review.DvhSelectionPolicy.ShouldSelect(series.StructureId, requested);
        }
    }
}
