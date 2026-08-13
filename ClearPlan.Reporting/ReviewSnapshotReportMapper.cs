using System;
using System.Collections.Generic;
using System.Linq;
using ClearPlan.Core.Review;

namespace ClearPlan.Reporting
{
    public sealed class ReviewSnapshotReportMapper
    {
        public const string SyntheticWatermark =
            "SYNTHETIC DEMONSTRATION — NOT FOR CLINICAL USE";

        public ReviewReportDocument Map(ReviewSnapshot snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException("snapshot");
            }

            ReviewReportMetadata metadata =
                snapshot.Report ?? new ReviewReportMetadata();
            var report = new ReviewReportDocument
            {
                SchemaVersion = snapshot.SchemaVersion,
                ScenarioId = snapshot.ScenarioId,
                ScenarioTitle = snapshot.ScenarioTitle,
                ScenarioDescription = snapshot.ScenarioDescription,
                Seed = snapshot.Seed,
                Synthetic = snapshot.Synthetic,
                GeneratedUtc = snapshot.GeneratedUtc,
                PatientDisplayLabel = snapshot.PatientDisplayLabel,
                PlanDisplayLabel = snapshot.PlanDisplayLabel,
                ProvenanceText = snapshot.ProvenanceText,
                ActivePlanKey = snapshot.ActivePlanKey,
                Title = metadata.Title,
                Subtitle = metadata.Subtitle,
                ModeLabel = snapshot.Synthetic
                    ? SyntheticWatermark
                    : metadata.ModeLabel,
                Watermark = snapshot.Synthetic
                    ? SyntheticWatermark
                    : metadata.Watermark,
                OutputFileLabel = metadata.OutputFileLabel,
                Notes = Copy(metadata.Notes),
                Sources = (snapshot.Sources ?? new List<ReviewSourceStatus>())
                    .Select(MapSource)
                    .ToList(),
                Plans = (snapshot.Plans ?? new List<ReviewPlanRow>())
                    .Select(MapPlan)
                    .ToList(),
                PqmRows = (snapshot.PqmRows ?? new List<ReviewPqmRow>())
                    .Select(MapPqm)
                    .ToList(),
                PlanCheckRows =
                    (snapshot.PlanCheckRows ?? new List<ReviewCheckRow>())
                    .Select(MapCheck)
                    .ToList(),
                FieldRows = (snapshot.FieldRows ?? new List<ReviewFieldRow>())
                    .Select(MapField)
                    .ToList(),
                StructureMappings =
                    (snapshot.StructureMappings ??
                     new List<ReviewStructureMapping>())
                    .Select(MapStructureMapping)
                    .ToList(),
                DvhSeries =
                    (snapshot.DvhSeries ?? new List<ReviewDvhSeries>())
                    .Select(MapDvh)
                    .ToList()
            };

            return report;
        }

        private static ReviewReportSourceRow MapSource(
            ReviewSourceStatus source)
        {
            return new ReviewReportSourceRow
            {
                StableId = source.StableId,
                SourceCode = source.SourceCode,
                SourceType = source.SourceType,
                Status = source.Status,
                Optional = source.Optional,
                UsedFallback = source.UsedFallback,
                PathDisplayLabel = source.PathDisplayLabel,
                Message = source.Message
            };
        }

        private static ReviewReportPlanRow MapPlan(ReviewPlanRow plan)
        {
            return new ReviewReportPlanRow
            {
                PlanKey = plan.PlanKey,
                DisplayLabel = plan.DisplayLabel,
                CreatedUtc = plan.CreatedUtc,
                DosePerFractionGy = plan.DosePerFractionGy,
                TotalDoseGy = plan.TotalDoseGy,
                FractionCount = plan.FractionCount,
                TargetDisplayLabel = plan.TargetDisplayLabel,
                Status = plan.Status
            };
        }

        private static ReviewReportPqmRow MapPqm(ReviewPqmRow pqm)
        {
            return new ReviewReportPqmRow
            {
                StableId = pqm.StableId,
                TemplateCode = pqm.TemplateCode,
                TemplateStructure = pqm.TemplateStructure,
                ResolvedStructureId = pqm.ResolvedStructureId,
                StructureOptions = Copy(pqm.StructureOptions),
                Objective = pqm.Objective,
                Comparator = pqm.Comparator,
                Goal = pqm.Goal,
                Variation = pqm.Variation,
                AchievedValue = pqm.AchievedValue,
                Unit = pqm.Unit,
                Status = pqm.Status,
                Severity = pqm.Severity,
                Explanation = pqm.Explanation
            };
        }

        private static ReviewReportCheckRow MapCheck(ReviewCheckRow check)
        {
            return new ReviewReportCheckRow
            {
                CheckCode = check.CheckCode,
                Category = check.Category,
                Status = check.Status,
                Severity = check.Severity,
                ObservedValue = check.ObservedValue,
                ExpectedValue = check.ExpectedValue,
                Unit = check.Unit,
                Message = check.Message
            };
        }

        private static ReviewReportFieldRow MapField(ReviewFieldRow field)
        {
            return new ReviewReportFieldRow
            {
                StableId = field.StableId,
                TreatmentOrder = field.TreatmentOrder,
                BeamNumber = field.BeamNumber,
                CurrentId = field.CurrentId,
                ExpectedId = field.ExpectedId,
                CurrentName = field.CurrentName,
                SuggestedName = field.SuggestedName,
                IdStatus = field.IdStatus,
                NameStatus = field.NameStatus
            };
        }

        private static ReviewReportStructureMappingRow MapStructureMapping(
            ReviewStructureMapping mapping)
        {
            return new ReviewReportStructureMappingRow
            {
                StableId = mapping.StableId,
                TemplateStructure = mapping.TemplateStructure,
                SelectedStructureId = mapping.SelectedStructureId,
                AvailableStructureIds = Copy(mapping.AvailableStructureIds),
                Status = mapping.Status,
                Message = mapping.Message
            };
        }

        private static ReviewReportDvhSeries MapDvh(ReviewDvhSeries series)
        {
            return new ReviewReportDvhSeries
            {
                StableId = series.StableId,
                StructureId = series.StructureId,
                DisplayName = series.DisplayName,
                Role = series.Role,
                ColorHex = series.ColorHex,
                LineStyle = series.LineStyle,
                Selected = series.Selected,
                VolumeCc = series.VolumeCc,
                DoseUnit = ReviewUnitCodes.Gray,
                VolumeUnit = ReviewUnitCodes.Percent,
                Points = (series.Points ?? new List<ReviewDvhPoint>())
                    .Select(point => new ReviewReportDvhPoint
                    {
                        DoseGy = point.DoseGy,
                        VolumePercent = point.VolumePercent
                    })
                    .ToList()
            };
        }

        private static List<string> Copy(IEnumerable<string> values)
        {
            return values == null
                ? new List<string>()
                : new List<string>(values);
        }
    }
}
