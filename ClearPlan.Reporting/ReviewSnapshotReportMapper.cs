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
                DisabledCheckCount = snapshot.DisabledCheckCount,
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
                    .Where(plan => string.Equals(plan.PlanKey, snapshot.ActivePlanKey, StringComparison.Ordinal))
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
                    .Where(series => series != null)
                    .Select(series => MapDvh(series, snapshot.Synthetic))
                    .ToList(),
                PlanImages = (snapshot.PlanImages ?? new List<ReviewPlanImage>())
                    .Where(image => image != null && string.Equals(image.PlanKey, snapshot.ActivePlanKey, StringComparison.Ordinal))
                    .Select(MapImage).ToList(),
                PlanAnalysis = snapshot.PlanAnalysis == null ? null : ClearPlan.Core.PlanAnalysis.PlanAnalysisSnapshot.Copy(snapshot.PlanAnalysis)
            };

            return report;
        }

        private static ReviewPlanImage MapImage(ReviewPlanImage source)
        {
            return new ReviewPlanImage
            {
                PlanKey = source.PlanKey, Kind = source.Kind, Title = source.Title,
                Caption = source.Caption, SourceStatus = source.SourceStatus,
                UnavailableReason = source.UnavailableReason, Synthetic = source.Synthetic,
                WidthPixels = source.WidthPixels, HeightPixels = source.HeightPixels,
                PixelSpacingXMillimeters = source.PixelSpacingXMillimeters,
                PixelSpacingYMillimeters = source.PixelSpacingYMillimeters,
                GrayscalePixels = source.GrayscalePixels == null ? null : (byte[])source.GrayscalePixels.Clone(),
                LeftOrientation = source.LeftOrientation, RightOrientation = source.RightOrientation,
                TopOrientation = source.TopOrientation, BottomOrientation = source.BottomOrientation,
                IsocenterPixelX = source.IsocenterPixelX, IsocenterPixelY = source.IsocenterPixelY,
                OverlaySummary = source.OverlaySummary,
                DoseFocusRegion = source.DoseFocusRegion == null ? null : new ReviewImageDoseRegion {
                    PrescriptionPercent = source.DoseFocusRegion.PrescriptionPercent, ThresholdGy = source.DoseFocusRegion.ThresholdGy,
                    MinPixelX = source.DoseFocusRegion.MinPixelX, MaxPixelX = source.DoseFocusRegion.MaxPixelX,
                    MinPixelY = source.DoseFocusRegion.MinPixelY, MaxPixelY = source.DoseFocusRegion.MaxPixelY,
                    CoverageLimited = source.DoseFocusRegion.CoverageLimited },
                Overlays = (source.Overlays ?? new List<ReviewImageOverlay>()).Where(item => item != null).Select(item => new ReviewImageOverlay
                {
                    Kind = item.Kind, Label = item.Label, ColorHex = item.ColorHex, Source = item.Source,
                    SourceStatus = item.SourceStatus, UnavailableReason = item.UnavailableReason, DoseGy = item.DoseGy,
                    Paths = (item.Paths ?? new List<ReviewImagePath>()).Where(path => path != null).Select(path => new ReviewImagePath
                    {
                        Closed = path.Closed,
                        Points = (path.Points ?? new List<ReviewImagePoint>()).Where(point => point != null)
                            .Select(point => new ReviewImagePoint { X = point.X, Y = point.Y }).ToList()
                    }).ToList()
                }).ToList()
            };
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
                Explanation = pqm.Explanation, SourceLabel = pqm.SourceLabel, MappingDescription = pqm.MappingDescription
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

        private static ReviewReportDvhSeries MapDvh(ReviewDvhSeries series, bool synthetic)
        {
            var statistics = ReviewDvhStatisticsCalculator.Calculate(
                series.Points, series.MinimumDoseGy, series.MeanDoseGy,
                series.MaximumDoseGy, allowCurveEstimates: synthetic);
            if (series.D98DoseGy.HasValue && !double.IsNaN(series.D98DoseGy.Value) &&
                !double.IsInfinity(series.D98DoseGy.Value) && series.D98DoseGy.Value >= 0)
                statistics.D98Gy = series.D98DoseGy;
            return new ReviewReportDvhSeries
            {
                StableId = series.StableId,
                StructureId = series.StructureId,
                DisplayName = series.DisplayName,
                Role = series.Role,
                ColorHex = series.ColorHex,
                LineStyle = series.LineStyle,
                Selected = series.Selected || series.RequiredForTargetReview,
                TargetKind = series.TargetKind,
                RequiredForTargetReview = series.RequiredForTargetReview,
                TargetSelectionReason = series.TargetSelectionReason,
                VolumeCc = series.VolumeCc,
                DoseUnit = ReviewUnitCodes.Gray,
                VolumeUnit = ReviewUnitCodes.Percent,
                Statistics = statistics,
                Points = (series.Points ?? new List<ReviewDvhPoint>())
                    .Where(point => point != null)
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
