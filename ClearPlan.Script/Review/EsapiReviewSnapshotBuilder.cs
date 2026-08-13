using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClearPlan.Core.Constraints;
using ClearPlan.Core.Review;
using ClearPlan.Helpers;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Review
{
    /// <summary>
    /// Projects already-calculated clinical review data into detached DTOs.
    /// Build must run synchronously on the owning ESAPI/WPF thread.
    /// </summary>
    public sealed class EsapiReviewSnapshotBuilder
    {
        private const double DvhBinWidthGy = 0.1;
        private const double DvhTolerance = 0.001;

        public ReviewSnapshot Build(
            MainViewModel source,
            ClearPlanSettings settings)
        {
            VerifyOwningDispatcher();

            if (source == null)
            {
                throw new ArgumentNullException("source");
            }

            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }
            if (source.ActivePlanningItem == null ||
                source.ActivePlanningItem.PlanningItemObject == null)
            {
                throw new InvalidOperationException(
                    "An active Eclipse planning item is required.");
            }

            var snapshot = new ReviewSnapshot
            {
                SchemaVersion = ReviewSnapshot.CurrentSchemaVersion,
                ScenarioId = "clinical-read-only",
                ScenarioTitle = "Clinical plan review",
                ScenarioDescription =
                    "Read-only projection of the active Eclipse review.",
                Seed = 0,
                Synthetic = false,
                GeneratedUtc = DateTimeOffset.UtcNow,
                PatientDisplayLabel = "Current Eclipse patient",
                PlanDisplayLabel = BuildPlanDisplayLabel(
                    source.ActivePlanningItem),
                ProvenanceText =
                    "Eclipse ESAPI clinical read-only projection.",
                Report = new ReviewReportMetadata
                {
                    Title = "ClearPlan clinical review",
                    Subtitle = "Active Eclipse planning item",
                    ModeLabel = "CLINICAL READ-ONLY",
                    Watermark = string.Empty,
                    OutputFileLabel = "clearplan-clinical-review.pdf",
                    Notes = new List<string>
                    {
                        "Dose is expressed in Gy.",
                        "Cumulative DVH volume is expressed as relative volume in percent.",
                        "PlanCheck rows are projected from existing results and are not recalculated."
                    }
                }
            };

            snapshot.Sources = BuildSources(source, settings);

            var usedCheckIds = new HashSet<string>(StringComparer.Ordinal);
            snapshot.PqmRows = BuildPqmRows(source);
            snapshot.PlanCheckRows = BuildPlanCheckRows(
                source,
                usedCheckIds);
            snapshot.FieldRows = BuildFieldRows(source);
            snapshot.DvhSeries = BuildDvhSeries(
                source,
                snapshot.PlanCheckRows,
                usedCheckIds);
            snapshot.StructureMappings = BuildStructureMappings(
                snapshot.PqmRows,
                snapshot.DvhSeries);
            snapshot.Plans = BuildPlans(
                source,
                GetAggregateStatus(snapshot),
                out string activePlanKey);
            snapshot.ActivePlanKey = activePlanKey;

            ReviewSnapshotValidationResult validation =
                ReviewSnapshotValidator.Validate(snapshot);
            if (!validation.IsValid)
            {
                string details = string.Join(
                    "; ",
                    validation.Issues
                        .Take(8)
                        .Select(issue =>
                            issue.Code + " at " + issue.Path));
                throw new InvalidOperationException(
                    "Clinical review snapshot validation failed: " + details);
            }

            return snapshot;
        }

        private static void VerifyOwningDispatcher()
        {
            if (System.Windows.Application.Current != null &&
                System.Windows.Application.Current.Dispatcher != null)
            {
                System.Windows.Application.Current.Dispatcher.VerifyAccess();
            }
        }

        private static List<ReviewSourceStatus> BuildSources(
            MainViewModel source,
            ClearPlanSettings settings)
        {
            var sources = new List<ReviewSourceStatus>
            {
                new ReviewSourceStatus
                {
                    StableId = "source-esapi",
                    SourceCode = "eclipse-esapi",
                    SourceType = "clinical-session",
                    Status = ReviewStatusCodes.Available,
                    Optional = false,
                    UsedFallback = false,
                    PathDisplayLabel = "Active Eclipse planning item",
                    Message = "Clinical data projected on the ESAPI thread."
                }
            };

            ConstraintCatalogLoadResult catalog = source.ConstraintCatalogStatus;
            string status;
            bool usedFallback;
            int warnings = catalog == null || catalog.Warnings == null
                ? 0
                : catalog.Warnings.Count;
            int errors = catalog == null || catalog.Errors == null
                ? 0
                : catalog.Errors.Count;

            if (catalog == null)
            {
                status = ReviewStatusCodes.NotConfigured;
                usedFallback = false;
            }
            else if (!catalog.IsUsable)
            {
                status = ReviewStatusCodes.Unavailable;
                usedFallback = false;
            }
            else if (warnings > 0)
            {
                status = ReviewStatusCodes.Fallback;
                usedFallback = true;
            }
            else
            {
                status = ReviewStatusCodes.Available;
                usedFallback = false;
            }

            string sourceLabel = ClinicalReviewValueMapper.SanitizeClinicalLabel(
                catalog == null ? null : catalog.ActiveSource,
                "Constraint source");
            string mode = settings.ConstraintSource == null
                ? "Automatic"
                : settings.ConstraintSource.Mode.ToString();
            sources.Add(new ReviewSourceStatus
            {
                StableId = "source-constraints",
                SourceCode = "constraint-catalog",
                SourceType = ClinicalReviewValueMapper.SanitizeClinicalLabel(
                    mode,
                    "configured"),
                Status = status,
                Optional = false,
                UsedFallback = usedFallback,
                PathDisplayLabel = sourceLabel + " constraints",
                Message = string.Format(
                    CultureInfo.InvariantCulture,
                    "Constraint catalog status: {0} warning(s), {1} error(s).",
                    warnings,
                    errors)
            });

            return sources;
        }

        private static List<ReviewPqmRow> BuildPqmRows(MainViewModel source)
        {
            var rows = new List<ReviewPqmRow>();
            var usedIds = new HashSet<string>(StringComparer.Ordinal);
            List<string> structureOptions = (source.StructureList ??
                    new System.Collections.ObjectModel.ObservableCollection<
                        StructureViewModel>())
                .Where(item => item != null)
                .Select(item => ClinicalReviewValueMapper.SanitizeClinicalLabel(
                    item.StructureName,
                    "Structure"))
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(item => item, StringComparer.Ordinal)
                .ToList();

            int index = 0;
            foreach (PQMSummaryViewModel item in source.PqmSummaries ??
                     new System.Collections.ObjectModel.ObservableCollection<
                         PQMSummaryViewModel>())
            {
                index++;
                if (item == null)
                {
                    continue;
                }

                ReviewStatusAndSeverity mapped =
                    ClinicalReviewValueMapper.MapPqmStatus(item.Met);
                string templateStructure =
                    ClinicalReviewValueMapper.SanitizeClinicalLabel(
                        item.TemplateId,
                        "Template " + index);
                string resolvedStructure =
                    string.IsNullOrWhiteSpace(item.StructureName)
                        ? string.Empty
                        : ClinicalReviewValueMapper.SanitizeClinicalLabel(
                            item.StructureName,
                            "Structure");
                string objective =
                    ClinicalReviewValueMapper.SanitizeClinicalLabel(
                        item.DVHObjective,
                        "Objective");
                string unit = InferPqmUnit(
                    item.DVHObjective,
                    item.Goal,
                    item.Achieved);

                rows.Add(new ReviewPqmRow
                {
                    StableId =
                        ClinicalReviewValueMapper.CreateUniqueStableId(
                            string.Join(
                                "-",
                                new[]
                                {
                                    item.ConstraintId,
                                    item.TemplateId,
                                    item.DVHObjective
                                }.Where(value =>
                                    !string.IsNullOrWhiteSpace(value))),
                            "pqm-" + index,
                            usedIds),
                    TemplateCode =
                        ClinicalReviewValueMapper.SanitizeClinicalLabel(
                            string.IsNullOrWhiteSpace(item.ConstraintId)
                                ? item.TemplateId
                                : item.ConstraintId,
                            "PQM " + index),
                    TemplateStructure = templateStructure,
                    ResolvedStructureId = resolvedStructure,
                    StructureOptions = new List<string>(structureOptions),
                    Objective = objective,
                    Comparator =
                        ClinicalReviewValueMapper.ExtractComparator(item.Goal),
                    Goal = ParsePqmValue(item.Goal, unit),
                    Variation = ParsePqmValue(item.Variation, unit),
                    AchievedValue = ParsePqmValue(item.Achieved, unit),
                    Unit = unit,
                    Status = mapped.Status,
                    Severity = mapped.Severity,
                    Explanation = BuildPqmExplanation(item)
                });
            }

            return rows;
        }

        private static List<ReviewCheckRow> BuildPlanCheckRows(
            MainViewModel source,
            ISet<string> usedIds)
        {
            var rows = new List<ReviewCheckRow>();
            int index = 0;
            foreach (ErrorViewModel item in source.ErrorGrid ??
                     new List<ErrorViewModel>())
            {
                index++;
                if (item == null)
                {
                    continue;
                }

                ReviewStatusAndSeverity mapped =
                    ClinicalReviewValueMapper.MapPlanCheckStatus(item.Status);
                rows.Add(new ReviewCheckRow
                {
                    CheckCode =
                        ClinicalReviewValueMapper.CreateUniqueStableId(
                            item.Severity,
                            "plan-check-" + index,
                            usedIds),
                    Category = "PlanCheck",
                    Status = mapped.Status,
                    Severity = mapped.Severity,
                    ObservedValue = string.Empty,
                    ExpectedValue = string.Empty,
                    Unit = ReviewUnitCodes.Text,
                    Message =
                        ClinicalReviewValueMapper.SanitizeClinicalLabel(
                            item.Description,
                            "Clinical PlanCheck finding")
                });
            }

            int referenceIndex = 0;
            foreach (RefViewModel item in source.RefGrid ??
                     new List<RefViewModel>())
            {
                referenceIndex++;
                if (item == null)
                {
                    continue;
                }

                string referenceLabel =
                    ClinicalReviewValueMapper.SanitizeClinicalLabel(
                        item.RefPointId,
                        "Reference " + referenceIndex);
                rows.Add(new ReviewCheckRow
                {
                    CheckCode =
                        ClinicalReviewValueMapper.CreateUniqueStableId(
                            "ref-" + referenceLabel,
                            "ref-" + referenceIndex,
                            usedIds),
                    Category = "Reference point",
                    Status = ReviewStatusCodes.Info,
                    Severity = ReviewSeverityCodes.Info,
                    ObservedValue = string.Format(
                        CultureInfo.InvariantCulture,
                        "Rx {0:0.###}; session {1:0.###}; fractions {2:0.###}; D50 {3:0.###}; D95 {4:0.###}; D2 {5:0.###}; D98 {6:0.###}",
                        item.Prescription,
                        item.Session,
                        item.Fractions,
                        item.D50,
                        item.D95,
                        item.D2,
                        item.D98),
                    ExpectedValue = string.Empty,
                    Unit = ReviewUnitCodes.Gray,
                    Message = "Reference-point dose summary: " +
                              referenceLabel
                });
            }

            return rows;
        }

        private static List<ReviewFieldRow> BuildFieldRows(MainViewModel source)
        {
            var rows = new List<ReviewFieldRow>();
            var usedIds = new HashSet<string>(StringComparer.Ordinal);
            int index = 0;
            foreach (FieldNamePreviewViewModel item in
                     source.FieldNamePreviews ??
                     new System.Collections.ObjectModel.ObservableCollection<
                         FieldNamePreviewViewModel>())
            {
                index++;
                if (item == null)
                {
                    continue;
                }

                string currentId =
                    ClinicalReviewValueMapper.SanitizeClinicalLabel(
                        item.CurrentId,
                        "Field " + index);
                string expectedId =
                    ClinicalReviewValueMapper.SanitizeClinicalLabel(
                        item.ExpectedId,
                        currentId);
                string currentName =
                    ClinicalReviewValueMapper.SanitizeClinicalLabel(
                        item.CurrentName,
                        "Field name " + index);
                string suggestedName =
                    ClinicalReviewValueMapper.SanitizeClinicalLabel(
                        item.SuggestedName,
                        currentName);

                rows.Add(new ReviewFieldRow
                {
                    StableId =
                        ClinicalReviewValueMapper.CreateUniqueStableId(
                            currentId,
                            "field-" + index,
                            usedIds),
                    TreatmentOrder = item.Order,
                    BeamNumber = item.BeamNumber,
                    CurrentId = currentId,
                    ExpectedId = expectedId,
                    CurrentName = currentName,
                    SuggestedName = suggestedName,
                    IdStatus =
                        ClinicalReviewValueMapper.MapFieldStatus(
                            currentId,
                            expectedId),
                    NameStatus =
                        ClinicalReviewValueMapper.MapFieldStatus(
                            currentName,
                            suggestedName)
                });
            }

            return rows;
        }

        private static List<ReviewDvhSeries> BuildDvhSeries(
            MainViewModel source,
            IList<ReviewCheckRow> findings,
            ISet<string> usedCheckIds)
        {
            var series = new List<ReviewDvhSeries>();
            var usedSeriesIds = new HashSet<string>(StringComparer.Ordinal);
            int index = 0;

            foreach (DvhStructureViewModel item in source.DvhStructures ??
                     new System.Collections.ObjectModel.ObservableCollection<
                         DvhStructureViewModel>())
            {
                index++;
                if (item == null || item.Structure == null)
                {
                    continue;
                }

                string structureId =
                    ClinicalReviewValueMapper.SanitizeClinicalLabel(
                        item.Id,
                        "Structure " + index);
                try
                {
                    DVHData data =
                        source.ActivePlanningItem.PlanningItemObject
                            .GetDVHCumulativeData(
                                item.Structure,
                                DoseValuePresentation.Absolute,
                                VolumePresentation.Relative,
                                DvhBinWidthGy);
                    if (data == null ||
                        data.CurveData == null ||
                        data.CurveData.Length == 0)
                    {
                        throw new InvalidOperationException(
                            "DVH data are unavailable.");
                    }

                    series.Add(new ReviewDvhSeries
                    {
                        StableId =
                            ClinicalReviewValueMapper.CreateUniqueStableId(
                                "dvh-" + structureId,
                                "dvh-" + index,
                                usedSeriesIds),
                        StructureId = structureId,
                        DisplayName = structureId,
                        Role = GetDvhRole(item.Structure, structureId),
                        ColorHex = GetColorHex(item.Structure, index),
                        LineStyle = ReviewLineStyleCodes.Solid,
                        Selected = item.IsSelected,
                        VolumeCc = GetFiniteVolume(item.Structure),
                        Points = BuildDvhPoints(data)
                    });
                }
                catch (Exception)
                {
                    findings.Add(new ReviewCheckRow
                    {
                        CheckCode =
                            ClinicalReviewValueMapper.CreateUniqueStableId(
                                "dvh-unavailable-" + structureId,
                                "dvh-unavailable-" + index,
                                usedCheckIds),
                        Category = "DVH",
                        Status = ReviewStatusCodes.Info,
                        Severity = ReviewSeverityCodes.Info,
                        ObservedValue = "Unavailable",
                        ExpectedValue = "Cumulative DVH",
                        Unit = ReviewUnitCodes.Text,
                        Message = "DVH unavailable for " + structureId + "."
                    });
                }
            }

            return series;
        }

        private static List<ReviewDvhPoint> BuildDvhPoints(DVHData data)
        {
            var points = new List<ReviewDvhPoint>();
            double? previousDose = null;
            double? previousVolume = null;

            foreach (DVHPoint point in data.CurveData)
            {
                double doseGy = ConvertDoseToGray(point.DoseValue);
                double volume = point.Volume;
                if (double.IsNaN(volume) || double.IsInfinity(volume))
                {
                    throw new InvalidOperationException(
                        "DVH volume must be finite.");
                }

                if (doseGy < -DvhTolerance ||
                    (previousDose.HasValue &&
                     doseGy + DvhTolerance < previousDose.Value))
                {
                    throw new InvalidOperationException(
                        "DVH dose coordinates are not increasing.");
                }

                if (volume < -DvhTolerance ||
                    volume > 100.0 + DvhTolerance)
                {
                    throw new InvalidOperationException(
                        "Relative DVH volume is outside 0 to 100 percent.");
                }

                doseGy = Math.Max(0.0, doseGy);
                volume = Math.Max(0.0, Math.Min(100.0, volume));
                if (previousVolume.HasValue &&
                    volume > previousVolume.Value)
                {
                    if (volume - previousVolume.Value > DvhTolerance)
                    {
                        throw new InvalidOperationException(
                            "Cumulative DVH volume is increasing.");
                    }

                    volume = previousVolume.Value;
                }

                points.Add(new ReviewDvhPoint(doseGy, volume));
                previousDose = doseGy;
                previousVolume = volume;
            }

            return points;
        }

        private static List<ReviewStructureMapping> BuildStructureMappings(
            IList<ReviewPqmRow> pqmRows,
            IList<ReviewDvhSeries> dvhSeries)
        {
            var mappings = new List<ReviewStructureMapping>();
            var usedIds = new HashSet<string>(StringComparer.Ordinal);
            List<string> available = dvhSeries
                .Where(item => item != null)
                .Select(item => item.StructureId)
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(item => item, StringComparer.Ordinal)
                .ToList();

            int index = 0;
            foreach (IGrouping<string, ReviewPqmRow> group in pqmRows
                         .Where(item => item != null)
                         .GroupBy(
                             item => item.TemplateStructure ?? string.Empty,
                             StringComparer.Ordinal))
            {
                index++;
                List<string> selected = group
                    .Select(item => item.ResolvedStructureId)
                    .Where(item =>
                        !string.IsNullOrWhiteSpace(item) &&
                        available.Contains(item, StringComparer.Ordinal))
                    .Distinct(StringComparer.Ordinal)
                    .ToList();
                string status;
                string message;
                string selectedStructure;
                if (selected.Count == 1)
                {
                    status = ReviewStatusCodes.Pass;
                    message = "Structure mapping resolved.";
                    selectedStructure = selected[0];
                }
                else if (selected.Count > 1)
                {
                    status = ReviewStatusCodes.Variation;
                    message = "Structure mapping is ambiguous.";
                    selectedStructure = string.Empty;
                }
                else
                {
                    status = ReviewStatusCodes.NotEvaluated;
                    message = "Structure mapping is not resolved.";
                    selectedStructure = string.Empty;
                }

                string template =
                    ClinicalReviewValueMapper.SanitizeClinicalLabel(
                        group.Key,
                        "Template " + index);
                mappings.Add(new ReviewStructureMapping
                {
                    StableId =
                        ClinicalReviewValueMapper.CreateUniqueStableId(
                            "mapping-" + template,
                            "mapping-" + index,
                            usedIds),
                    TemplateStructure = template,
                    SelectedStructureId = selectedStructure,
                    AvailableStructureIds = new List<string>(available),
                    Status = status,
                    Message = message
                });
            }

            return mappings;
        }

        private static List<ReviewPlanRow> BuildPlans(
            MainViewModel source,
            string activeStatus,
            out string activePlanKey)
        {
            var rows = new List<ReviewPlanRow>();
            var usedIds = new HashSet<string>(StringComparer.Ordinal);
            activePlanKey = null;

            IEnumerable<PlanningItemDetailsViewModel> summaries =
                source.PlanningItemSummaries ??
                new System.Collections.ObjectModel.ObservableCollection<
                    PlanningItemDetailsViewModel>();
            foreach (PlanningItemDetailsViewModel summary in summaries)
            {
                if (summary == null)
                {
                    continue;
                }

                PlanningItemViewModel item = FindPlanningItem(source, summary);
                bool active = IsSamePlan(item, source.ActivePlanningItem);
                ReviewPlanRow row = BuildPlanRow(
                    item,
                    summary,
                    active ? activeStatus : ReviewStatusCodes.Info,
                    usedIds,
                    rows.Count + 1);
                rows.Add(row);
                if (active && string.IsNullOrWhiteSpace(activePlanKey))
                {
                    activePlanKey = row.PlanKey;
                }
            }

            if (string.IsNullOrWhiteSpace(activePlanKey))
            {
                ReviewPlanRow activeRow = BuildPlanRow(
                    source.ActivePlanningItem,
                    null,
                    activeStatus,
                    usedIds,
                    rows.Count + 1);
                rows.Add(activeRow);
                activePlanKey = activeRow.PlanKey;
            }

            return rows;
        }

        private static ReviewPlanRow BuildPlanRow(
            PlanningItemViewModel item,
            PlanningItemDetailsViewModel summary,
            string status,
            ISet<string> usedIds,
            int index)
        {
            string course = ClinicalReviewValueMapper.SanitizeClinicalLabel(
                item == null
                    ? summary == null ? null : summary.PlanningItemCourseId
                    : item.PlanningItemCourse,
                "Course");
            string plan = ClinicalReviewValueMapper.SanitizeClinicalLabel(
                item == null
                    ? summary == null ? null : summary.PlanName
                    : item.PlanningItemId,
                "Planning item " + index);
            string type = ClinicalReviewValueMapper.SanitizeClinicalLabel(
                item == null
                    ? summary == null ? null : summary.IsPlanSum
                    : item.PlanningItemType,
                "Planning item");

            double? dosePerFraction = null;
            double? totalDose = null;
            int? fractions = null;
            DateTimeOffset? created = null;
            if (item != null)
            {
                created = new DateTimeOffset(item.Creation.ToUniversalTime());
                PlanSetup planSetup = item.PlanningItemObject as PlanSetup;
                if (planSetup != null)
                {
                    dosePerFraction =
                        ConvertDoseToGray(planSetup.DosePerFraction);
                    totalDose = ConvertDoseToGray(planSetup.TotalDose);
                    fractions = planSetup.NumberOfFractions;
                }
            }

            string target =
                ClinicalReviewValueMapper.SanitizeClinicalLabel(
                    summary == null
                        ? item == null ? null : item.PlanningItemTargetId
                        : summary.PlanTarget,
                    "Target not specified");

            return new ReviewPlanRow
            {
                PlanKey = ClinicalReviewValueMapper.CreateUniqueStableId(
                    course + "-" + type + "-" + plan,
                    "planning-item-" + index,
                    usedIds),
                DisplayLabel = course + " · " + plan + " (" + type + ")",
                CreatedUtc = created,
                DosePerFractionGy = dosePerFraction,
                TotalDoseGy = totalDose,
                FractionCount = fractions,
                TargetDisplayLabel = target,
                Status = status
            };
        }

        private static PlanningItemViewModel FindPlanningItem(
            MainViewModel source,
            PlanningItemDetailsViewModel summary)
        {
            return (source.PlanningItemList ??
                    new System.Collections.ObjectModel.ObservableCollection<
                        PlanningItemViewModel>())
                .FirstOrDefault(item =>
                    item != null &&
                    (ReferenceEquals(
                         item.PlanningItemObject,
                         summary.PlanningItemObject) ||
                     string.Equals(
                         item.PlanningItemIdWithCourse,
                         summary.PlanningItemIdWithCourse,
                         StringComparison.Ordinal)));
        }

        private static bool IsSamePlan(
            PlanningItemViewModel left,
            PlanningItemViewModel right)
        {
            return left != null &&
                   right != null &&
                   (ReferenceEquals(
                        left.PlanningItemObject,
                        right.PlanningItemObject) ||
                    (string.Equals(
                         left.PlanningItemCourse,
                         right.PlanningItemCourse,
                         StringComparison.Ordinal) &&
                     string.Equals(
                         left.PlanningItemId,
                         right.PlanningItemId,
                         StringComparison.Ordinal) &&
                     string.Equals(
                         left.PlanningItemType,
                         right.PlanningItemType,
                         StringComparison.Ordinal)));
        }

        private static string GetAggregateStatus(ReviewSnapshot snapshot)
        {
            var statuses = snapshot.PqmRows
                .Select(item => item.Status)
                .Concat(snapshot.PlanCheckRows.Select(item => item.Status))
                .Concat(snapshot.FieldRows.Select(item => item.IdStatus))
                .Concat(snapshot.FieldRows.Select(item => item.NameStatus))
                .Concat(snapshot.StructureMappings.Select(item => item.Status))
                .ToList();

            if (statuses.Contains(ReviewStatusCodes.Fail))
            {
                return ReviewStatusCodes.Fail;
            }

            if (statuses.Contains(ReviewStatusCodes.Variation) ||
                statuses.Contains(ReviewStatusCodes.NotEvaluated))
            {
                return ReviewStatusCodes.Variation;
            }

            return statuses.Count == 0
                ? ReviewStatusCodes.NotEvaluated
                : ReviewStatusCodes.Pass;
        }

        private static string InferPqmUnit(params string[] values)
        {
            foreach (string value in values)
            {
                string unit = ClinicalReviewValueMapper.InferUnit(value);
                if (unit != ReviewUnitCodes.Text)
                {
                    return unit;
                }
            }

            return ReviewUnitCodes.Text;
        }

        private static double? ParsePqmValue(string value, string unit)
        {
            if (!ClinicalReviewValueMapper.TryParseClinicalDouble(
                    value,
                    out double parsed))
            {
                return null;
            }

            if (unit == ReviewUnitCodes.Gray &&
                (value ?? string.Empty).IndexOf(
                    "cGy",
                    StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return ClinicalReviewValueMapper.ConvertDoseToGray(
                    parsed,
                    "cGy");
            }

            return parsed;
        }

        private static string BuildPqmExplanation(PQMSummaryViewModel item)
        {
            string raw = string.Format(
                CultureInfo.InvariantCulture,
                "Goal: {0}; variation: {1}; achieved: {2}.",
                item.Goal ?? "not specified",
                item.Variation ?? "not specified",
                item.Achieved ?? "not evaluated");
            return ClinicalReviewValueMapper.SanitizeClinicalLabel(
                raw,
                "PQM values contain no publishable text.");
        }

        private static double ConvertDoseToGray(DoseValue value)
        {
            if (value.Unit.CompareTo(DoseValue.DoseUnit.cGy) == 0)
            {
                return ClinicalReviewValueMapper.ConvertDoseToGray(
                    value.Dose,
                    "cGy");
            }

            if (value.Unit.CompareTo(DoseValue.DoseUnit.Gy) == 0)
            {
                return ClinicalReviewValueMapper.ConvertDoseToGray(
                    value.Dose,
                    "Gy");
            }

            throw new InvalidOperationException(
                "Only Gy and cGy dose values can be projected.");
        }

        private static double? GetFiniteVolume(Structure structure)
        {
            double volume = structure.Volume;
            return double.IsNaN(volume) ||
                   double.IsInfinity(volume) ||
                   volume < 0.0
                ? (double?)null
                : volume;
        }

        private static string GetDvhRole(
            Structure structure,
            string structureId)
        {
            string normalized = (structureId ?? string.Empty).ToUpperInvariant();
            if (normalized.Contains("PTV") ||
                normalized.Contains("CTV") ||
                normalized.Contains("GTV") ||
                normalized.Contains("ITV"))
            {
                return ReviewDvhRoleCodes.Target;
            }

            if (normalized == "BODY" ||
                normalized.Contains("EXTERNAL") ||
                string.Equals(
                    structure.DicomType,
                    "EXTERNAL",
                    StringComparison.OrdinalIgnoreCase))
            {
                return ReviewDvhRoleCodes.External;
            }

            return ReviewDvhRoleCodes.OrganAtRisk;
        }

        private static string GetColorHex(Structure structure, int index)
        {
            try
            {
                return string.Format(
                    CultureInfo.InvariantCulture,
                    "#{0:X2}{1:X2}{2:X2}",
                    structure.Color.R,
                    structure.Color.G,
                    structure.Color.B);
            }
            catch
            {
                string[] fallback =
                {
                    "#2563EB",
                    "#DC2626",
                    "#059669",
                    "#7C3AED",
                    "#D97706",
                    "#0891B2"
                };
                return fallback[(index - 1) % fallback.Length];
            }
        }

        private static string BuildPlanDisplayLabel(
            PlanningItemViewModel item)
        {
            if (item == null)
            {
                return "Active planning item";
            }

            string course = ClinicalReviewValueMapper.SanitizeClinicalLabel(
                item.PlanningItemCourse,
                "Course");
            string plan = ClinicalReviewValueMapper.SanitizeClinicalLabel(
                item.PlanningItemId,
                "Planning item");
            return course + " · " + plan;
        }
    }
}
