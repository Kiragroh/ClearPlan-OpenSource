using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace ClearPlan.Core.Review
{
    public static class ReviewSnapshotValidationCodes
    {
        public const string SnapshotMissing = "snapshot.missing";
        public const string UnsupportedSchemaVersion = "schema.unsupported";
        public const string SimulatorRequiresSynthetic = "simulator.synthetic.required";
        public const string MissingRequiredValue = "value.required";
        public const string DuplicateStableId = "stable-id.duplicate";
        public const string MissingActivePlan = "active-plan.missing";
        public const string InvalidStatus = "status.invalid";
        public const string InvalidSeverity = "severity.invalid";
        public const string InvalidUnit = "unit.invalid";
        public const string InvalidDose = "dvh.dose.invalid";
        public const string DecreasingDose = "dvh.dose.decreasing";
        public const string InvalidVolume = "dvh.volume.invalid";
        public const string IncreasingCumulativeVolume = "dvh.volume.increasing";
        public const string MissingReferencedStructure = "structure.reference.missing";
        public const string InvalidDvhRole = "dvh.role.invalid";
        public const string InvalidLineStyle = "dvh.line-style.invalid";
        public const string UnsafeDisplayLabel = "display-label.unsafe";
    }

    public sealed class ReviewSnapshotValidationIssue
    {
        public ReviewSnapshotValidationIssue(string code, string path, string message)
        {
            Code = code ?? string.Empty;
            Path = path ?? string.Empty;
            Message = message ?? string.Empty;
        }

        public string Code { get; private set; }

        public string Path { get; private set; }

        public string Message { get; private set; }
    }

    public sealed class ReviewSnapshotValidationResult
    {
        internal ReviewSnapshotValidationResult(
            IList<ReviewSnapshotValidationIssue> issues)
        {
            Issues = new List<ReviewSnapshotValidationIssue>(
                issues ?? new List<ReviewSnapshotValidationIssue>());
        }

        public bool IsValid
        {
            get { return Issues.Count == 0; }
        }

        public IList<ReviewSnapshotValidationIssue> Issues { get; private set; }
    }

    public static class ReviewSnapshotValidator
    {
        private const double DvhTolerance = 1e-9;
        private static readonly DefaultContractResolver PublicContractResolver = new DefaultContractResolver();

        private static readonly HashSet<string> AllowedStatuses =
            new HashSet<string>(
                new[]
                {
                    ReviewStatusCodes.Pass,
                    ReviewStatusCodes.Variation,
                    ReviewStatusCodes.Fail,
                    ReviewStatusCodes.Info,
                    ReviewStatusCodes.NotEvaluated,
                    ReviewStatusCodes.Available,
                    ReviewStatusCodes.Fallback,
                    ReviewStatusCodes.Unavailable,
                    ReviewStatusCodes.NotConfigured
                },
                StringComparer.Ordinal);

        private static readonly HashSet<string> AllowedSeverities =
            new HashSet<string>(
                new[]
                {
                    ReviewSeverityCodes.None,
                    ReviewSeverityCodes.Info,
                    ReviewSeverityCodes.Warning,
                    ReviewSeverityCodes.Error
                },
                StringComparer.Ordinal);

        private static readonly HashSet<string> AllowedUnits =
            new HashSet<string>(
                new[]
                {
                    ReviewUnitCodes.Gray,
                    ReviewUnitCodes.Percent,
                    ReviewUnitCodes.CubicCentimeter,
                    ReviewUnitCodes.Count,
                    ReviewUnitCodes.Text,
                    ReviewUnitCodes.Boolean,
                    ReviewUnitCodes.Degree
                },
                StringComparer.Ordinal);

        private static readonly HashSet<string> AllowedDvhRoles =
            new HashSet<string>(
                new[]
                {
                    ReviewDvhRoleCodes.Target,
                    ReviewDvhRoleCodes.OrganAtRisk,
                    ReviewDvhRoleCodes.External,
                    ReviewDvhRoleCodes.Other
                },
                StringComparer.Ordinal);

        private static readonly HashSet<string> AllowedLineStyles =
            new HashSet<string>(
                new[]
                {
                    ReviewLineStyleCodes.Solid,
                    ReviewLineStyleCodes.Dash,
                    ReviewLineStyleCodes.Dot
                },
                StringComparer.Ordinal);

        private static readonly Regex DicomUidPattern =
            new Regex(
                @"(?<![\d.])\d+(?:\.\d+){4,}(?![\d.])",
                RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex AbsoluteWindowsPathPattern =
            new Regex(
                @"(?<![A-Za-z0-9])(?:[A-Za-z]:[\\/]|\\\\)",
                RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public static ReviewSnapshotValidationResult Validate(
            ReviewSnapshot snapshot)
        {
            return ValidateInternal(snapshot, false);
        }

        public static ReviewSnapshotValidationResult ValidateForSimulator(
            ReviewSnapshot snapshot)
        {
            return ValidateInternal(snapshot, true);
        }

        private static ReviewSnapshotValidationResult ValidateInternal(
            ReviewSnapshot snapshot,
            bool requireSynthetic)
        {
            var issues = new List<ReviewSnapshotValidationIssue>();
            if (snapshot == null)
            {
                Add(
                    issues,
                    ReviewSnapshotValidationCodes.SnapshotMissing,
                    "$",
                    "A review snapshot is required.");
                return new ReviewSnapshotValidationResult(issues);
            }

            if (snapshot.SchemaVersion != ReviewSnapshot.CurrentSchemaVersion)
            {
                Add(
                    issues,
                    ReviewSnapshotValidationCodes.UnsupportedSchemaVersion,
                    "$.schemaVersion",
                    string.Format(
                        "Schema version {0} is unsupported; expected {1}.",
                        snapshot.SchemaVersion,
                        ReviewSnapshot.CurrentSchemaVersion));
            }

            if (requireSynthetic && !snapshot.Synthetic)
            {
                Add(
                    issues,
                    ReviewSnapshotValidationCodes.SimulatorRequiresSynthetic,
                    "$.synthetic",
                    "The simulator accepts synthetic snapshots only.");
            }

            RequireValue(issues, snapshot.ScenarioId, "$.scenarioId");

            IList<ReviewSourceStatus> sources =
                snapshot.Sources ?? new List<ReviewSourceStatus>();
            IList<ReviewPlanRow> plans =
                snapshot.Plans ?? new List<ReviewPlanRow>();
            IList<ReviewPqmRow> pqmRows =
                snapshot.PqmRows ?? new List<ReviewPqmRow>();
            IList<ReviewCheckRow> checks =
                snapshot.PlanCheckRows ?? new List<ReviewCheckRow>();
            IList<ReviewFieldRow> fields =
                snapshot.FieldRows ?? new List<ReviewFieldRow>();
            IList<ReviewStructureMapping> mappings =
                snapshot.StructureMappings ?? new List<ReviewStructureMapping>();
            IList<ReviewDvhSeries> dvhSeries =
                snapshot.DvhSeries ?? new List<ReviewDvhSeries>();

            ValidateStableIds(
                issues,
                sources.Select(item => item == null ? null : item.StableId),
                "$.sources");
            ValidateStableIds(
                issues,
                plans.Select(item => item == null ? null : item.PlanKey),
                "$.plans");
            ValidateStableIds(
                issues,
                pqmRows.Select(item => item == null ? null : item.StableId),
                "$.pqmRows");
            ValidateStableIds(
                issues,
                checks.Select(item => item == null ? null : item.CheckCode),
                "$.planCheckRows");
            ValidateStableIds(
                issues,
                fields.Select(item => item == null ? null : item.StableId),
                "$.fieldRows");
            ValidateStableIds(
                issues,
                mappings.Select(item => item == null ? null : item.StableId),
                "$.structureMappings");
            ValidateStableIds(
                issues,
                dvhSeries.Select(item => item == null ? null : item.StableId),
                "$.dvhSeries");

            if (string.IsNullOrWhiteSpace(snapshot.ActivePlanKey) ||
                !plans.Any(plan =>
                    plan != null &&
                    string.Equals(
                        plan.PlanKey,
                        snapshot.ActivePlanKey,
                        StringComparison.Ordinal)))
            {
                Add(
                    issues,
                    ReviewSnapshotValidationCodes.MissingActivePlan,
                    "$.activePlanKey",
                    "The active plan key must identify one plan row.");
            }

            ValidateSources(issues, sources);
            ValidatePlans(issues, plans);
            ValidatePqmRows(issues, pqmRows);
            ValidateChecks(issues, checks);
            ValidateFields(issues, fields);
            ValidateMappings(issues, mappings);
            ValidateDvhSeries(issues, dvhSeries);
            ValidateStructureReferences(issues, mappings, dvhSeries);
            ValidatePublishSafeText(issues, snapshot);
            ValidateExtendedReview(issues, snapshot);

            return new ReviewSnapshotValidationResult(issues);
        }

        private static void ValidateSources(
            IList<ReviewSnapshotValidationIssue> issues,
            IList<ReviewSourceStatus> sources)
        {
            for (int index = 0; index < sources.Count; index++)
            {
                ReviewSourceStatus source = sources[index];
                if (source == null)
                {
                    RequireValue(issues, null, Path("$.sources", index));
                    continue;
                }

                ValidateStatus(
                    issues,
                    source.Status,
                    Path("$.sources", index) + ".status");
            }
        }

        private static void ValidatePlans(
            IList<ReviewSnapshotValidationIssue> issues,
            IList<ReviewPlanRow> plans)
        {
            for (int index = 0; index < plans.Count; index++)
            {
                ReviewPlanRow plan = plans[index];
                if (plan == null)
                {
                    RequireValue(issues, null, Path("$.plans", index));
                    continue;
                }

                ValidateStatus(
                    issues,
                    plan.Status,
                    Path("$.plans", index) + ".status");
                ValidateOptionalNonNegativeFinite(
                    issues,
                    plan.DosePerFractionGy,
                    Path("$.plans", index) + ".dosePerFractionGy");
                ValidateOptionalNonNegativeFinite(
                    issues,
                    plan.TotalDoseGy,
                    Path("$.plans", index) + ".totalDoseGy");
            }
        }

        private static void ValidatePqmRows(
            IList<ReviewSnapshotValidationIssue> issues,
            IList<ReviewPqmRow> rows)
        {
            for (int index = 0; index < rows.Count; index++)
            {
                ReviewPqmRow row = rows[index];
                if (row == null)
                {
                    RequireValue(issues, null, Path("$.pqmRows", index));
                    continue;
                }

                string root = Path("$.pqmRows", index);
                ValidateStatus(issues, row.Status, root + ".status");
                ValidateSeverity(issues, row.Severity, root + ".severity");
                ValidateUnit(issues, row.Unit, root + ".unit");
                ValidateOptionalFinite(issues, row.Goal, root + ".goal");
                ValidateOptionalFinite(issues, row.Variation, root + ".variation");
                ValidateOptionalFinite(
                    issues,
                    row.AchievedValue,
                    root + ".achievedValue");
            }
        }

        private static void ValidateChecks(
            IList<ReviewSnapshotValidationIssue> issues,
            IList<ReviewCheckRow> rows)
        {
            for (int index = 0; index < rows.Count; index++)
            {
                ReviewCheckRow row = rows[index];
                if (row == null)
                {
                    RequireValue(issues, null, Path("$.planCheckRows", index));
                    continue;
                }

                string root = Path("$.planCheckRows", index);
                ValidateStatus(issues, row.Status, root + ".status");
                ValidateSeverity(issues, row.Severity, root + ".severity");
                ValidateUnit(issues, row.Unit, root + ".unit");
            }
        }

        private static void ValidateFields(
            IList<ReviewSnapshotValidationIssue> issues,
            IList<ReviewFieldRow> rows)
        {
            for (int index = 0; index < rows.Count; index++)
            {
                ReviewFieldRow row = rows[index];
                if (row == null)
                {
                    RequireValue(issues, null, Path("$.fieldRows", index));
                    continue;
                }

                string root = Path("$.fieldRows", index);
                ValidateStatus(issues, row.IdStatus, root + ".idStatus");
                ValidateStatus(issues, row.NameStatus, root + ".nameStatus");
            }
        }

        private static void ValidateMappings(
            IList<ReviewSnapshotValidationIssue> issues,
            IList<ReviewStructureMapping> mappings)
        {
            for (int index = 0; index < mappings.Count; index++)
            {
                ReviewStructureMapping mapping = mappings[index];
                if (mapping == null)
                {
                    RequireValue(issues, null, Path("$.structureMappings", index));
                    continue;
                }

                ValidateStatus(
                    issues,
                    mapping.Status,
                    Path("$.structureMappings", index) + ".status");
            }
        }

        private static void ValidateDvhSeries(
            IList<ReviewSnapshotValidationIssue> issues,
            IList<ReviewDvhSeries> series)
        {
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                ReviewDvhSeries item = series[seriesIndex];
                string root = Path("$.dvhSeries", seriesIndex);
                if (item == null)
                {
                    RequireValue(issues, null, root);
                    continue;
                }

                ValidateToken(
                    issues,
                    item.Role,
                    AllowedDvhRoles,
                    ReviewSnapshotValidationCodes.InvalidDvhRole,
                    root + ".role",
                    "DVH role");
                ValidateToken(
                    issues,
                    item.LineStyle,
                    AllowedLineStyles,
                    ReviewSnapshotValidationCodes.InvalidLineStyle,
                    root + ".lineStyle",
                    "line style");
                ValidateOptionalNonNegativeFinite(
                    issues,
                    item.VolumeCc,
                    root + ".volumeCc");

                IList<ReviewDvhPoint> points =
                    item.Points ?? new List<ReviewDvhPoint>();
                double? previousDose = null;
                double? previousVolume = null;
                for (int pointIndex = 0; pointIndex < points.Count; pointIndex++)
                {
                    ReviewDvhPoint point = points[pointIndex];
                    string pointRoot = Path(root + ".points", pointIndex);
                    if (point == null)
                    {
                        RequireValue(issues, null, pointRoot);
                        continue;
                    }

                    if (!IsFinite(point.DoseGy) || point.DoseGy < 0.0)
                    {
                        Add(
                            issues,
                            ReviewSnapshotValidationCodes.InvalidDose,
                            pointRoot + ".doseGy",
                            "DVH dose must be finite and non-negative.");
                    }
                    else if (previousDose.HasValue &&
                             point.DoseGy + DvhTolerance < previousDose.Value)
                    {
                        Add(
                            issues,
                            ReviewSnapshotValidationCodes.DecreasingDose,
                            pointRoot + ".doseGy",
                            "DVH dose coordinates must be ordered increasingly.");
                    }

                    if (!IsFinite(point.VolumePercent) ||
                        point.VolumePercent < 0.0 ||
                        point.VolumePercent > 100.0)
                    {
                        Add(
                            issues,
                            ReviewSnapshotValidationCodes.InvalidVolume,
                            pointRoot + ".volumePercent",
                            "Cumulative volume must be finite and between 0 and 100 percent.");
                    }
                    else if (previousVolume.HasValue &&
                             point.VolumePercent >
                             previousVolume.Value + DvhTolerance)
                    {
                        Add(
                            issues,
                            ReviewSnapshotValidationCodes.IncreasingCumulativeVolume,
                            pointRoot + ".volumePercent",
                            "Cumulative volume must not increase with dose.");
                    }

                    previousDose = IsFinite(point.DoseGy)
                        ? (double?)point.DoseGy
                        : null;
                    previousVolume = IsFinite(point.VolumePercent)
                        ? (double?)point.VolumePercent
                        : null;
                }
            }
        }

        private static void ValidateStructureReferences(
            IList<ReviewSnapshotValidationIssue> issues,
            IList<ReviewStructureMapping> mappings,
            IList<ReviewDvhSeries> series)
        {
            var structureIds = new HashSet<string>(
                series
                    .Where(item =>
                        item != null &&
                        !string.IsNullOrWhiteSpace(item.StructureId))
                    .Select(item => item.StructureId),
                StringComparer.Ordinal);

            for (int index = 0; index < mappings.Count; index++)
            {
                ReviewStructureMapping mapping = mappings[index];
                if (mapping == null ||
                    string.IsNullOrWhiteSpace(mapping.SelectedStructureId))
                {
                    continue;
                }

                if (!structureIds.Contains(mapping.SelectedStructureId))
                {
                    Add(
                        issues,
                        ReviewSnapshotValidationCodes.MissingReferencedStructure,
                        Path("$.structureMappings", index) + ".selectedStructureId",
                        "The selected structure must identify a DVH series.");
                }
            }
        }

        private static void ValidatePublishSafeText(
            IList<ReviewSnapshotValidationIssue> issues,
            ReviewSnapshot snapshot)
        {
            ValidatePublishSafeTextValue(issues, snapshot, "$");
        }

        private static void ValidatePublishSafeTextValue(
            IList<ReviewSnapshotValidationIssue> issues,
            object value,
            string path)
        {
            if (value == null)
            {
                return;
            }
            if (value is byte[]) return;

            string text = value as string;
            if (text != null)
            {
                ValidateTextContent(issues, text, path);
                return;
            }

            IEnumerable sequence = value as IEnumerable;
            if (sequence != null)
            {
                int index = 0;
                foreach (object item in sequence)
                {
                    ValidatePublishSafeTextValue(
                        issues,
                        item,
                        Path(path, index));
                    index++;
                }

                return;
            }

            Type valueType = value.GetType();
            if (valueType.Namespace != typeof(ReviewSnapshot).Namespace &&
                valueType.Namespace != typeof(PlanAnalysis.ReviewPlanAnalysis).Namespace)
            {
                return;
            }

            // Follow the actual JSON contract, including default opt-out DTOs,
            // but never inspect deliberately ignored patient geometry/identifiers.
            var contract = PublicContractResolver.ResolveContract(valueType) as JsonObjectContract;
            if (contract == null) return;
            var serializedProperties = contract.Properties.Where(item => item.Readable && !item.Ignored);

            foreach (var item in serializedProperties)
            {
                ValidatePublishSafeTextValue(
                    issues,
                    item.ValueProvider.GetValue(value),
                    path + "." + item.PropertyName);
            }
        }

        private static void ValidateExtendedReview(IList<ReviewSnapshotValidationIssue> issues, ReviewSnapshot snapshot)
        {
            var analysis = snapshot.PlanAnalysis;
            if (analysis != null)
            {
                foreach (double? value in new[] { analysis.TotalMetersetMu, analysis.MuPerGy, analysis.Pam,
                    analysis.MeanApertureAreaCm2, analysis.SmallApertureFraction, analysis.PlanNormalizationPercent })
                    if (value.HasValue && (!ReviewComparison.Finite(value) || value < 0))
                        Add(issues, "analysis.numeric.invalid", "$.planAnalysis", "Plan metrics must be nonnegative finite numbers or unavailable.");
                if (analysis.Pam > 1 || analysis.SmallApertureFraction > 1)
                    Add(issues, "analysis.fraction.invalid", "$.planAnalysis", "Dimensionless modulation fractions must be in [0,1].");
                foreach (var beam in analysis.Beams ?? new List<PlanAnalysis.ReviewBeamAnalysis>())
                foreach (var cp in beam.ControlPoints ?? new List<PlanAnalysis.ReviewControlPointSample>())
                    if (cp.BevImage != null && cp.BevImage.Synthetic != snapshot.Synthetic)
                        Add(issues, "bev.mode.mismatch", "$.planAnalysis", "Clinical and synthetic BEV pixels must remain separate.");
            }
            foreach (var image in snapshot.PlanImages ?? new List<ReviewPlanImage>())
            {
                if (image == null) continue;
                if (image.DosePlane != null && !image.DosePlane.IsValid)
                    Add(issues, "image.doseplane.invalid", "$.planImages", "Detached dose samples require a bounded grid, finite nonnegative Gy or unavailable NaN, and positive prescription Gy.");
                if (snapshot.Synthetic && !image.Synthetic)
                    Add(issues, "image.synthetic.required", "$.planImages", "A synthetic snapshot cannot contain a clinical image.");
                if (image.GrayscalePixels == null) continue;
                if (image.WidthPixels <= 0 || image.HeightPixels <= 0 ||
                    (long)image.WidthPixels * image.HeightPixels != image.GrayscalePixels.Length || image.GrayscalePixels.Length > 4194304 ||
                    !ReviewComparison.Finite(image.PixelSpacingXMillimeters) || image.PixelSpacingXMillimeters <= 0 ||
                    !ReviewComparison.Finite(image.PixelSpacingYMillimeters) || image.PixelSpacingYMillimeters <= 0)
                    Add(issues, "image.geometry.invalid", "$.planImages", "Overview pixels require a bounded complete grid and positive physical spacing.");
            }
            if (snapshot.IsodoseDisplay != null)
            {
                try { snapshot.IsodoseDisplay.Validate(); }
                catch (FormatException) { Add(issues, "image.palette.invalid", "$.isodoseDisplay", "Invalid isodose display configuration."); }
            }
        }

        private static void ValidateTextContent(
            IList<ReviewSnapshotValidationIssue> issues,
            string value,
            string path)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            string trimmed = value.Trim();
            bool looksLikePath =
                trimmed.StartsWith("/", StringComparison.Ordinal) ||
                AbsoluteWindowsPathPattern.IsMatch(value);
            if (looksLikePath || DicomUidPattern.IsMatch(trimmed))
            {
                Add(
                    issues,
                    ReviewSnapshotValidationCodes.UnsafeDisplayLabel,
                    path,
                    "Snapshot text must not contain private paths or DICOM UIDs.");
            }
        }

        private static void ValidateStableIds(
            IList<ReviewSnapshotValidationIssue> issues,
            IEnumerable<string> stableIds,
            string root)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            int index = 0;
            foreach (string stableId in stableIds)
            {
                string path = Path(root, index);
                if (string.IsNullOrWhiteSpace(stableId))
                {
                    RequireValue(issues, stableId, path);
                }
                else if (!seen.Add(stableId))
                {
                    Add(
                        issues,
                        ReviewSnapshotValidationCodes.DuplicateStableId,
                        path,
                        string.Format("Stable ID '{0}' occurs more than once.", stableId));
                }

                index++;
            }
        }

        private static void ValidateStatus(
            IList<ReviewSnapshotValidationIssue> issues,
            string status,
            string path)
        {
            ValidateToken(
                issues,
                status,
                AllowedStatuses,
                ReviewSnapshotValidationCodes.InvalidStatus,
                path,
                "status");
        }

        private static void ValidateSeverity(
            IList<ReviewSnapshotValidationIssue> issues,
            string severity,
            string path)
        {
            ValidateToken(
                issues,
                severity,
                AllowedSeverities,
                ReviewSnapshotValidationCodes.InvalidSeverity,
                path,
                "severity");
        }

        private static void ValidateUnit(
            IList<ReviewSnapshotValidationIssue> issues,
            string unit,
            string path)
        {
            ValidateToken(
                issues,
                unit,
                AllowedUnits,
                ReviewSnapshotValidationCodes.InvalidUnit,
                path,
                "unit");
        }

        private static void ValidateToken(
            IList<ReviewSnapshotValidationIssue> issues,
            string value,
            ISet<string> allowed,
            string code,
            string path,
            string label)
        {
            if (!string.IsNullOrWhiteSpace(value) && !allowed.Contains(value))
            {
                Add(
                    issues,
                    code,
                    path,
                    string.Format("Unsupported {0} token '{1}'.", label, value));
            }
        }

        private static void ValidateOptionalFinite(
            IList<ReviewSnapshotValidationIssue> issues,
            double? value,
            string path)
        {
            if (value.HasValue && !IsFinite(value.Value))
            {
                Add(
                    issues,
                    ReviewSnapshotValidationCodes.InvalidDose,
                    path,
                    "Numeric values must be finite.");
            }
        }

        private static void ValidateOptionalNonNegativeFinite(
            IList<ReviewSnapshotValidationIssue> issues,
            double? value,
            string path)
        {
            if (value.HasValue &&
                (!IsFinite(value.Value) || value.Value < 0.0))
            {
                Add(
                    issues,
                    ReviewSnapshotValidationCodes.InvalidDose,
                    path,
                    "Numeric values must be finite and non-negative.");
            }
        }

        private static bool IsFinite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static void RequireValue(
            IList<ReviewSnapshotValidationIssue> issues,
            string value,
            string path)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                Add(
                    issues,
                    ReviewSnapshotValidationCodes.MissingRequiredValue,
                    path,
                    "A stable non-empty value is required.");
            }
        }

        private static string Path(string root, int index)
        {
            return string.Format("{0}[{1}]", root, index);
        }

        private static void Add(
            IList<ReviewSnapshotValidationIssue> issues,
            string code,
            string path,
            string message)
        {
            issues.Add(new ReviewSnapshotValidationIssue(code, path, message));
        }
    }
}
