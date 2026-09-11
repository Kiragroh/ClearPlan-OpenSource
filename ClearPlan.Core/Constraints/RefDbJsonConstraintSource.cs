using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace ClearPlan.Core.Constraints
{
    public sealed class RefDbJsonConstraintSource : IConstraintCatalogSource
    {
        public const string SupportedSchema = "RSAlign.local_refdb_constraints.v1";

        private readonly bool _includeInactive;

        public RefDbJsonConstraintSource(bool includeInactive = false)
        {
            _includeInactive = includeInactive;
        }

        public ConstraintCatalog Load(string path)
        {
            var catalog = new ConstraintCatalog
            {
                SourceKind = "RefDB",
                SourcePath = path
            };

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                catalog.Issues.Add(Error(
                    "source_path",
                    path,
                    "RefDB JSON file does not exist.",
                    "RefDB"));
                return catalog;
            }

            RefDbRootDto source;
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream, Encoding.UTF8, true))
                using (var jsonReader = new JsonTextReader(reader))
                {
                    source = JsonSerializer.CreateDefault().Deserialize<RefDbRootDto>(jsonReader);
                }
            }
            catch (Exception exception)
            {
                catalog.Issues.Add(Error(
                    "json",
                    null,
                    "RefDB JSON could not be read: " + exception.Message,
                    "RefDB"));
                return catalog;
            }

            if (source == null)
            {
                catalog.Issues.Add(Error("json", null, "RefDB JSON is empty.", "RefDB"));
                return catalog;
            }

            catalog.Schema = source.Schema;
            if (!string.Equals(source.Schema, SupportedSchema, StringComparison.Ordinal))
            {
                catalog.Issues.Add(Error(
                    "schema",
                    source.Schema,
                    "Unsupported RefDB schema. Expected " + SupportedSchema + ".",
                    "RefDB"));
                return catalog;
            }

            List<RefDbStructureDto> structures = source.Structures ?? new List<RefDbStructureDto>();
            List<RefDbTableDto> tables = source.Tables ?? new List<RefDbTableDto>();
            Dictionary<string, RefDbTableDto> details =
                source.Details ?? new Dictionary<string, RefDbTableDto>();

            catalog.Statistics.TotalStructures = structures.Count;
            catalog.Statistics.ActiveStructures = structures.Count(IsActive);
            catalog.Statistics.TotalTables = tables.Count;
            catalog.Statistics.ActiveTables = tables.Count(IsActive);
            catalog.Statistics.TotalConstraints = details.Values.Sum(
                detail => detail.Constraints == null ? 0 : detail.Constraints.Count);

            foreach (RefDbStructureDto structure in structures
                .Where(item => _includeInactive || IsActive(item)))
            {
                catalog.Structures.Add(MapStructure(structure));
            }

            var detailsById = details.Values
                .Where(detail => detail != null && detail.Id.HasValue)
                .GroupBy(detail => detail.Id.Value)
                .ToDictionary(group => group.Key, group => group.First());

            foreach (RefDbTableDto tableSummary in tables
                .Where(item => _includeInactive || IsActive(item)))
            {
                RefDbTableDto detail;
                if (!tableSummary.Id.HasValue ||
                    !detailsById.TryGetValue(tableSummary.Id.Value, out detail))
                {
                    detail = tableSummary;
                    catalog.Issues.Add(Warning(
                        "details",
                        tableSummary.Id.HasValue
                            ? tableSummary.Id.Value.ToString(CultureInfo.InvariantCulture)
                            : null,
                        "RefDB table has no matching details entry.",
                        "Tables"));
                }

                catalog.Tables.Add(MapTable(tableSummary, detail, catalog.Issues));
            }

            foreach (CatalogValidationIssue issue in ConstraintCatalogValidator.Validate(catalog))
            {
                catalog.Issues.Add(issue);
            }

            return catalog;
        }

        private static StructureDefinition MapStructure(RefDbStructureDto source)
        {
            var structure = new StructureDefinition
            {
                StructureId = source.Id.HasValue
                    ? source.Id.Value.ToString(CultureInfo.InvariantCulture)
                    : null,
                CanonicalName = source.CanonicalName,
                Active = IsActive(source),
                Laterality = source.Laterality
            };

            foreach (string alias in (source.Aliases ?? new List<RefDbAliasDto>())
                .Select(item => item == null ? null : item.Alias)
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                structure.Aliases.Add(alias.Trim());
            }

            if (source.SideAliases != null)
            {
                CopyDistinct(source.SideAliases, "L", structure.SideAliasesLeft);
                CopyDistinct(source.SideAliases, "R", structure.SideAliasesRight);
            }

            return structure;
        }

        private static ConstraintTableDefinition MapTable(
            RefDbTableDto summary,
            RefDbTableDto detail,
            IList<CatalogValidationIssue> issues)
        {
            string tableId = summary.Id.HasValue
                ? summary.Id.Value.ToString(CultureInfo.InvariantCulture)
                : null;
            var table = new ConstraintTableDefinition
            {
                TableId = tableId,
                DisplayName = FirstNonBlank(detail.Name, summary.Name, tableId),
                Active = IsActive(summary),
                IsPlanSum = ContainsPlanSum(detail.Name) || ContainsPlanSum(summary.Name),
                Site = FirstNonBlank(detail.Site, summary.Site),
                Regime = FirstNonBlank(detail.Regime, summary.Regime),
                Source = FirstNonBlank(detail.SourceNote, summary.SourceNote)
            };

            TryAssignInteger(
                FirstPopulated(detail.FractionCountMinimum, summary.FractionCountMinimum),
                value => table.FractionCountMinimum = value,
                issues,
                tableId,
                "fx_min");
            TryAssignInteger(
                FirstPopulated(detail.FractionCountMaximum, summary.FractionCountMaximum),
                value => table.FractionCountMaximum = value,
                issues,
                tableId,
                "fx_max");
            TryAssignDecimal(
                FirstPopulated(detail.DosePerFractionMinimum, summary.DosePerFractionMinimum),
                value => table.DosePerFractionMinimumGy = value,
                issues,
                tableId,
                "dpf_min");
            TryAssignDecimal(
                FirstPopulated(detail.DosePerFractionMaximum, summary.DosePerFractionMaximum),
                value => table.DosePerFractionMaximumGy = value,
                issues,
                tableId,
                "dpf_max");
            TryAssignDecimal(
                FirstPopulated(detail.TotalDoseMinimum, summary.TotalDoseMinimum),
                value => table.TotalDoseMinimumGy = value,
                issues,
                tableId,
                "td_min");
            TryAssignDecimal(
                FirstPopulated(detail.TotalDoseMaximum, summary.TotalDoseMaximum),
                value => table.TotalDoseMaximumGy = value,
                issues,
                tableId,
                "td_max");

            foreach (string prescription in (detail.Prescriptions ?? summary.Prescriptions ?? new List<RefDbPrescriptionDto>())
                .Where(item => item != null &&
                               (_includeInactivePrescription(item) ||
                                string.Equals(item.Status, "active", StringComparison.OrdinalIgnoreCase)))
                .Select(item => item.Name)
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                table.PrescriptionLabels.Add(prescription.Trim());
            }

            int row = 0;
            foreach (RefDbConstraintDto constraint in detail.Constraints ?? new List<RefDbConstraintDto>())
            {
                row++;
                ConstraintDefinition mapped = MapConstraint(constraint, tableId, row, issues);
                if (mapped != null)
                {
                    table.Constraints.Add(mapped);
                }
            }

            return table;
        }

        private static object FirstPopulated(object detail, object summary)
        {
            // RefDB exports may keep selection metadata in the summary only.
            // Nonempty detail values remain authoritative, including numeric zero.
            return detail == null || string.IsNullOrWhiteSpace(Convert.ToString(detail, CultureInfo.InvariantCulture))
                ? summary : detail;
        }

        private static bool _includeInactivePrescription(RefDbPrescriptionDto item)
        {
            return string.IsNullOrWhiteSpace(item.Status);
        }

        private static ConstraintDefinition MapConstraint(
            RefDbConstraintDto source,
            string tableId,
            int row,
            IList<CatalogValidationIssue> issues)
        {
            string location = string.Format(
                CultureInfo.InvariantCulture,
                "details/{0}/constraints/{1}",
                tableId ?? "?",
                row);
            try
            {
                string metric = ConstraintValueNormalizer.NormalizeMetric(source.Metric);
                string unit = ConstraintValueNormalizer.NormalizeUnit(source.Unit);
                string comparator = ConstraintValueNormalizer.NormalizeComparator(source.Comparator);
                decimal? goal = ConstraintValueNormalizer.ParseNullableDecimal(source.LimitOptimal);
                decimal? variation = ConstraintValueNormalizer.ParseNullableDecimal(source.LimitMaximal);

                var constraint = new ConstraintDefinition
                {
                    ConstraintId = source.Id.HasValue
                        ? source.Id.Value.ToString(CultureInfo.InvariantCulture)
                        : string.Format(CultureInfo.InvariantCulture, "{0}-{1}", tableId, row),
                    TableId = source.TableId.HasValue
                        ? source.TableId.Value.ToString(CultureInfo.InvariantCulture)
                        : tableId,
                    StructureId = source.StructureId.HasValue
                        ? source.StructureId.Value.ToString(CultureInfo.InvariantCulture)
                        : null,
                    StructureName = source.StructureName,
                    RawStructureName = source.RawStructureName,
                    Metric = metric,
                    Unit = unit,
                    Comparator = comparator,
                    ExpectedDirection = DirectionFromComparator(comparator),
                    Goal = goal,
                    Variation = variation,
                    Priority = ParsePriority(source.Priority),
                    Source = source.Source,
                    Comment = source.Comment,
                    DvhObjective = BuildDvhObjective(metric, unit),
                    EvaluationPoint = BuildEvaluationPoint(comparator, goal)
                };

                constraint.RawValues["metric"] = source.Metric;
                constraint.RawValues["unit"] = source.Unit;
                constraint.RawValues["comparator"] = source.Comparator;
                constraint.RawValues["limit_optimal"] = Convert.ToString(
                    source.LimitOptimal,
                    CultureInfo.InvariantCulture);
                constraint.RawValues["limit_maximal"] = Convert.ToString(
                    source.LimitMaximal,
                    CultureInfo.InvariantCulture);
                constraint.RawValues["priority"] = source.Priority;
                return constraint;
            }
            catch (Exception exception)
            {
                issues.Add(Error(
                    "constraint",
                    source == null ? null : source.Metric,
                    "Constraint row could not be normalized: " + exception.Message,
                    location));
                return null;
            }
        }

        private static int? ParsePriority(string value)
        {
            Match match = Regex.Match(value ?? string.Empty, @"^\s*(\d+)");
            int priority;
            return match.Success &&
                   int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out priority)
                ? (int?)priority
                : null;
        }

        private static ConstraintDirection DirectionFromComparator(string comparator)
        {
            if (comparator == "<" || comparator == "<=")
            {
                return ConstraintDirection.LowerIsBetter;
            }

            if (comparator == ">" || comparator == ">=")
            {
                return ConstraintDirection.HigherIsBetter;
            }

            return ConstraintDirection.Unknown;
        }

        private static string BuildDvhObjective(string metric, string unit)
        {
            string objective = metric ?? string.Empty;
            if (objective.Equals("Dmean", StringComparison.OrdinalIgnoreCase))
            {
                objective = "Mean";
            }
            else if (objective.Equals("Dmax", StringComparison.OrdinalIgnoreCase))
            {
                objective = "Max";
            }
            else if (objective.Equals("Dmin", StringComparison.OrdinalIgnoreCase))
            {
                objective = "Min";
            }

            return string.IsNullOrWhiteSpace(unit)
                ? objective
                : string.Format("{0}[{1}]", objective, unit);
        }

        private static string BuildEvaluationPoint(string comparator, decimal? goal)
        {
            if (!goal.HasValue)
            {
                return comparator ?? string.Empty;
            }

            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}{1:0.###}",
                comparator,
                goal.Value);
        }

        private static void TryAssignDecimal(
            object rawValue,
            Action<decimal?> assign,
            IList<CatalogValidationIssue> issues,
            string tableId,
            string field)
        {
            try
            {
                assign(ConstraintValueNormalizer.ParseNullableDecimal(rawValue));
            }
            catch (FormatException exception)
            {
                issues.Add(Warning(
                    field,
                    Convert.ToString(rawValue, CultureInfo.InvariantCulture),
                    exception.Message,
                    "Tables/" + tableId));
            }
        }

        private static void TryAssignInteger(
            object rawValue,
            Action<int?> assign,
            IList<CatalogValidationIssue> issues,
            string tableId,
            string field)
        {
            try
            {
                assign(ConstraintValueNormalizer.ParseNullableInteger(rawValue));
            }
            catch (FormatException exception)
            {
                issues.Add(Warning(
                    field,
                    Convert.ToString(rawValue, CultureInfo.InvariantCulture),
                    exception.Message,
                    "Tables/" + tableId));
            }
        }

        private static void CopyDistinct(
            IDictionary<string, List<string>> aliases,
            string side,
            IList<string> destination)
        {
            List<string> values;
            if (!aliases.TryGetValue(side, out values) || values == null)
            {
                return;
            }

            foreach (string value in values
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                destination.Add(value.Trim());
            }
        }

        private static bool IsActive(RefDbStructureDto item)
        {
            return item != null &&
                   string.Equals(item.Status, "active", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsActive(RefDbTableDto item)
        {
            return item != null &&
                   string.Equals(item.Status, "active", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ContainsPlanSum(string value)
        {
            return !string.IsNullOrWhiteSpace(value) &&
                   (value.IndexOf("plansum", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    value.IndexOf("sum", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static string FirstNonBlank(params string[] values)
        {
            return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        }

        private static CatalogValidationIssue Error(
            string field,
            string rawValue,
            string message,
            string location)
        {
            return new CatalogValidationIssue
            {
                Severity = CatalogIssueSeverity.Error,
                Field = field,
                RawValue = rawValue,
                Message = message,
                SourceLocation = location
            };
        }

        private static CatalogValidationIssue Warning(
            string field,
            string rawValue,
            string message,
            string location)
        {
            return new CatalogValidationIssue
            {
                Severity = CatalogIssueSeverity.Warning,
                Field = field,
                RawValue = rawValue,
                Message = message,
                SourceLocation = location
            };
        }
    }
}
