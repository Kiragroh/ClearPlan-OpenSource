using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ClearPlan.Core.Constraints
{
    public sealed class ExcelConstraintSource : IConstraintCatalogSource
    {
        public const string WorkbookSchema = "ClearPlan.constraint_catalog.xlsx.v1";

        private static readonly string[] TableHeaders =
        {
            "table_id", "display_name", "active", "is_plan_sum", "fx_min", "fx_max",
            "dpf_min_gy", "dpf_max_gy", "total_dose_min_gy", "total_dose_max_gy",
            "site", "regime", "source"
        };

        private static readonly string[] ConstraintHeaders =
        {
            "constraint_id", "table_id", "structure_id", "structure_name", "metric",
            "unit", "comparator", "goal", "variation", "priority", "source", "comment"
        };

        private static readonly string[] StructureHeaders =
        {
            "structure_id", "canonical_name", "active", "laterality", "aliases",
            "side_aliases_left", "side_aliases_right", "dicom_type", "codes"
        };

        private readonly bool _includeInactive;

        public ExcelConstraintSource(bool includeInactive = false)
        {
            _includeInactive = includeInactive;
        }

        public ConstraintCatalog Load(string path)
        {
            var catalog = new ConstraintCatalog
            {
                Schema = WorkbookSchema,
                SourceKind = "Excel",
                SourcePath = path
            };
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                catalog.Issues.Add(Error(
                    "source_path",
                    path,
                    "Excel constraint workbook does not exist.",
                    "Excel"));
                return catalog;
            }

            XlsxWorkbookData workbook;
            try
            {
                workbook = XlsxWorksheetReader.Read(path);
            }
            catch (Exception exception)
            {
                catalog.Issues.Add(Error(
                    "xlsx",
                    null,
                    "Excel constraint workbook could not be read: " + exception.Message,
                    "Excel"));
                return catalog;
            }

            XlsxWorksheetData tablesSheet;
            XlsxWorksheetData constraintsSheet;
            XlsxWorksheetData structuresSheet;
            if (!TryGetSheet(workbook, "Tables", catalog.Issues, out tablesSheet) |
                !TryGetSheet(workbook, "Constraints", catalog.Issues, out constraintsSheet) |
                !TryGetSheet(workbook, "Structures", catalog.Issues, out structuresSheet))
            {
                return catalog;
            }

            SheetRows tableRows = ReadRows(tablesSheet, TableHeaders, catalog.Issues);
            SheetRows constraintRows = ReadRows(
                constraintsSheet,
                ConstraintHeaders,
                catalog.Issues);
            SheetRows structureRows = ReadRows(
                structuresSheet,
                StructureHeaders,
                catalog.Issues);
            if (catalog.Issues.Any(issue => issue.IsFatal))
            {
                return catalog;
            }

            catalog.Statistics.TotalTables = tableRows.Rows.Count;
            catalog.Statistics.ActiveTables = tableRows.Rows.Count(
                row => ParseBoolean(tableRows.Get(row, "active"), true));
            catalog.Statistics.TotalConstraints = constraintRows.Rows.Count;
            catalog.Statistics.TotalStructures = structureRows.Rows.Count;
            catalog.Statistics.ActiveStructures = structureRows.Rows.Count(
                row => ParseBoolean(structureRows.Get(row, "active"), true));

            ReportDuplicates(tableRows, "table_id", catalog.Issues);
            ReportDuplicates(structureRows, "structure_id", catalog.Issues);
            ReportDuplicates(constraintRows, "constraint_id", catalog.Issues);

            foreach (XlsxRowData row in structureRows.Rows)
            {
                bool active = ParseBoolean(structureRows.Get(row, "active"), true);
                if (!_includeInactive && !active)
                {
                    continue;
                }

                var structure = new StructureDefinition
                {
                    StructureId = structureRows.Get(row, "structure_id"),
                    CanonicalName = structureRows.Get(row, "canonical_name"),
                    Active = active,
                    Laterality = structureRows.Get(row, "laterality")
                };
                AddDelimited(structure.Aliases, structureRows.Get(row, "aliases"));
                AddDelimited(structure.SideAliasesLeft, structureRows.Get(row, "side_aliases_left"));
                AddDelimited(structure.SideAliasesRight, structureRows.Get(row, "side_aliases_right"));
                AddDelimited(structure.DicomTypes, structureRows.Get(row, "dicom_type"));
                AddDelimited(structure.Codes, structureRows.Get(row, "codes"));
                catalog.Structures.Add(structure);
            }

            var tablesById = new Dictionary<string, ConstraintTableDefinition>(
                StringComparer.OrdinalIgnoreCase);
            var allTableIds = new HashSet<string>(
                tableRows.Rows.Select(row => tableRows.Get(row, "table_id")),
                StringComparer.OrdinalIgnoreCase);
            foreach (XlsxRowData row in tableRows.Rows)
            {
                string tableId = tableRows.Get(row, "table_id");
                bool active = ParseBoolean(tableRows.Get(row, "active"), true);
                if (!_includeInactive && !active)
                {
                    continue;
                }

                var table = new ConstraintTableDefinition
                {
                    TableId = tableId,
                    DisplayName = tableRows.Get(row, "display_name"),
                    Active = active,
                    IsPlanSum = ParseBoolean(tableRows.Get(row, "is_plan_sum"), false),
                    RequiresConfirmation = ParseBoolean(tableRows.Get(row, "requires_confirmation"), false),
                    Site = tableRows.Get(row, "site"),
                    Regime = tableRows.Get(row, "regime"),
                    Source = tableRows.Get(row, "source")
                };
                AssignInteger(tableRows, row, "fx_min", value => table.FractionCountMinimum = value, catalog.Issues);
                AssignInteger(tableRows, row, "fx_max", value => table.FractionCountMaximum = value, catalog.Issues);
                AssignDecimal(tableRows, row, "dpf_min_gy", value => table.DosePerFractionMinimumGy = value, catalog.Issues);
                AssignDecimal(tableRows, row, "dpf_max_gy", value => table.DosePerFractionMaximumGy = value, catalog.Issues);
                AssignDecimal(tableRows, row, "total_dose_min_gy", value => table.TotalDoseMinimumGy = value, catalog.Issues);
                AssignDecimal(tableRows, row, "total_dose_max_gy", value => table.TotalDoseMaximumGy = value, catalog.Issues);
                if (!tablesById.ContainsKey(tableId))
                {
                    tablesById.Add(tableId, table);
                    catalog.Tables.Add(table);
                }
            }

            foreach (XlsxRowData row in constraintRows.Rows)
            {
                string tableId = constraintRows.Get(row, "table_id");
                if (!allTableIds.Contains(tableId))
                {
                    catalog.Issues.Add(Error(
                        "constraint.table_id",
                        tableId,
                        "Constraint references an unknown table ID.",
                        RowLocation("Constraints", row)));
                    continue;
                }

                ConstraintTableDefinition table;
                if (!tablesById.TryGetValue(tableId, out table))
                {
                    continue;
                }

                string comparator = ConstraintValueNormalizer.NormalizeComparator(
                    constraintRows.Get(row, "comparator"));
                var constraint = new ConstraintDefinition
                {
                    ConstraintId = constraintRows.Get(row, "constraint_id"),
                    TableId = tableId,
                    StructureId = EmptyToNull(constraintRows.Get(row, "structure_id")),
                    StructureName = EmptyToNull(constraintRows.Get(row, "structure_name")),
                    RawStructureName = EmptyToNull(constraintRows.Get(row, "structure_name")),
                    Metric = ConstraintValueNormalizer.NormalizeMetric(
                        constraintRows.Get(row, "metric")),
                    Unit = ConstraintValueNormalizer.NormalizeUnit(
                        constraintRows.Get(row, "unit")),
                    Comparator = comparator,
                    ExpectedDirection = DirectionFromComparator(comparator),
                    Source = constraintRows.Get(row, "source"),
                    Comment = constraintRows.Get(row, "comment")
                };
                AssignDecimal(constraintRows, row, "goal", value => constraint.Goal = value, catalog.Issues);
                AssignDecimal(constraintRows, row, "variation", value => constraint.Variation = value, catalog.Issues);
                AssignInteger(constraintRows, row, "priority", value => constraint.Priority = value, catalog.Issues);
                constraint.DvhObjective = BuildDvhObjective(constraint.Metric, constraint.Unit);
                constraint.EvaluationPoint = BuildEvaluationPoint(
                    constraint.Comparator,
                    constraint.Goal);
                foreach (KeyValuePair<string, int> header in constraintRows.Headers)
                {
                    constraint.RawValues[header.Key] = row.GetCell(header.Value);
                }

                table.Constraints.Add(constraint);
            }

            foreach (CatalogValidationIssue issue in ConstraintCatalogValidator.Validate(catalog))
            {
                catalog.Issues.Add(issue);
            }

            return catalog;
        }

        private static bool TryGetSheet(
            XlsxWorkbookData workbook,
            string name,
            IList<CatalogValidationIssue> issues,
            out XlsxWorksheetData worksheet)
        {
            if (workbook.Worksheets.TryGetValue(name, out worksheet))
            {
                return true;
            }

            issues.Add(Error(
                "worksheet",
                name,
                "Required worksheet is missing.",
                "Excel"));
            return false;
        }

        private static SheetRows ReadRows(
            XlsxWorksheetData worksheet,
            IEnumerable<string> requiredHeaders,
            IList<CatalogValidationIssue> issues)
        {
            var result = new SheetRows();
            XlsxRowData headerRow = worksheet.Rows.FirstOrDefault();
            if (headerRow == null)
            {
                issues.Add(Error(
                    "missing_header",
                    null,
                    "Worksheet has no header row.",
                    worksheet.Name));
                return result;
            }

            foreach (KeyValuePair<int, string> cell in headerRow.Cells)
            {
                string header = (cell.Value ?? string.Empty).Trim();
                if (!string.IsNullOrWhiteSpace(header) && !result.Headers.ContainsKey(header))
                {
                    result.Headers.Add(header, cell.Key);
                }
            }

            foreach (string required in requiredHeaders)
            {
                if (!result.Headers.ContainsKey(required))
                {
                    issues.Add(Error(
                        "missing_header",
                        required,
                        "Required header is missing: " + required,
                        worksheet.Name));
                }
            }

            foreach (XlsxRowData row in worksheet.Rows.Skip(1))
            {
                if (row.Cells.Values.Any(value => !string.IsNullOrWhiteSpace(value)))
                {
                    result.Rows.Add(row);
                }
            }

            result.SheetName = worksheet.Name;
            return result;
        }

        private static void ReportDuplicates(
            SheetRows rows,
            string field,
            IList<CatalogValidationIssue> issues)
        {
            foreach (IGrouping<string, XlsxRowData> duplicate in rows.Rows
                .GroupBy(row => rows.Get(row, field), StringComparer.OrdinalIgnoreCase)
                .Where(group => !string.IsNullOrWhiteSpace(group.Key) && group.Count() > 1))
            {
                issues.Add(Error(
                    field,
                    duplicate.Key,
                    field + " is duplicate.",
                    rows.SheetName));
            }
        }

        private static void AssignDecimal(
            SheetRows rows,
            XlsxRowData row,
            string field,
            Action<decimal?> assign,
            IList<CatalogValidationIssue> issues)
        {
            try
            {
                assign(ConstraintValueNormalizer.ParseNullableDecimal(rows.Get(row, field)));
            }
            catch (FormatException exception)
            {
                issues.Add(Error(
                    field,
                    rows.Get(row, field),
                    exception.Message,
                    RowLocation(rows.SheetName, row)));
            }
        }

        private static void AssignInteger(
            SheetRows rows,
            XlsxRowData row,
            string field,
            Action<int?> assign,
            IList<CatalogValidationIssue> issues)
        {
            try
            {
                assign(ConstraintValueNormalizer.ParseNullableInteger(rows.Get(row, field)));
            }
            catch (FormatException exception)
            {
                issues.Add(Error(
                    field,
                    rows.Get(row, field),
                    exception.Message,
                    RowLocation(rows.SheetName, row)));
            }
        }

        private static void AddDelimited(IList<string> destination, string value)
        {
            foreach (string item in (value ?? string.Empty)
                .Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim())
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                destination.Add(item);
            }
        }

        private static bool ParseBoolean(string value, bool fallback)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return fallback;
            }

            return normalized.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                   normalized.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                   normalized.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
                   normalized.Equals("ja", StringComparison.OrdinalIgnoreCase);
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
            string objective = DvhMetricName(metric);
            return string.IsNullOrWhiteSpace(unit)
                ? objective
                : string.Format("{0}[{1}]", objective, unit);
        }

        private static string DvhMetricName(string metric)
        {
            if (string.Equals(metric, "Dmean", StringComparison.OrdinalIgnoreCase))
            {
                return "Mean";
            }

            if (string.Equals(metric, "Dmax", StringComparison.OrdinalIgnoreCase))
            {
                return "Max";
            }

            if (string.Equals(metric, "Dmin", StringComparison.OrdinalIgnoreCase))
            {
                return "Min";
            }

            return metric;
        }

        private static string BuildEvaluationPoint(string comparator, decimal? goal)
        {
            return goal.HasValue
                ? string.Format(
                    CultureInfo.InvariantCulture,
                    "{0}{1:0.###}",
                    comparator,
                    goal.Value)
                : comparator;
        }

        private static string EmptyToNull(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }

        private static string RowLocation(string sheetName, XlsxRowData row)
        {
            return string.Format("{0}!{1}", sheetName, row.RowNumber);
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

        private sealed class SheetRows
        {
            public SheetRows()
            {
                Headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                Rows = new List<XlsxRowData>();
            }

            public string SheetName { get; set; }
            public IDictionary<string, int> Headers { get; private set; }
            public IList<XlsxRowData> Rows { get; private set; }

            public string Get(XlsxRowData row, string header)
            {
                int index;
                return Headers.TryGetValue(header, out index)
                    ? row.GetCell(index).Trim()
                    : string.Empty;
            }
        }
    }
}
