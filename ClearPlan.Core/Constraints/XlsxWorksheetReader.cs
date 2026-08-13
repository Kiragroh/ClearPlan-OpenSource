using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace ClearPlan.Core.Constraints
{
    internal sealed class XlsxWorkbookData
    {
        public XlsxWorkbookData()
        {
            Worksheets = new Dictionary<string, XlsxWorksheetData>(
                StringComparer.OrdinalIgnoreCase);
        }

        public IDictionary<string, XlsxWorksheetData> Worksheets { get; private set; }
    }

    internal sealed class XlsxWorksheetData
    {
        public XlsxWorksheetData()
        {
            Rows = new List<XlsxRowData>();
        }

        public string Name { get; set; }
        public IList<XlsxRowData> Rows { get; private set; }
    }

    internal sealed class XlsxRowData
    {
        public XlsxRowData()
        {
            Cells = new Dictionary<int, string>();
        }

        public int RowNumber { get; set; }
        public IDictionary<int, string> Cells { get; private set; }

        public string GetCell(int columnIndex)
        {
            string value;
            return Cells.TryGetValue(columnIndex, out value) ? value : string.Empty;
        }
    }

    internal static class XlsxWorksheetReader
    {
        private static readonly XNamespace SpreadsheetNamespace =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace OfficeDocumentRelationshipNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PackageRelationshipNamespace =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        public static XlsxWorkbookData Read(string path)
        {
            using (var archive = ZipFile.OpenRead(path))
            {
                XDocument workbook = LoadXml(archive, "xl/workbook.xml");
                XDocument relationships = LoadXml(archive, "xl/_rels/workbook.xml.rels");
                IList<string> sharedStrings = ReadSharedStrings(archive);
                var relationshipTargets = relationships
                    .Descendants(PackageRelationshipNamespace + "Relationship")
                    .Where(item => item.Attribute("Id") != null &&
                                   item.Attribute("Target") != null)
                    .ToDictionary(
                        item => item.Attribute("Id").Value,
                        item => NormalizePartPath("xl", item.Attribute("Target").Value),
                        StringComparer.Ordinal);

                var result = new XlsxWorkbookData();
                foreach (XElement sheet in workbook.Descendants(SpreadsheetNamespace + "sheet"))
                {
                    string name = (string)sheet.Attribute("name");
                    string relationshipId = (string)sheet.Attribute(
                        OfficeDocumentRelationshipNamespace + "id");
                    string target;
                    if (string.IsNullOrWhiteSpace(name) ||
                        string.IsNullOrWhiteSpace(relationshipId) ||
                        !relationshipTargets.TryGetValue(relationshipId, out target))
                    {
                        continue;
                    }

                    XlsxWorksheetData worksheet = ReadWorksheet(
                        archive,
                        target,
                        name,
                        sharedStrings);
                    result.Worksheets[name] = worksheet;
                }

                return result;
            }
        }

        private static XlsxWorksheetData ReadWorksheet(
            ZipArchive archive,
            string partPath,
            string name,
            IList<string> sharedStrings)
        {
            XDocument document = LoadXml(archive, partPath);
            var worksheet = new XlsxWorksheetData { Name = name };
            int inferredRowNumber = 0;
            foreach (XElement rowElement in document.Descendants(SpreadsheetNamespace + "row"))
            {
                inferredRowNumber++;
                int rowNumber;
                if (!int.TryParse(
                    (string)rowElement.Attribute("r"),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out rowNumber))
                {
                    rowNumber = inferredRowNumber;
                }

                var row = new XlsxRowData { RowNumber = rowNumber };
                int inferredColumn = 0;
                foreach (XElement cell in rowElement.Elements(SpreadsheetNamespace + "c"))
                {
                    string reference = (string)cell.Attribute("r");
                    int column = string.IsNullOrWhiteSpace(reference)
                        ? inferredColumn
                        : ColumnIndex(reference);
                    inferredColumn = column + 1;
                    row.Cells[column] = ReadCellValue(cell, sharedStrings);
                }

                worksheet.Rows.Add(row);
            }

            return worksheet;
        }

        private static string ReadCellValue(XElement cell, IList<string> sharedStrings)
        {
            string type = (string)cell.Attribute("t");
            if (string.Equals(type, "inlineStr", StringComparison.Ordinal))
            {
                return string.Concat(
                    cell.Descendants(SpreadsheetNamespace + "t").Select(item => item.Value));
            }

            string value = (string)cell.Element(SpreadsheetNamespace + "v") ?? string.Empty;
            if (string.Equals(type, "s", StringComparison.Ordinal))
            {
                int index;
                return int.TryParse(
                           value,
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture,
                           out index) &&
                       index >= 0 &&
                       index < sharedStrings.Count
                    ? sharedStrings[index]
                    : string.Empty;
            }

            if (string.Equals(type, "b", StringComparison.Ordinal))
            {
                return value == "1" ? "true" : "false";
            }

            return value;
        }

        private static IList<string> ReadSharedStrings(ZipArchive archive)
        {
            ZipArchiveEntry entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
            {
                return new List<string>();
            }

            using (Stream stream = entry.Open())
            {
                XDocument document = XDocument.Load(stream);
                return document
                    .Descendants(SpreadsheetNamespace + "si")
                    .Select(item => string.Concat(
                        item.Descendants(SpreadsheetNamespace + "t")
                            .Select(text => text.Value)))
                    .ToList();
            }
        }

        private static XDocument LoadXml(ZipArchive archive, string partPath)
        {
            ZipArchiveEntry entry = archive.GetEntry(partPath.Replace('\\', '/'));
            if (entry == null)
            {
                throw new InvalidDataException("XLSX part is missing: " + partPath);
            }

            using (Stream stream = entry.Open())
            {
                return XDocument.Load(stream);
            }
        }

        private static string NormalizePartPath(string basePath, string target)
        {
            string candidate = target.Replace('\\', '/');
            if (candidate.StartsWith("/", StringComparison.Ordinal))
            {
                return candidate.TrimStart('/');
            }

            var parts = new List<string>();
            foreach (string part in (basePath + "/" + candidate).Split('/'))
            {
                if (part == "..")
                {
                    if (parts.Count > 0)
                    {
                        parts.RemoveAt(parts.Count - 1);
                    }
                }
                else if (part != "." && part.Length > 0)
                {
                    parts.Add(part);
                }
            }

            return string.Join("/", parts);
        }

        private static int ColumnIndex(string reference)
        {
            int column = 0;
            foreach (char character in reference)
            {
                if (!char.IsLetter(character))
                {
                    break;
                }

                column = (column * 26) + (char.ToUpperInvariant(character) - 'A' + 1);
            }

            return Math.Max(0, column - 1);
        }
    }
}
