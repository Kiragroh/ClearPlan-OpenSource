using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using ClearPlan.Reporting.MigraDoc.Internal;

namespace ClearPlan.Reporting.MigraDoc
{
    public class ReportPdf : IReport
    {
        public void Export(string path, ReportData data)
        {
            ExportPdf(path, CreateReport(data));
        }

        public void Export(string path, ReviewReportDocument data)
        {
            if (data == null)
            {
                throw new ArgumentNullException("data");
            }

            ExportPdf(path, CreateReviewReport(data));
        }

        private void ExportPdf(string path, Document report)
        {
            var pdfRenderer = new PdfDocumentRenderer();
            pdfRenderer.Document = report;
            pdfRenderer.RenderDocument();
            pdfRenderer.PdfDocument.Save(path);
        }

        private Document CreateReport(ReportData data)
        {
            var doc = new Document();
            CustomStyles.Define(doc);
            doc.Add(CreateMainSection(data));
            return doc;
        }

        private Document CreateReviewReport(ReviewReportDocument data)
        {
            var document = new Document();
            CustomStyles.Define(document);
            string watermark = data.Synthetic
                ? ReviewSnapshotReportMapper.SyntheticWatermark
                : data.Watermark;

            document.Info.Title = Safe(
                data.Title,
                "ClearPlan review report");
            document.Info.Subject = Safe(watermark, data.ModeLabel);

            var section = new Section();
            SetUpReviewPage(section);
            AddReviewHeaderAndFooter(section, data, watermark);
            AddReviewContents(section, data, watermark);
            document.Add(section);
            return document;
        }

        private Section CreateMainSection(ReportData data)
        {
            var section = new Section();
            SetUpPage(section);
            AddHeaderAndFooter(section, data);
            AddContents(section, data);
            return section;
        }

        private void SetUpPage(Section section)
        {
            section.PageSetup.PageFormat = PageFormat.Letter;
            //section.PageSetup.Orientation = Orientation.Landscape;
            section.PageSetup.LeftMargin = Size.LeftRightPageMargin;
            section.PageSetup.TopMargin = Size.TopBottomPageMargin;
            section.PageSetup.RightMargin = Size.LeftRightPageMargin;
            section.PageSetup.BottomMargin = Size.TopBottomPageMargin;

            section.PageSetup.HeaderDistance = Size.HeaderFooterMargin;
            section.PageSetup.FooterDistance = Size.HeaderFooterMargin;
        }

        private void SetUpReviewPage(Section section)
        {
            section.PageSetup.PageFormat = PageFormat.A4;
            section.PageSetup.Orientation = Orientation.Landscape;
            section.PageSetup.LeftMargin = Unit.FromCentimeter(1.2);
            section.PageSetup.TopMargin = Unit.FromCentimeter(1.8);
            section.PageSetup.RightMargin = Unit.FromCentimeter(1.2);
            section.PageSetup.BottomMargin = Unit.FromCentimeter(1.5);
            section.PageSetup.HeaderDistance = Unit.FromCentimeter(0.6);
            section.PageSetup.FooterDistance = Unit.FromCentimeter(0.6);
        }

        private void AddReviewHeaderAndFooter(
            Section section,
            ReviewReportDocument data,
            string watermark)
        {
            var header = section.Headers.Primary.AddParagraph();
            header.Format.Alignment = ParagraphAlignment.Center;
            header.Format.Font.Bold = true;
            header.Format.Font.Color = Colors.DarkRed;
            header.AddText(Safe(watermark, data.ModeLabel));

            var footer = section.Footers.Primary.AddParagraph();
            footer.Format.AddTabStop(
                Unit.FromCentimeter(25.0),
                TabAlignment.Right);
            footer.AddText(
                "ClearPlan · " +
                Safe(data.ScenarioId, "detached review"));
            footer.AddTab();
            footer.AddText("Page ");
            footer.AddPageField();
            footer.AddText(" of ");
            footer.AddNumPagesField();
        }

        private void AddReviewContents(
            Section section,
            ReviewReportDocument data,
            string watermark)
        {
            Paragraph banner = section.AddParagraph(
                Safe(watermark, data.ModeLabel));
            banner.Format.Alignment = ParagraphAlignment.Center;
            banner.Format.Font.Size = 15;
            banner.Format.Font.Bold = true;
            banner.Format.Font.Color = Colors.White;
            banner.Format.Shading.Color = Colors.DarkRed;
            banner.Format.SpaceAfter = Unit.FromCentimeter(0.3);

            Paragraph title = section.AddParagraph(
                Safe(data.Title, "ClearPlan review report"));
            title.Style = StyleNames.Heading1;
            title.AddLineBreak();
            title.AddFormattedText(
                Safe(data.Subtitle, data.ScenarioTitle),
                TextFormat.NotBold);

            Paragraph provenance = section.AddParagraph();
            provenance.AddFormattedText("Scenario: ", TextFormat.Bold);
            provenance.AddText(Safe(data.ScenarioId, "detached review"));
            provenance.AddText(" · ");
            provenance.AddFormattedText("Generated: ", TextFormat.Bold);
            provenance.AddText(
                data.GeneratedUtc.ToUniversalTime().ToString(
                    "yyyy-MM-dd HH:mm 'UTC'",
                    CultureInfo.InvariantCulture));
            provenance.AddLineBreak();
            provenance.AddText(Safe(data.ProvenanceText, string.Empty));

            AddPlanRows(section, data.Plans);
            AddSourceRows(section, data.Sources);
            AddPqmRows(section, data.PqmRows);
            AddPlanCheckRows(section, data.PlanCheckRows);
            AddFieldRows(section, data.FieldRows);
            AddMappingRows(section, data.StructureMappings);
            AddDvhRows(section, data.DvhSeries);
            AddNotes(section, data.Notes);
        }

        private void AddPlanRows(
            Section section,
            IList<ReviewReportPlanRow> rows)
        {
            AddHeading(section, "Plans");
            Table table = CreateTable(
                section,
                4.0,
                3.0,
                2.0,
                2.0,
                1.8,
                3.0,
                2.0);
            AddHeader(
                table,
                "Plan",
                "Created",
                "Dose/Fx [Gy]",
                "Total [Gy]",
                "Fractions",
                "Target",
                "Status");
            foreach (ReviewReportPlanRow row in rows ??
                     new List<ReviewReportPlanRow>())
            {
                Row target = table.AddRow();
                SetCells(
                    target,
                    row.DisplayLabel,
                    FormatUtc(row.CreatedUtc),
                    FormatNumber(row.DosePerFractionGy),
                    FormatNumber(row.TotalDoseGy),
                    FormatNumber(row.FractionCount),
                    row.TargetDisplayLabel,
                    row.Status);
                ShadeStatus(target.Cells[6], row.Status);
            }
        }

        private void AddSourceRows(
            Section section,
            IList<ReviewReportSourceRow> rows)
        {
            AddHeading(section, "Source status");
            Table table = CreateTable(section, 4.0, 3.0, 2.0, 2.0, 9.0);
            AddHeader(
                table,
                "Source",
                "Type",
                "Status",
                "Fallback",
                "Message");
            foreach (ReviewReportSourceRow row in rows ??
                     new List<ReviewReportSourceRow>())
            {
                Row target = table.AddRow();
                SetCells(
                    target,
                    row.PathDisplayLabel,
                    row.SourceType,
                    row.Status,
                    row.UsedFallback ? "yes" : "no",
                    row.Message);
                ShadeStatus(target.Cells[2], row.Status);
            }
        }

        private void AddPqmRows(
            Section section,
            IList<ReviewReportPqmRow> rows)
        {
            AddHeading(section, "Plan quality metrics (PQM)");
            Table table = CreateTable(
                section,
                3.0,
                3.0,
                2.2,
                1.4,
                2.0,
                2.0,
                2.0,
                2.0);
            AddHeader(
                table,
                "Template",
                "Resolved structure",
                "Objective",
                "Rule",
                "Goal",
                "Variation",
                "Achieved",
                "Status");
            foreach (ReviewReportPqmRow row in rows ??
                     new List<ReviewReportPqmRow>())
            {
                Row target = table.AddRow();
                SetCells(
                    target,
                    row.TemplateStructure,
                    row.ResolvedStructureId,
                    row.Objective,
                    row.Comparator,
                    FormatValue(row.Goal, row.Unit),
                    FormatValue(row.Variation, row.Unit),
                    FormatValue(row.AchievedValue, row.Unit),
                    StatusWithSeverity(row.Status, row.Severity));
                ShadeStatus(target.Cells[7], row.Status);
            }
        }

        private void AddPlanCheckRows(
            Section section,
            IList<ReviewReportCheckRow> rows)
        {
            AddHeading(section, "PlanCheck");
            Table table = CreateTable(
                section,
                3.0,
                2.5,
                3.0,
                3.0,
                2.0,
                2.5,
                5.0);
            AddHeader(
                table,
                "Check",
                "Category",
                "Observed",
                "Expected",
                "Unit",
                "Status",
                "Message");
            foreach (ReviewReportCheckRow row in rows ??
                     new List<ReviewReportCheckRow>())
            {
                Row target = table.AddRow();
                SetCells(
                    target,
                    row.CheckCode,
                    row.Category,
                    row.ObservedValue,
                    row.ExpectedValue,
                    row.Unit,
                    StatusWithSeverity(row.Status, row.Severity),
                    row.Message);
                ShadeStatus(target.Cells[5], row.Status);
            }
        }

        private void AddFieldRows(
            Section section,
            IList<ReviewReportFieldRow> rows)
        {
            AddHeading(section, "Field identifiers and names");
            Table table = CreateTable(
                section,
                1.5,
                1.5,
                2.5,
                2.5,
                4.0,
                4.0,
                2.0,
                2.0);
            AddHeader(
                table,
                "Order",
                "Beam #",
                "Current ID",
                "Expected ID",
                "Current name",
                "Suggested name",
                "ID status",
                "Name status");
            foreach (ReviewReportFieldRow row in rows ??
                     new List<ReviewReportFieldRow>())
            {
                Row target = table.AddRow();
                SetCells(
                    target,
                    row.TreatmentOrder.ToString(
                        CultureInfo.InvariantCulture),
                    row.BeamNumber.ToString(CultureInfo.InvariantCulture),
                    row.CurrentId,
                    row.ExpectedId,
                    row.CurrentName,
                    row.SuggestedName,
                    row.IdStatus,
                    row.NameStatus);
                ShadeStatus(target.Cells[6], row.IdStatus);
                ShadeStatus(target.Cells[7], row.NameStatus);
            }
        }

        private void AddMappingRows(
            Section section,
            IList<ReviewReportStructureMappingRow> rows)
        {
            AddHeading(section, "Structure mapping");
            Table table = CreateTable(section, 4.0, 4.0, 2.5, 9.0);
            AddHeader(
                table,
                "Template structure",
                "Resolved structure",
                "Status",
                "Message");
            foreach (ReviewReportStructureMappingRow row in rows ??
                     new List<ReviewReportStructureMappingRow>())
            {
                Row target = table.AddRow();
                SetCells(
                    target,
                    row.TemplateStructure,
                    row.SelectedStructureId,
                    row.Status,
                    row.Message);
                ShadeStatus(target.Cells[2], row.Status);
            }
        }

        private void AddDvhRows(
            Section section,
            IList<ReviewReportDvhSeries> rows)
        {
            AddHeading(section, "Dose-volume histogram (DVH)");
            Table table = CreateTable(
                section,
                4.0,
                3.0,
                2.5,
                2.5,
                2.5,
                2.5,
                3.0);
            AddHeader(
                table,
                "Structure",
                "Role",
                "Volume [cm3]",
                "Dose unit",
                "Volume unit",
                "Samples",
                "Dose range");
            foreach (ReviewReportDvhSeries row in rows ??
                     new List<ReviewReportDvhSeries>())
            {
                IList<ReviewReportDvhPoint> points =
                    row.Points ?? new List<ReviewReportDvhPoint>();
                string doseRange = points.Count == 0
                    ? string.Empty
                    : string.Format(
                        CultureInfo.InvariantCulture,
                        "{0:0.###}–{1:0.###} {2}",
                        points.First().DoseGy,
                        points.Last().DoseGy,
                        row.DoseUnit);
                Row target = table.AddRow();
                SetCells(
                    target,
                    row.DisplayName,
                    row.Role,
                    FormatNumber(row.VolumeCc),
                    row.DoseUnit,
                    row.VolumeUnit,
                    points.Count.ToString(CultureInfo.InvariantCulture),
                    doseRange);
            }

            Paragraph note = section.AddParagraph(
                "The detached report model retains every sampled " +
                "(dose in Gy, cumulative relative volume in %) point. " +
                "This table summarizes each series for the PDF.");
            note.Format.Font.Size = 7;
            note.Format.Font.Italic = true;
        }

        private void AddNotes(Section section, IList<string> notes)
        {
            if (notes == null || notes.Count == 0)
            {
                return;
            }

            AddHeading(section, "Notes");
            foreach (string note in notes)
            {
                Paragraph paragraph = section.AddParagraph();
                paragraph.Format.LeftIndent = Unit.FromCentimeter(0.4);
                paragraph.AddText("• " + Safe(note, string.Empty));
            }
        }

        private void AddHeading(Section section, string text)
        {
            section.AddParagraph(text, StyleNames.Heading2);
        }

        private Table CreateTable(
            Section section,
            params double[] columnWidthsCentimeter)
        {
            Table table = section.AddTable();
            table.Borders.Width = 0.25;
            table.Rows.LeftIndent = 0;
            table.Format.Font.Size = 7.5;
            table.LeftPadding = Unit.FromMillimeter(1.2);
            table.RightPadding = Unit.FromMillimeter(1.2);
            table.TopPadding = Unit.FromMillimeter(0.8);
            table.BottomPadding = Unit.FromMillimeter(0.8);
            foreach (double width in columnWidthsCentimeter)
            {
                table.AddColumn(Unit.FromCentimeter(width));
            }

            return table;
        }

        private void AddHeader(Table table, params string[] labels)
        {
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Font.Bold = true;
            row.Shading.Color = Colors.LightGray;
            SetCells(row, labels);
        }

        private void SetCells(Row row, params string[] values)
        {
            int columnCount = row.Table == null
                ? row.Cells.Count
                : row.Table.Columns.Count;
            for (int index = 0;
                 index < values.Length && index < columnCount;
                 index++)
            {
                row.Cells[index].AddParagraph(
                    Safe(values[index], string.Empty));
            }
        }

        private void ShadeStatus(Cell cell, string status)
        {
            switch ((status ?? string.Empty).ToLowerInvariant())
            {
                case "pass":
                case "available":
                    cell.Shading.Color = Color.FromRgb(218, 242, 224);
                    break;
                case "variation":
                case "fallback":
                    cell.Shading.Color = Color.FromRgb(255, 239, 184);
                    break;
                case "fail":
                case "unavailable":
                    cell.Shading.Color = Color.FromRgb(255, 214, 214);
                    break;
            }
        }

        private string StatusWithSeverity(string status, string severity)
        {
            if (string.IsNullOrWhiteSpace(severity) ||
                severity.Equals("none", StringComparison.OrdinalIgnoreCase))
            {
                return Safe(status, string.Empty);
            }

            return string.Format(
                CultureInfo.InvariantCulture,
                "{0} ({1})",
                Safe(status, string.Empty),
                severity);
        }

        private string FormatValue(double? value, string unit)
        {
            string formatted = FormatNumber(value);
            if (formatted.Length == 0)
            {
                return formatted;
            }

            return string.IsNullOrWhiteSpace(unit)
                ? formatted
                : formatted + " " + unit;
        }

        private string FormatNumber(double? value)
        {
            return value.HasValue
                ? value.Value.ToString("0.###", CultureInfo.InvariantCulture)
                : string.Empty;
        }

        private string FormatNumber(int? value)
        {
            return value.HasValue
                ? value.Value.ToString(CultureInfo.InvariantCulture)
                : string.Empty;
        }

        private string FormatUtc(DateTimeOffset? value)
        {
            return value.HasValue
                ? value.Value.ToUniversalTime().ToString(
                    "yyyy-MM-dd HH:mm",
                    CultureInfo.InvariantCulture)
                : string.Empty;
        }

        private string Safe(string value, string fallback)
        {
            return string.IsNullOrWhiteSpace(value)
                ? (fallback ?? string.Empty)
                : value;
        }
        

        private void AddHeaderAndFooter(Section section, ReportData data)
        {
            new HeaderAndFooter().Add(section, data);
        }

        private void AddContents(Section section, ReportData data)
        {

            AddPatientInfo(section, data.ReportPatient, data.ReportPlanningItem);
            AddRefs(section, data.ReportStructureSet, data.ReportPlanningItem, data.ReportPQMs, data.ReportRefs, data.ReportDVH);
            AddPCs(section, data.ReportPCs, data.ReportPlanningItem);
            //section.AddPageBreak();
            AddPQMs(section, data.ReportStructureSet, data.ReportPlanningItem, data.ReportPQMs, data.ReportPatient);
            AddDVHstats(section, data.ReportDVHstats);
        }

        private void AddPatientInfo(Section section, ReportPatient patient, ReportPlanningItem reportPlanningItem)
        {
            new PatientInfo().Add(section, patient, reportPlanningItem);
        }

        private void AddPQMs(Section section, ReportStructureSet structureSet, ReportPlanningItem reportPlanningItem, ReportPQMs reportPQMs, ReportPatient reportPatient)
        {
            new PQMsContent().Add(section, structureSet, reportPlanningItem, reportPQMs, reportPatient);
        }

        private void AddRefs(Section section, ReportStructureSet structureSet, ReportPlanningItem reportPlanningItem, ReportPQMs reportPQMs, ReportRefs reportRefs, ReportDVH reportDVH)
        {
            new RefsContent().Add(section, structureSet, reportPlanningItem, reportPQMs, reportRefs, reportDVH);
        }

        private void AddPCs(Section section, ReportPCs reportPCs,ReportPlanningItem reportPlanningItem)
        {
            if (reportPlanningItem.PrintPC_CheckboxState == true)
            { new PCsContent().Add(section, reportPCs); }
        }
        private void AddDVHstats(Section section, ReportDVHstats reportDVHstats)
        {
            new DVHstatsContent().Add(section, reportDVHstats);
        }

    }
}
