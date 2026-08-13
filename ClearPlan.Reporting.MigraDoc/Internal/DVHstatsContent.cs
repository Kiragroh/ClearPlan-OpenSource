using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel.Shapes;
using System;
using System.IO;
using System.Reflection;

namespace ClearPlan.Reporting.MigraDoc.Internal
{
    internal class DVHstatsContent
    {
        public void Add(Section section, ReportDVHstats reportDVHstats)
        {
            AddDVHstats(section, reportDVHstats);
        }

        private Table AddHeadingTable(Section section)
        {
            var table2 = section.AddTable();
            //table.Shading.Color = Shading;

            table2.Rows.LeftIndent = 0;

            table2.LeftPadding = Size.TableCellPadding;
            table2.TopPadding = Size.TableCellPadding;
            table2.RightPadding = Size.TableCellPadding;
            table2.BottomPadding = Size.TableCellPadding;

            // Use two columns of equal width
            var columnWidth1 = Size.GetWidth(section) * 2.5 / 10;
            var columnWidth2 = Size.GetWidth(section) * 2.5 / 10;
            var columnWidth3 = Size.GetWidth(section) * 2.5 / 10;
            var columnWidth4 = Size.GetWidth(section) * 2.5 / 10;
            var columnWidth5 = Size.GetWidth(section) * 2.5 / 10;
            table2.AddColumn(columnWidth1);
            table2.AddColumn(columnWidth2);
            table2.AddColumn(columnWidth3);
            table2.AddColumn(columnWidth4);
            table2.AddColumn(columnWidth5);

            // Only one row is needed
            table2.AddRow();

            return table2;
        }



        private void AddDVHstats(Section section, ReportDVHstats DVHstats)
        {
            AddTableTitle(section, "DVH Metrics");
            AddDVHstatsTable(section, DVHstats);
        }

        private void AddTableTitle(Section section, string title)
        {
            var p = section.AddParagraph(title, StyleNames.Heading2);
            p.Format.KeepWithNext = true;
        }

        private void AddDVHstatsTable(Section section, ReportDVHstats DVHstats)
        {
            var table = section.AddTable();

            FormatTable(table);
            AddColumnsAndHeaders(table);
            AddDVHstatsRows(table, DVHstats);

            AddLastRowBorder(table);
            AlternateRowShading(table);
        }

        private static void FormatTable(Table table)
        {
            table.LeftPadding = 0;
            table.TopPadding = Size.TableCellPadding;
            table.RightPadding = 0;
            table.BottomPadding = Size.TableCellPadding;
            table.Format.LeftIndent = Size.TableCellPadding;
            table.Format.RightIndent = Size.TableCellPadding;
        }

        private void AddColumnsAndHeaders(Table table)
        {
            var width = Size.GetWidth(table.Section);
            table.AddColumn(width * 0.25);
            table.AddColumn(width * 0.15);
            table.AddColumn(width * 0.2);
            table.AddColumn(width * 0.2);
            table.AddColumn(width * 0.2);
            //table.AddColumn(width * 0.1);

            var headerRow = table.AddRow();
            headerRow.Borders.Bottom.Width = 1;

            AddHeader(headerRow.Cells[0], "Structure");
            AddHeader(headerRow.Cells[1], "Volume [cc]");
            AddHeader(headerRow.Cells[2], "Max_D1% [Gy]");
            AddHeader(headerRow.Cells[3], "Min_D99% [Gy]");
            AddHeader(headerRow.Cells[4], "Mean_D50% [Gy]");

        }

        private void AddHeader(Cell cell, string header)
        {
            var p = cell.AddParagraph(header);
            p.Style = CustomStyles.ColumnHeader;
        }

        private void AddDVHstatsRows(Table table, ReportDVHstats DVHstats)
        {
            //int w = 0;
            foreach (var dvhstat in DVHstats.DVHstats)
            {
                try
                {
                    var row = table.AddRow();
                    row.Format.Font.Size = 10;
                    row.VerticalAlignment = VerticalAlignment.Center;

                    row.Cells[0].AddParagraph(dvhstat.dsStructureID);
                    row.Cells[1].AddParagraph(dvhstat.dsVolume.ToString());
                    row.Cells[2].AddParagraph(dvhstat.dsMaxNear);
                    row.Cells[3].AddParagraph(dvhstat.dsMinNear);
                    row.Cells[4].AddParagraph(dvhstat.dsMeanMedian);
                }
                catch { }
                //w++;
            }
        }

        private void AddLastRowBorder(Table table)
        {
            var lastRow = table.Rows[table.Rows.Count - 1];
            lastRow.Borders.Bottom.Width = 2;
        }

        private void AlternateRowShading(Table table)
        {
            // Start at i = 1 to skip column headers
            for (var i = 1; i < table.Rows.Count; i++)
            {
                if (i % 2 == 0)  // Even rows
                {
                    table.Rows[i].Shading.Color = Color.FromRgb(240, 240, 240);
                }
            }
        }

    }
}
