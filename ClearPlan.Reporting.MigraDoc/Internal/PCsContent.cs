using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel.Shapes;
using System;
using System.IO;
using System.Reflection;

namespace ClearPlan.Reporting.MigraDoc.Internal
{
    internal class PCsContent
    {
        public void Add(Section section, ReportPCs reportPCs)
        {
            //var table2 = AddHeadingTable(section);
            

            //AddHeading(section, structureSet, reportPlanningItem);
            AddPCs(section, reportPCs);
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
            var columnWidth1 = Size.GetWidth(section) * 9.0/10;
            var columnWidth2 = Size.GetWidth(section) * 1.0/10;
            table2.AddColumn(columnWidth1);
            table2.AddColumn(columnWidth2);

            // Only one row is needed
            table2.AddRow();

            return table2;
        }

        

        private void AddPCs(Section section, ReportPCs PCs)
        {
            AddTableTitle(section, "ClearPlans");
            AddPCTable(section, PCs);
        }

        private void AddTableTitle(Section section, string title)
        {
            var p = section.AddParagraph(title, StyleNames.Heading2);
            p.Format.KeepWithNext = true;
        }

        private void AddPCTable(Section section, ReportPCs PCs)
        {
            var table = section.AddTable();

            FormatTable(table);
            AddColumnsAndHeaders(table);
            AddPCRows(table, PCs);

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
            table.AddColumn(width * 0.85);
            table.AddColumn(width * 0.15);
            //table.AddColumn(width * 0.1);

            var headerRow = table.AddRow();
            headerRow.Borders.Bottom.Width = 1;

            AddHeader(headerRow.Cells[0], "Description");
            AddHeader(headerRow.Cells[1], "Status");
           
        }

        private void AddHeader(Cell cell, string header)
        {
            var p = cell.AddParagraph(header);
            p.Style = CustomStyles.ColumnHeader;
        }

        private void AddPCRows(Table table, ReportPCs PCs)
        {
            int w = 0;
            foreach (var pc in PCs.PCs)
            {
                if (pc.Status.StartsWith("1"))
                {
                    var row = table.AddRow();
                    row.Format.Font.Size = 10;
                    row.VerticalAlignment = VerticalAlignment.Center;

                    row.Cells[0].AddParagraph(pc.Description);
                    row.Cells[1].AddParagraph(pc.Status.Replace("Deviation", "Warning"));
                    w++;
                }
              
            }
            if (w == 0)
            {
                var row = table.AddRow();
                row.Format.Font.Size = 10;
                row.VerticalAlignment = VerticalAlignment.Center;
                row.Cells[0].AddParagraph("No Errors found.");
                row.Cells[1].AddParagraph("3 - OK");
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
