using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel.Shapes;
using System;
using System.IO;
using System.Reflection;


namespace ClearPlan.Reporting.MigraDoc.Internal
{
    internal class RefsContent
    {
        public void Add(Section section, ReportStructureSet structureSet, ReportPlanningItem reportPlanningItem, ReportPQMs reportPQMs, ReportRefs reportRefs, ReportDVH reportDVH)
        {
            var table2 = AddHeadingTable2(section);
            AddHeadingLeft(table2.Rows[0].Cells[0], section, structureSet, reportPlanningItem);
            AddHeadingRight(table2.Rows[0].Cells[1], section);

            //AddHeading(section, structureSet, reportPlanningItem);
            //AddPQMs(section, reportPQMs, reportPlanningItem);
            //var table3 = AddHeadingTable(section);
            //AddHeading(section, structureSet, reportPlanningItem);
            AddRefs(section, reportRefs);

            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.LineSpacingRule = LineSpacingRule.Exactly;
            paragraph.Format.LineSpacing = Unit.FromMillimeter(6);

            AddDVH(section, reportDVH);
           // section.AddPageBreak();

        }

        private Table AddHeadingTable2(Section section)
        {
            var table2 = section.AddTable();
            //table.Shading.Color = Shading;

            table2.Rows.LeftIndent = 0;

            table2.LeftPadding = Size.TableCellPadding;
            table2.TopPadding = Size.TableCellPadding;
            table2.RightPadding = Size.TableCellPadding;
            table2.BottomPadding = Size.TableCellPadding;

            // Use two columns of equal width
            var columnWidth1 = Size.GetWidth(section) * 9.0 / 10;
            var columnWidth2 = Size.GetWidth(section) * 1.0 / 10;
            table2.AddColumn(columnWidth1);
            table2.AddColumn(columnWidth2);

            // Only one row is needed
            table2.AddRow();

            return table2;
        }

        static string MigraDocFilenameFromByteArray(byte[] image)
        {
            return "base64:" +
                   Convert.ToBase64String(image);
        }

        private byte[] LoadResource(string name)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var fullName = $"{assembly.GetName().Name}.{name}";
            using (var stream = assembly.GetManifestResourceStream(fullName))
            {
                if (stream == null)
                {
                    throw new ArgumentException($"No resource with name {name}");
                }

                var count = (int)stream.Length;
                var data = new byte[count];
                stream.Read(data, 0, count);
                return data;
            }
        }

        static byte[] LoadImage(string name)
        {
            var assembly = Assembly.GetExecutingAssembly();

            using (Stream stream = assembly.GetManifestResourceStream(name))
            {
                if (stream == null)
                    throw new ArgumentException("No resource with name " + name);

                int count = (int)stream.Length;
                byte[] data = new byte[count];
                stream.Read(data, 0, count);
                return data;
            }
        }

        private string ConvertToMigraDocFileName(byte[] image)
        {
            return $"base64:{Convert.ToBase64String(image)}";
        }

        private void AddHeadingLeft(Cell cell, Section section, ReportStructureSet structureSet, ReportPlanningItem reportPlanningItem)
        {
            Paragraph p1 = cell.AddParagraph();
            Paragraph p2 = cell.AddParagraph();
            p1.AddSpace(1);
            //p.Style = StyleNames.Heading1;
            AddPlanningItemSymbol(p2, reportPlanningItem);
            //var section2 = section.Headers.Primary.AddParagraph();
            //p.Format.AddTabStop(Size.GetWidth(section), TabAlignment.Right);
            //p.AddText(reportPlanningItem.Type + " : " + reportPlanningItem.Id);
            var myFont = new Font();
            myFont.Size = 21;
            p2.AddFormattedText(reportPlanningItem.Type + " : " + reportPlanningItem.Id, myFont);

            cell.AddParagraph($"created {reportPlanningItem.Created:g}");
            cell.AddParagraph($"Image: '{structureSet.Image.Id}' " +
                                 $"taken {structureSet.Image.CreationTime:g}" + $"{reportPlanningItem.orientation}");
            if (reportPlanningItem.Type == "Plan")
            {
                cell.AddParagraph($"Prescription: GD = {reportPlanningItem.prinTtotaldose} Gy, " +
                    $"ED = {reportPlanningItem.prinTfractiondose} Gy, " +
                                 $"Fx = {reportPlanningItem.prinTfractions}" + $" | " + $"{reportPlanningItem.gatingExtenedString}");
                cell.AddParagraph($"MLC: {reportPlanningItem.technique}" + $" | " + $"Calculation Model: {reportPlanningItem.algorithmus}" + reportPlanningItem.planningApprover);
            }

        }
        private void AddHeadingRight(Cell cell, Section section)
        {
            Paragraph p = cell.AddParagraph();
            //p.Style = StyleNames.Heading1;
            //AddPlanningItemSymbol(p, reportPlanningItem);
            //var section2 = section.Headers.Primary.AddParagraph();
            //p.Format.AddTabStop(Size.GetWidth(section), TabAlignment.Right);
            //p.AddText(reportPlanningItem.Type + " : " + reportPlanningItem.Id);

            //section.AddParagraph($"created {reportPlanningItem.Created:g}");
            //section.AddParagraph($"Image: '{structureSet.Image.Id}' " +
            //                    $"taken {structureSet.Image.CreationTime:g}");
            byte[] image = LoadResource("Resources.logo.png");
            //logo
            string imageFilename = ConvertToMigraDocFileName(image);
            p.AddSpace(3);
            Image image1 = p.AddImage(imageFilename);
            image1.Resolution = 300;
            image1.Width = "1.45cm";
            //image1.Height = "1.5cm";
            image1.LockAspectRatio = true;
            image1.Left = ShapePosition.Right;
        }

        private void AddPlanningItemSymbol(Paragraph p, ReportPlanningItem reportPlanningItem)
        {
            Image image1 = p.AddImage(new PlanningItemSymbol(reportPlanningItem).GetMigraDocFileName());
            image1.Height = "0.54cm";
        }

        /*public void Add(Section section, ReportRefs reportRefs)
        {
            var table2 = AddHeadingTable(section);
            

            //AddHeading(section, structureSet, reportPlanningItem);
            AddRefs(section, reportRefs);
        }*/

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
            var columnWidth1 = Size.GetWidth(section) * 2.0/10;
            var columnWidth2 = Size.GetWidth(section) * 2.0/10;
            var columnWidth3 = Size.GetWidth(section) * 1.0 / 10;
            var columnWidth4 = Size.GetWidth(section) * 1.0 / 10;
            var columnWidth5 = Size.GetWidth(section) * 1.0 / 10;
            var columnWidth6 = Size.GetWidth(section) * 1.0 / 10;
            var columnWidth7 = Size.GetWidth(section) * 1.0 / 10;
            var columnWidth8 = Size.GetWidth(section) * 1.0 / 10;
            table2.AddColumn(columnWidth1);
            table2.AddColumn(columnWidth2);
            table2.AddColumn(columnWidth3);
            table2.AddColumn(columnWidth4);
            table2.AddColumn(columnWidth5);
            table2.AddColumn(columnWidth6);
            table2.AddColumn(columnWidth7);
            table2.AddColumn(columnWidth8);

            // Only one row is needed
            table2.AddRow();

            return table2;
        }


        private void AddDVH(Section section, ReportDVH DVH)
        {
            var byteDVH = ImageToByte(DVH._dvhBitmap);

            //Now converting it to string as seen in the WikiPage

            string stringDVH = Convert.ToBase64String(byteDVH);
            string FinalDVH = "base64:" + stringDVH;
            Image img = section.AddImage(FinalDVH);
            img.Width = 530;
            img.Left = ShapePosition.Center;
        }
        //Code for converting bitmap to byte[]
        public static byte[] ImageToByte(System.Drawing.Image img)
        {
            using (var ms = new MemoryStream())
            {
                System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(img);
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                return ms.ToArray();
            }
        }


        private void AddRefs(Section section, ReportRefs Refs)
        {
            
            AddTableTitle(section, "Target Table");
            AddRefTable(section, Refs);
        }

        private void AddTableTitle(Section section, string title)
        {
            var p = section.AddParagraph(title, StyleNames.Heading2);
            p.Format.KeepWithNext = true;
        }

        private void AddRefTable(Section section, ReportRefs Refs)
        {
            var table = section.AddTable();

            FormatTable(table);
            AddColumnsAndHeaders(table);
            AddRefRows(table, Refs);

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
            table.AddColumn(width * 0.40);
            table.AddColumn(width * 0.09);
            table.AddColumn(width * 0.09);
            table.AddColumn(width * 0.06);
            table.AddColumn(width * 0.09);
            table.AddColumn(width * 0.09);
            table.AddColumn(width * 0.09);
            table.AddColumn(width * 0.09);


            var headerRow = table.AddRow();
            headerRow.Borders.Bottom.Width = 1;

            AddHeader(headerRow.Cells[0], "RP-ID (Plan-ID)");
            AddHeader(headerRow.Cells[1], "GD [Gy]");
            AddHeader(headerRow.Cells[2], "ED [Gy]");
            AddHeader(headerRow.Cells[3], "Fx");
            AddHeader(headerRow.Cells[4], "D50[Gy]");
            AddHeader(headerRow.Cells[5], "D95[%]");
            AddHeader(headerRow.Cells[6], "D2[%]");
            AddHeader(headerRow.Cells[7], "D98[%]");

        }

        private void AddHeader(Cell cell, string header)
        {
            var p = cell.AddParagraph(header);
            p.Style = CustomStyles.ColumnHeader;
        }

        private void AddRefRows(Table table, ReportRefs Refs)
        {
            foreach (var re in Refs.Refs)
            {

                var row = table.AddRow();
                row.Format.Font.Size = 10;
                row.VerticalAlignment = VerticalAlignment.Center;

               
                var id = row.Cells[0].AddParagraph();
                if (re.RefPointId.Contains("[PRIMÄR"))
                    id.AddFormattedText(re.RefPointId.Replace("[PRIMÄR] ",""), TextFormat.Bold);
                else
                    id.AddFormattedText(re.RefPointId, TextFormat.NotBold);
                row.Cells[1].AddParagraph(re.Prescription.ToString("#.##"));
                row.Cells[2].AddParagraph(re.Session.ToString("#.##"));
                row.Cells[3].AddParagraph(re.Fractions.ToString("#.##"));
                row.Cells[4].AddParagraph(re.D50.ToString("#.##"));
                row.Cells[5].AddParagraph(re.D95.ToString("#.##"));
                row.Cells[6].AddParagraph(re.D2.ToString("#.##"));
                row.Cells[7].AddParagraph(re.D98.ToString("#.##"));


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
