using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel.Shapes;
using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;


namespace ClearPlan.Reporting.MigraDoc.Internal
{
    internal class PQMsContent
    {
        public void Add(Section section, ReportStructureSet structureSet, ReportPlanningItem reportPlanningItem, ReportPQMs reportPQMs, ReportPatient reportPatient)
        {
            //var table2 = AddHeadingTable(section);
            //AddHeadingLeft(table2.Rows[0].Cells[0], section, structureSet, reportPlanningItem);
            //AddHeadingRight(table2.Rows[0].Cells[1], section);
            //AddHeading(section, structureSet, reportPlanningItem);

            AddPQMs(section, reportPQMs, reportPlanningItem, reportPatient);
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

        private void AddHeadingLeft(Cell cell,Section section, ReportStructureSet structureSet, ReportPlanningItem reportPlanningItem)
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
                                 $"taken {structureSet.Image.CreationTime:g}");
            if (reportPlanningItem.Type=="Plan")
            {
                cell.AddParagraph($"Prescription: GD = {reportPlanningItem.prinTtotaldose} Gy, " +
                    $"ED = {reportPlanningItem.prinTfractiondose} Gy, " +
                                 $"Fx = {reportPlanningItem.prinTfractions}");
            }
            
        }
        // moved to RefsContent
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

        private void AddPQMs(Section section, ReportPQMs PQMs, ReportPlanningItem reportPlanningItem, ReportPatient reportPatient)
        {
            AddTableTitle(section, "PQMs - " + reportPlanningItem.ConstraintPath);
            AddPQMTable(section, PQMs, reportPlanningItem, reportPatient);
        }

        private void AddTableTitle(Section section, string title)
        {
            var p = section.AddParagraph(title, StyleNames.Heading2);
            p.Format.KeepWithNext = true;
        }

        private void AddPQMTable(Section section, ReportPQMs PQMs, ReportPlanningItem reportPlanningItem, ReportPatient reportPatient)
        {
            var table = section.AddTable();

            FormatTable(table);
            AddColumnsAndHeaders(table);
            AddPQMRows(table, PQMs, reportPlanningItem, reportPatient);

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
            table.AddColumn(width * 0.16);
            table.AddColumn(width * 0.21);
            table.AddColumn(width * 0.10);
            table.AddColumn(width * 0.14);
            table.AddColumn(width * 0.16);
            table.AddColumn(width * 0.13);
            table.AddColumn(width * 0.1);
            //table.AddColumn(width * 0.1);

            var headerRow = table.AddRow();
            headerRow.Borders.Bottom.Width = 1;

            AddHeader(headerRow.Cells[0], "Template ID");
            AddHeader(headerRow.Cells[1], "Structure");
            AddHeader(headerRow.Cells[2], "Vol [cc]");
            AddHeader(headerRow.Cells[3], "Objective");
            AddHeader(headerRow.Cells[4], "Goal (Variation)");
            //AddHeader(headerRow.Cells[5], "Variation");
            AddHeader(headerRow.Cells[5], "Achieved");
            AddHeader(headerRow.Cells[6], "Met");
        }

        private void AddHeader(Cell cell, string header)
        {
            var p = cell.AddParagraph(header);
            p.Style = CustomStyles.ColumnHeader;
        }

        private void AddPQMRows(Table table, ReportPQMs PQMs, ReportPlanningItem reportPlanningItem, ReportPatient reportPatient)
        {
            #region CSV-Export Part1
            string MakeFilenameValid(string s)
            {
                char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
                foreach (char ch in invalidChars)
                {
                    s = s.Replace(ch, '_');
                }
                return s;
            }

            string userLogPath;
            string filename = string.Format("{0}_{1}_{2}.csv", MakeFilenameValid(reportPatient.Id), MakeFilenameValid(reportPlanningItem.Id.Replace(@"\", "_")), MakeFilenameValid(reportPlanningItem.ConstraintPath));
            StringBuilder userLogCsvContent = new StringBuilder();
            string exportDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Exports", "PQM-CSV");
            Directory.CreateDirectory(exportDirectory);
            userLogPath = Path.Combine(exportDirectory, filename);


            // add headers if the file doesn't exist
            // list of target headers for desired dose stats
            // in this case I want to display the headers every time so i can verify which target the distance is being measured for
            // this is due to the inconsistency in target naming (PTV1/2 vs ptv45/79.2) -- these can be removed later when cleaning up the data
            if (File.Exists(userLogPath))
            {
                File.Delete(userLogPath);
            }
            if (!File.Exists(userLogPath))
            {
                List<string> dataHeaderList = new List<string>();
                dataHeaderList.Add("User");
                // dataHeaderList.Add("Domain");
                dataHeaderList.Add("PC");
                dataHeaderList.Add("Script");
                //dataHeaderList.Add("Version");
                dataHeaderList.Add("Date");                
                dataHeaderList.Add("Time");
                dataHeaderList.Add("Patient");
                dataHeaderList.Add("TreatmentApproved by");
                dataHeaderList.Add("Plan");
                dataHeaderList.Add("Total Dose");
                dataHeaderList.Add("Fraction Dose");
                dataHeaderList.Add("Fractions");
                dataHeaderList.Add("Constraint-Template");

                string concatDataHeader = string.Join(";", dataHeaderList.ToArray());

                userLogCsvContent.AppendLine(concatDataHeader);
            }


            List<object> userStatsList = new List<object>();

            var culture = new System.Globalization.CultureInfo("de-DE");
            var day2 = culture.DateTimeFormat.GetDayName(System.DateTime.Today.DayOfWeek);

            string pc = Environment.MachineName.ToString();
            //string domain = Environment.UserDomainName.ToString();
            string userId = Environment.UserName.ToString();
            string scriptId = "ClearPlan";
            string date = System.DateTime.Now.ToString("yyyy-MM-dd");
            // string dayOfWeek = day2;
            string time = string.Format("{0}:{1}", System.DateTime.Now.ToLocalTime().ToString("HH"), System.DateTime.Now.ToLocalTime().ToString("mm"));

            userStatsList.Add(userId);
            // userStatsList.Add(domain);
            userStatsList.Add(pc);
            userStatsList.Add(scriptId);
            // userStatsList.Add(version);
            userStatsList.Add(date);
            //userStatsList.Add(dayOfWeek);
            userStatsList.Add(time);
            userStatsList.Add(MakeFilenameValid(reportPatient.Id));
            
            if (reportPlanningItem.TreatmentApproverDisplayName != "" & !reportPlanningItem.TreatmentApproverDisplayName.StartsWith("1."))
            {
                //p.AddFormattedText($" {reportPlanningItem.TreatmentApproverDisplayName}", TextFormat.Bold);
                userStatsList.Add(reportPlanningItem.TreatmentApproverDisplayName);
            }
            //PlanSum
            else if (reportPlanningItem.TreatmentApproverDisplayName.StartsWith("1."))
            {
                //p.AddFormattedText($" {reportPlanningItem.PlansInPlansum} Plans in PlanSum");
                //p.AddLineBreak();
                //p.AddFormattedText($"{reportPlanningItem.TreatmentApproverDisplayName}", TextFormat.Bold);
                userStatsList.Add(reportPlanningItem.PlansInPlansum + " Plans in PlanSum_"+ reportPlanningItem.TreatmentApproverDisplayName);
            }
            else if (reportPlanningItem.TreatmentApproverDisplayName == "")
            {
                //p.AddFormattedText($" No TreatmentApproval", TextFormat.Bold);
                userStatsList.Add(" No TreatmentApproval");
            }
            else
            {
                //p.AddFormattedText($" {reportPlanningItem.TreatmentApproverDisplayName}", TextFormat.Bold);
                userStatsList.Add(reportPlanningItem.TreatmentApproverDisplayName);
            }
            userStatsList.Add(MakeFilenameValid(reportPlanningItem.Id));
            userStatsList.Add(MakeFilenameValid(reportPlanningItem.prinTtotaldose.ToString()));
            userStatsList.Add(MakeFilenameValid(reportPlanningItem.prinTfractiondose.ToString()));
            userStatsList.Add(MakeFilenameValid(reportPlanningItem.prinTfractions));
            userStatsList.Add(MakeFilenameValid(reportPlanningItem.ConstraintPath));

            string concatUserStats = string.Join(";", userStatsList.ToArray());

            userLogCsvContent.AppendLine(concatUserStats);

            
            List<string> dataHeaderList2 = new List<string>();
            dataHeaderList2.Add("Template ID");
            dataHeaderList2.Add("Structure");
            dataHeaderList2.Add("Vol [cc]");
            dataHeaderList2.Add("Objective");
            dataHeaderList2.Add("Goal");
            dataHeaderList2.Add("Variation");
            dataHeaderList2.Add("Achieved");
            dataHeaderList2.Add("Met");

            string concatDataHeader2 = string.Join(";", dataHeaderList2.ToArray());

            userLogCsvContent.AppendLine(concatDataHeader2);

            File.AppendAllText(userLogPath, userLogCsvContent.ToString());
            userLogCsvContent.Clear();
            string concatUserStats2 = "";
            StringBuilder userLogCsvContent2 = new StringBuilder();
            List<object> userStatsList2 = new List<object>();

            #endregion CSV-Export Part1

            foreach (var pqm in PQMs.PQMs)
            {
                if (pqm.Ignore.ToString() != "True")
                {

                    var row = table.AddRow();
                    row.Format.Font.Size = 10;

                    if (pqm.Met.ToUpper() == "NOT MET")
                    {
                        row.Cells[5].Format.Font.Color = Color.FromRgb(220, 20, 6);
                        row.Cells[5].Format.Font.Bold = true;
                    }
                    if (pqm.Met == "Variation" && pqm.TemplateId.ToString().ToUpper() != "TARGET" && pqm.TemplateId.ToString().ToUpper() != "BODY")
                    {
                        row.Cells[6].Format.Font.Color = Color.FromRgb(253, 106, 2);
                        row.Cells[6].Format.Font.Bold = true;

                        row.Cells[5].Format.Font.Color = Color.FromRgb(253, 106, 2);
                        row.Cells[5].Format.Font.Bold = true;
                    }

                    row.VerticalAlignment = VerticalAlignment.Center;

                    row.Cells[0].AddParagraph(pqm.TemplateId);
                    row.Cells[1].AddParagraph(pqm.StructureNameWithCode);
                    row.Cells[2].AddParagraph(pqm.StructVolume);
                    row.Cells[3].AddParagraph(pqm.DVHObjective);
                    row.Cells[4].AddParagraph(pqm.Goal + " (" + pqm.Variation + ")");
                    //row.Cells[5].AddParagraph(pqm.Variation);
                    row.Cells[5].AddParagraph(pqm.Achieved);
                    if (reportPlanningItem.TreatmentApproverDisplayName != "" & !reportPlanningItem.TreatmentApproverDisplayName.ToString().ToUpper().Contains("NOT APPROVED") && pqm.Met.ToUpper() == "NOT MET")
                    {
                        row.Cells[6].Format.Font.Color = Color.FromRgb(0, 128, 0);
                        row.Cells[6].Format.Font.Bold = true;
                        row.Cells[6].AddParagraph("Accepted");
                    }
                    else if (reportPlanningItem.TreatmentApproverDisplayName == "" && pqm.Met.ToUpper() == "NOT MET")
                    {
                        row.Cells[6].Format.Font.Color = Color.FromRgb(220, 20, 6);
                        row.Cells[6].Format.Font.Bold = true;
                        row.Cells[6].AddParagraph(pqm.Met);
                    }
                    else if (reportPlanningItem.TreatmentApproverDisplayName != "" && reportPlanningItem.TreatmentApproverDisplayName.ToString().ToUpper().Contains("NOT APPROVED") && pqm.Met.ToUpper() == "NOT MET")
                    {
                        row.Cells[6].Format.Font.Color = Color.FromRgb(220, 20, 6);
                        row.Cells[6].Format.Font.Bold = true;
                        row.Cells[6].AddParagraph(pqm.Met);
                    }
                    else
                    {
                        row.Cells[6].AddParagraph(pqm.Met);
                    }
                    #region CSV-Report Part 2

                    //PQMs
                    userStatsList2.Clear();

                        userStatsList2.Add(pqm.TemplateId);
                        userStatsList2.Add(pqm.StructureNameWithCode);
                        userStatsList2.Add(pqm.StructVolume);
                        userStatsList2.Add(pqm.DVHObjective);
                        userStatsList2.Add(pqm.Goal);
                        userStatsList2.Add(pqm.Variation);
                        userStatsList2.Add(pqm.Achieved);
                        userStatsList2.Add(pqm.Met);

                    
                        concatUserStats2 = string.Join(";", userStatsList2.ToArray());

                        userLogCsvContent2.AppendLine(concatUserStats2);

                        
                    
                    #endregion
                }
                
            }
            File.AppendAllText(userLogPath, userLogCsvContent2.ToString());
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
