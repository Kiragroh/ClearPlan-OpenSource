using System;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;

namespace ClearPlan.Reporting.MigraDoc.Internal
{
    internal class PatientInfo
    {
        public static readonly Color Shading = new Color(243, 243, 243);

        public void Add(Section section, ReportPatient patient, ReportPlanningItem reportPlanningItem)
        {
            var table = AddPatientInfoTable(section);

            AddLeftInfo(table.Rows[0].Cells[0], patient);
            AddRightInfo(table.Rows[0].Cells[1], patient, reportPlanningItem);
        }

        private Table AddPatientInfoTable(Section section)
        {
            var table = section.AddTable();
            table.Shading.Color = Shading;

            table.Rows.LeftIndent = 0;

            table.LeftPadding = Size.TableCellPadding;
            table.TopPadding = Size.TableCellPadding;
            table.RightPadding = Size.TableCellPadding;
            table.BottomPadding = Size.TableCellPadding;

            // Use two columns of equal width
            var columnWidth = Size.GetWidth(section) / 2.0;
            table.AddColumn(columnWidth);
            table.AddColumn(columnWidth);

            // Only one row is needed
            table.AddRow();

            return table;
        }

        private void AddLeftInfo(Cell cell, ReportPatient patient)
        {
            // Add patient name and sex symbol
            var p1 = cell.AddParagraph();
            p1.Style = CustomStyles.PatientName;
            var myFont = new Font();
            myFont.Size = 11;
            myFont.Bold = true;
            //if (data.)
            if (patient.LastName.ToString() != "" && patient.FirstName.ToString() != "")
            {
                //p1.AddText($"{patient.LastName}, {patient.FirstName}");
                p1.AddFormattedText($"{patient.LastName}, {patient.FirstName}", myFont);
            }
            if (patient.LastName.ToString() == "")
            {
                //p1.AddText($"not assigned, {patient.FirstName}");
                p1.AddFormattedText($"not assigned, {patient.FirstName}", myFont);
            }
            if (patient.FirstName.ToString() == "")
            {
                //p1.AddText($"{patient.LastName}, not assigned");
                p1.AddFormattedText($"{patient.LastName}, not assigned", myFont);
            }
            

            // Add patient ID
            //var p2 = cell.AddParagraph();
            //p1.AddText(" (");

            p1.AddFormattedText(" ("+patient.Id+")", myFont);
            //p1.AddText(")");
            p1.AddSpace(1);
            if (patient.Sex.ToString().ToUpper() == "MALE" || patient.Sex.ToString().ToUpper() == "FEMALE" || patient.Sex.ToString().ToUpper() == "MÄNNLICH" || patient.Sex.ToString().ToUpper() == "WEIBLICH")
            {
                AddSexSymbol(p1, patient.Sex);
            }
        }

        private void AddSexSymbol(Paragraph p, Sex sex)
        {
            p.AddImage(new SexSymbol(sex).GetMigraDocFileName());
        }

        private void AddRightInfo(Cell cell, ReportPatient patient, ReportPlanningItem reportPlanningItem)
        {
            var p = cell.AddParagraph();
            
            // Add birthdate
            p.AddText("Birthdate: ");
            p.AddFormattedText(Format(patient.Birthdate), TextFormat.Bold);

            p.AddLineBreak();

            // Add TreatmentApprover
            p.AddText("TreatmentApproved by: ");
            if (reportPlanningItem.TreatmentApproverDisplayName != "" &! reportPlanningItem.TreatmentApproverDisplayName.StartsWith("1."))
            {
                p.AddFormattedText($"\n{reportPlanningItem.TreatmentApproverDisplayName} at {reportPlanningItem.treatmentapproveDate}", TextFormat.Bold);
            }
            //PlanSum
            else if (reportPlanningItem.TreatmentApproverDisplayName.StartsWith("1."))
            {
                p.AddFormattedText($" {reportPlanningItem.PlansInPlansum} Plans in PlanSum");
                p.AddLineBreak();
                p.AddFormattedText($"{reportPlanningItem.TreatmentApproverDisplayName}", TextFormat.Bold);
            }
            else if (reportPlanningItem.TreatmentApproverDisplayName == "")
            {
                p.AddFormattedText($" No TreatmentApproval", TextFormat.Bold);
            }
            else
            {
                p.AddFormattedText($" {reportPlanningItem.TreatmentApproverDisplayName}", TextFormat.Bold);
            }
        }

        private string Format(DateTime birthdate)
        {
            return $"{birthdate:d} (age {Age(birthdate)})";
        }

        // See http://stackoverflow.com/a/1404/1383366
        private int Age(DateTime birthdate)
        {
            var today = DateTime.Today;
            int age = today.Year - birthdate.Year;
            return birthdate.AddYears(age) > today ? age : age;
        }
    }
}
