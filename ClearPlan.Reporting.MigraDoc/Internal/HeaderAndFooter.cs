using System;
using MigraDoc.DocumentObjectModel;

namespace ClearPlan.Reporting.MigraDoc.Internal
{
    internal class HeaderAndFooter
    {
        public void Add(Section section, ReportData data)
        {
            AddHeader(section, data.ReportPatient);
            AddFooter(section, data.ReportPatient);
        }

        private void AddHeader(Section section, ReportPatient patient)
        {
            var header = section.Headers.Primary.AddParagraph();
            header.Format.AddTabStop(Size.GetWidth(section), TabAlignment.Right);

            header.AddText($"{patient.LastName}, {patient.FirstName} (ID: {patient.Id}, {patient.planType}: {patient.planId})");
            header.AddTab();
            //header.AddText($"Generated {DateTime.Now:g}");
        }

        private void AddFooter(Section section, ReportPatient patient)
        {
            var footer = section.Footers.Primary.AddParagraph();
            footer.Format.AddTabStop(Size.GetWidth(section), TabAlignment.Right);

            footer.AddText($"PQM-Report (Generated from {patient.userId} on {Environment.MachineName} - {DateTime.Now:g})");
            footer.AddTab();
            footer.AddText("Page ");
            footer.AddPageField();
            footer.AddText(" of ");
            footer.AddNumPagesField();
        }
    }
}
