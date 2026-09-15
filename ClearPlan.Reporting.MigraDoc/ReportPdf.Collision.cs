using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClearPlan.Core.Review;
using MigraDoc.DocumentObjectModel;

namespace ClearPlan.Reporting.MigraDoc
{
    public partial class ReportPdf
    {
        private void AddCollisionSweeps(Section section, IList<CollisionReportBeam> beams)
        {
            if (beams.Count == 0) return;
            var minimum = CollisionReportBeam.MinimumBeam(beams);
            section.AddPageBreak();
            AddHeading(section, "Collision / 3D | Minimum model distance");
            var heading = section.AddParagraph(minimum == null ? "Minimum model distance unavailable."
                : minimum.MinimumCaption + " · " + CollisionLabel(minimum.MinimumRow.Status));
            heading.Format.Font.Bold = true; heading.Format.Font.Size = 11;
            heading.Format.SpaceAfter = Unit.FromMillimeter(3);
            heading.Format.Shading.Color = CollisionColor(minimum == null ? "uncertain" : minimum.MinimumRow.Status);
            var panel = CreateTable(section, 15.4, 11.9); panel.Borders.Visible = false;
            var row = panel.AddRow();
            if (minimum != null && minimum.OverviewPng != null && minimum.OverviewPng.Length > 0)
            {
                var image = row.Cells[0].AddImage("base64:" + Convert.ToBase64String(minimum.OverviewPng));
                image.Width = Unit.FromCentimeter(15); image.LockAspectRatio = true;
            }
            else row.Cells[0].AddParagraph("Minimum-position image unavailable; no substitute pose shown.");
            if (minimum != null && minimum.OrientationPng != null && minimum.OrientationPng.Length > 0)
            {
                var strip = row.Cells[0].AddParagraph();
                strip.Format.SpaceBefore = Unit.FromMillimeter(2);
                var orientation = strip.AddImage("base64:" + Convert.ToBase64String(minimum.OrientationPng));
                orientation.Width = Unit.FromCentimeter(12); orientation.LockAspectRatio = true;
            }
            else row.Cells[0].AddParagraph("Orientation strip unavailable; no substitute pose shown.");
            var summary = row.Cells[0].AddParagraph(minimum == null ? "No finite conservative distance available." : minimum.Summary);
            summary.Format.Font.Size = 8;
            var table = row.Cells[1].Elements.AddTable();
            table.Borders.Width = 0.25; table.Borders.Color = Color.FromRgb(212,225,233);
            table.Format.Font.Size = 7.5; table.TopPadding = Unit.FromMillimeter(0.7); table.BottomPadding = Unit.FromMillimeter(0.7);
            table.LeftPadding = Unit.FromMillimeter(0.6); table.RightPadding = Unit.FromMillimeter(0.6);
            foreach (double width in new[] {1.35,1.0,1.0,2.0,2.5,2.5,1.45}) table.AddColumn(Unit.FromCentimeter(width));
            AddHeader(table, "Field", "Gantry", "Couch", "Status¹", "Body [mm]", "Table [mm]", "Min. [mm]");
            foreach (var beam in beams)
            {
                var point = beam.MinimumRow;
                var entry = table.AddRow();
                SetCells(entry, beam.BeamId, point == null ? "—" : point.GantryDegrees.ToString("0.#", CultureInfo.InvariantCulture) + "°",
                    point == null ? "—" : point.CouchDegrees.ToString("0.#", CultureInfo.InvariantCulture) + "°", CollisionLabel(beam.Status),
                    point == null ? "—" : point.BodyDistance ?? "—", point == null ? "—" : point.TableDistance ?? "—",
                    point == null ? "—" : point.MinimumDistanceMm.Value.ToString("0.0", CultureInfo.InvariantCulture));
                entry.Cells[3].Shading.Color = CollisionColor(beam.Status);
                if (ReferenceEquals(beam, minimum)) entry.Format.Font.Bold = true;
            }
            var key = row.Cells[1].AddParagraph("¹ Status covers all sampled positions in the field; distances and angles identify its minimum. 10° states plus endpoints. " +
                "Body/table: conservative lower bound … sampled upper bound. Min.: smallest finite lower bound (open bore: radial gap). " +
                "Green / Full clear*: captured body and table clear. Yellow / Still clear*: only part of the geometry assessed clear; see body/table results. Uncertain/Hit remain distinct. " +
                "Not continuous-path or clinical clearance; outside-CT anatomy and actual setup remain unverified. Orientation strip: nominal Gantry/Couch angles at the minimum pose. Gray C-arm reference, when present: schematic, not tested geometry.");
            key.Format.Font.Size = 8; key.Format.SpaceBefore = Unit.FromMillimeter(3);
        }
        private static string CollisionLabel(string status) { return status == "body-clear" || status == "partial-clear" ? "Still clear*" : status == "model-clear" || status == "clear" || status == "pass" ? "Full clear*" : status == "model-hit" ? "Model hit" : status == "hit" ? "Hit" : "Uncertain"; }
        private static Color CollisionColor(string status)
        { return status == "pass" || status == "clear" || status == "model-clear" ? Color.FromRgb(218,242,224) : status == "hit" || status == "model-hit" ? Color.FromRgb(253,224,220) : Color.FromRgb(255,241,204); }
    }
}
