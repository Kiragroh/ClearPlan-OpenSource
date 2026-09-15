using System;
using System.Collections;
using System.Linq;
using ClearPlan.Core.Collision;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;
using ClearPlan.Presentation.Views;
using ClearPlan.Reporting;
using ClearPlan.Reporting.MigraDoc;

namespace ClearPlan.Core.Tests
{
    internal static class CollisionSweepReportTests
    {
        public static void OrientationStripIsSeparateBoundToMinimumAndDetached()
        {
            var property = typeof(CollisionReportBeam).GetProperty("OrientationPng");
            TestAssert.NotNull(property, "Reports need a separate orientation strip below the overview.");
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            var scene = SyntheticCollisionFactory.Create(snapshot.ActivePlanKey); scene.PatientPosition = "HFS";
            scene.Poses = scene.Poses.Take(3).ToList();
            var sweep = new CollisionSweepResult { Frames = scene.Poses.Select((p,i) => new CollisionSweepFrame {
                Pose=p, Status="pass", BodyStatus="pass", TableStatus="pass", BodyLowerBoundMm=20-i,
                BodyClearanceMm=20-i, TableLowerBoundMm=40, TableClearanceMm=40 }).ToList() };
            var model = new CollisionViewModel(snapshot); model.SetScene(scene,null,sweep);
            model.IncludeInReport=true; model.PosePosition=0;
            CollisionView.CaptureReportSweep(model);
            var captured=snapshot.CollisionBeams.Single();
            var report=new ReviewSnapshotReportMapper().Map(snapshot);
            TestAssert.False(ReviewSnapshotJson.Serialize(snapshot).Contains("OrientationPng"),"No private strip in public JSON.");
            var strip=(byte[])property.GetValue(captured);
            TestAssert.True(strip != null && strip.Length > 1000, "Minimum pose needs a separately captured, nonempty strip; bytes="+(strip==null ? 0 : strip.Length));
            TestAssert.Equal(0,model.PosePosition,"Capture must restore the user's selected pose.");
            TestAssert.Equal(10d,captured.MinimumRow.GantryDegrees);
            model.PosePosition=2;
            var view=new CollisionView { DataContext=model };
            view.Measure(new System.Windows.Size(1340,800)); view.Arrange(new System.Windows.Rect(0,0,1340,800)); view.UpdateLayout();
            var flags=System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance;
            typeof(CollisionView).GetField("captureRequested",flags).SetValue(view,true);
            typeof(CollisionView).GetMethod("DrawScene",flags).Invoke(view,null);
            view.UpdateLayout();
            var capture=typeof(CollisionView).GetMethod("CapturePng",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Static);
            var root=(System.Windows.FrameworkElement)view.FindName("RoomReferenceRoot");
            var expected=(byte[])capture.Invoke(null,new object[]{root});
            TestAssert.True(strip.SequenceEqual(expected),"Orientation must show the minimum pose, not the restored GUI pose.");
            root.Visibility=System.Windows.Visibility.Hidden; view.UpdateLayout();
            var expectedOverview=(byte[])capture.Invoke(null,new object[]{view.FindName("IllustrationRoot")});
            TestAssert.True(captured.OverviewPng.SequenceEqual(expectedOverview),"Main report image must exclude the orientation overlay.");
            using(var stream=new System.IO.MemoryStream(strip))
            using(var pixels=new System.Drawing.Bitmap(stream))
            {
                TestAssert.True(pixels.Width>pixels.Height*2.8,"PDF/HTML keep a horizontal strip, independent of the GUI's vertical rail.");
                int light=0,teal=0;
                for(int y=0;y<pixels.Height;y++) for(int x=0;x<pixels.Width;x++)
                { var color=pixels.GetPixel(x,y); if(color.R>170 && color.G>170 && color.B>170)light++; if(color.G>120 && color.G>color.R*1.25 && color.B>color.R*1.1)teal++; }
                TestAssert.True(light>300 && teal>100,"Orientation strip needs rendered labels, gantry and couch geometry, not transparent pixels.");
            }
            var copied=(byte[])property.GetValue(report.CollisionBeams.Single());
            TestAssert.False(ReferenceEquals(strip,copied),"Report strip must be deep-copied.");
            TestAssert.True(strip.SequenceEqual(copied));
            var pdf=new ReportPdf();
            var pdfDocument=(MigraDoc.DocumentObjectModel.Document)typeof(ReportPdf).GetMethod("CreateReviewReport",flags).Invoke(pdf,new object[]{report});
            string ddl=MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToString(pdfDocument);
            int mainIndex=ddl.IndexOf(Convert.ToBase64String(captured.OverviewPng),StringComparison.Ordinal);
            int stripIndex=ddl.IndexOf(Convert.ToBase64String(strip),StringComparison.Ordinal);
            TestAssert.True(mainIndex>=0 && stripIndex>mainIndex,"PDF must place the same orientation image after the main view.");
            string html=new HtmlReviewReportRenderer().Render(report);
            string collision=html.Substring(html.IndexOf("id=\"collision\"",StringComparison.Ordinal));
            collision=collision.Substring(0,collision.IndexOf("</section>",StringComparison.Ordinal));
            TestAssert.Equal(2,System.Text.RegularExpressions.Regex.Matches(collision,"<img ").Count);
            TestAssert.True(collision.IndexOf("Minimum model distance position",StringComparison.Ordinal) <
                collision.IndexOf("Gantry and couch orientation",StringComparison.Ordinal),"Orientation belongs below the main image.");
            property.SetValue(report.CollisionBeams.Single(),null);
            string legacy=new HtmlReviewReportRenderer().Render(report);
            TestAssert.True(legacy.Contains("Orientation strip unavailable; no substitute pose shown."),"Old sidecars must not fabricate orientation.");
            var roundTrip=Newtonsoft.Json.JsonConvert.DeserializeObject<CollisionReportBeam>(Newtonsoft.Json.JsonConvert.SerializeObject(captured));
            TestAssert.True(strip.SequenceEqual((byte[])property.GetValue(roundTrip)),"Private collision sidecar roundtrip must retain strip.");
        }

        public static void UnavailableSweepDoesNotExportAnEmptyOverview()
        {
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            var model = new CollisionViewModel(snapshot);
            var scene = SyntheticCollisionFactory.Create(snapshot.ActivePlanKey);
            model.SetScene(scene, null, new CollisionSweepResult { Status = "uncertain", Summary = "Workload unavailable." });
            model.IncludeInReport = true;
            CollisionView.CaptureReportSweep(model);
            TestAssert.True(snapshot.CollisionBeams == null || snapshot.CollisionBeams.Count == 0, "No empty, apparently evaluated beam report is allowed.");
            TestAssert.True(snapshot.CollisionPreviewPng == null, "No pose-less image may masquerade as a evaluated collision overview.");
        }

        public static void ReportGetsEachBeamOverviewStatusAndDistances()
        {
            var capture = typeof(CollisionView).GetMethod("CaptureReportSweep");
            TestAssert.NotNull(capture, "Capture each beam's ten-degree sweep, not one last GUI control point.");
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            var model = new CollisionViewModel(snapshot);
            var scene = SyntheticCollisionFactory.Create(snapshot.ActivePlanKey);
            scene.Poses = scene.Poses.Take(5).ToList();
            scene.Poses.Add(new CollisionPose { BeamId = "SECOND <beam>", ControlPointIndex = 0, GantryDegrees = 0,
                Isocenter = scene.Poses[0].Isocenter, Source = scene.Poses[0].Source });
            var sweep = new CollisionSweepResult { Frames = scene.Poses.Select((p,i) => new CollisionSweepFrame {
                Pose=p, Status="pass", BodyStatus="pass", TableStatus="pass", BodyLowerBoundMm=20-i,
                BodyClearanceMm=20-i, TableLowerBoundMm=40, TableClearanceMm=40 }).ToList() };
            model.SetScene(scene, null, sweep); model.IncludeInReport = true;
            model.SelectedBeamId=scene.Poses[0].BeamId; model.PosePosition = 1;
            capture.Invoke(null, new object[] { model });
            TestAssert.Equal(1, model.PosePosition, "Export must restore GUI selection.");
            var mapper = new ReviewSnapshotReportMapper();
            var report = mapper.Map(snapshot);
            var property = typeof(ReviewReportDocument).GetProperty("CollisionBeams");
            TestAssert.NotNull(property);
            var beams = (IList)property.GetValue(report);
            TestAssert.Equal(2, beams.Count);
            var typed = report.CollisionBeams;
            TestAssert.Equal("SECOND <beam>", CollisionReportBeam.MinimumBeam(typed).BeamId);
            TestAssert.Equal(15d, CollisionReportBeam.MinimumBeam(typed).MinimumRow.MinimumDistanceMm.Value);
            TestAssert.True(typed.All(b => b.Rows.All(r => typeof(CollisionReportRow).GetProperty("MinimumDistanceMm") != null)),
                "Report minimum must use typed distances, not parse localized interval text.");
            string html = new HtmlReviewReportRenderer().Render(report);
            TestAssert.True(html.Contains("SECOND &lt;beam&gt;"), "Beam names must remain HTML encoded.");
            string visible = System.Net.WebUtility.HtmlDecode(html);
            TestAssert.True(visible.Contains("Minimum model distance"));
            string collision = html.Substring(html.IndexOf("id=\"collision\"", StringComparison.Ordinal));
            collision = collision.Substring(0, collision.IndexOf("</section>", StringComparison.Ordinal));
            TestAssert.Equal(1, System.Text.RegularExpressions.Regex.Matches(collision, "<img ").Count,
                "One global-minimum illustration, not one figure per field.");
            foreach (string value in new[] { "10°", "Body [mm]", "Table [mm]", "Status", "Uncertain", "Hit", "Full clear", "Still clear" })
                TestAssert.True(visible.Contains(value), "Collision report missing " + value);
            TestAssert.True(collision.Contains("background:#daf2e0;color:#202b38\">Full clear*"));
            report.CollisionBeams[0].Status = "body-clear";
            report.CollisionBeams[1].Status = "partial-clear";
            var limited = new HtmlReviewReportRenderer().Render(report);
            TestAssert.Equal(2, System.Text.RegularExpressions.Regex.Matches(limited, "background:#fff1cc;color:#202b38\">Still clear\\*").Count);
            foreach(var status in new[]{"body-clear","partial-clear"})
            {
                var color = typeof(ReportPdf).GetMethod("CollisionColor", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic).Invoke(null,new object[]{status});
                TestAssert.Equal(MigraDoc.DocumentObjectModel.Color.FromRgb(255,241,204), color, "PDF still-clear rows must be yellow.");
            }
            model.IncludeInReport = false;
            TestAssert.Equal(0, ((IList)property.GetValue(mapper.Map(snapshot))).Count);
            TestAssert.False(ReviewSnapshotJson.Serialize(snapshot).Contains("CollisionBeams"));
            model.IncludeInReport = true;
            model.SetFailure("Changed plan");
            TestAssert.Equal(0, ((IList)property.GetValue(mapper.Map(snapshot))).Count);
        }
    }
}
