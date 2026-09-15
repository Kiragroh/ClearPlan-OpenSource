using System;
using System.Collections;
using System.Linq;
using System.Reflection;
using ClearPlan.Core.Review;
using Newtonsoft.Json;

namespace ClearPlan.Core.Tests
{
    internal static class IsodoseDisplayTests
    {
        private static Type ConfigurationType()
        {
            var type = typeof(ReviewPlanImage).Assembly.GetType("ClearPlan.Core.Review.IsodoseDisplayConfiguration");
            TestAssert.NotNull(type, "Missing shared, validated isodose display palette.");
            return type;
        }
        private static dynamic Defaults() { return ConfigurationType().GetMethod("CreateDefault").Invoke(null, null); }
        private static dynamic Parse(string json) { return ConfigurationType().GetMethod("Parse").Invoke(null, new object[] { json }); }
        private static ReviewPlanImage Apply(ReviewPlanImage image, object config)
        { return (ReviewPlanImage)ConfigurationType().GetMethod("Apply").Invoke(config, new object[] { image }); }
        internal static ReviewPlanImage Plane()
        {
            return JsonConvert.DeserializeObject<ReviewPlanImage>(@"{
              'PlanKey':'p','Kind':'transversal','SourceStatus':'available',
              'WidthPixels':3,'HeightPixels':3,'PixelSpacingXMillimeters':1,'PixelSpacingYMillimeters':1,
              'GrayscalePixels':'AAAAAAAAAAAA','DosePlane':{'Columns':3,'Rows':3,'PrescriptionGy':50,
                'SamplesGy':[0,53,0,0,53,0,0,53,0],'Source':'Synthetic test plane'}}");
        }
        public static void DefaultsAndValidation()
        {
            dynamic defaults = Defaults();
            var levels = ((IEnumerable)defaults.Levels).Cast<dynamic>().ToList();
            foreach (var pair in new[] { new { P=95.0, C="#38D46A" }, new { P=100.0, C="#8DE4C1" },
                new { P=105.0, C="#FFD84A" }, new { P=107.0, C="#F58AB6" }, new { P=110.0, C="#F44336" }, new { P=125.0, C="#B56CFF" } })
                TestAssert.Equal(pair.C, (string)levels.Single(l => (double)l.Percent == pair.P).ColorHex);
            Reject("{'SchemaVersion':1,'Levels':[{'Percent':95,'ColorHex':'#38D46A'},{'Percent':95,'ColorHex':'#8DE4C1'}]}");
            Reject("{'SchemaVersion':1,'Levels':[{'Percent':0,'ColorHex':'#38D46A'}]}");
            Reject("{'SchemaVersion':1,'Levels':[{'Percent':95,'ColorHex':'transparent'}]}");
            Reject("{'SchemaVersion':2,'Levels':[]}");
            Reject("{'Levels':[]}");
            Reject("{'SchemaVersion':1}");
            Reject("{'SchemaVersion':1,'Levels':null}");
            foreach (double retained in new[] { 2.0, 20.0, 50.0, 80.0 })
                TestAssert.True(levels.Any(l => (double)l.Percent == retained && (bool)l.Enabled));
        }
        private static void Reject(string json)
        {
            bool rejected = false;
            try { Parse(json); } catch (TargetInvocationException ex) { rejected = ex.InnerException is FormatException; }
            TestAssert.True(rejected, "Invalid palette must fail atomically.");
        }
        public static void GuiEditsAndExportUseSamePalette()
        {
            var source = Plane();
            dynamic vm = new ClearPlan.Presentation.ViewModels.PlanImagesViewModel(new[] { source }, "p");
            var property = vm.GetType().GetProperty("IsodoseEditorRows");
            TestAssert.NotNull((object)property, "CT GUI needs a level/color editor.");
            dynamic row = ((IEnumerable)vm.IsodoseEditorRows).Cast<dynamic>().Single(l => (string)l.PercentText == "95");
            row.PercentText = "98"; row.ColorHex = "#123456";
            vm.ApplyIsodosesCommand.Execute(null);
            TestAssert.Equal("#123456", ((System.Collections.Generic.List<ReviewPlanImage>)vm.ApplyToReport(new[] { source })).Single().Overlays.Single(o => o.DoseGy == 49).ColorHex);
            string accepted = JsonConvert.SerializeObject((object)vm.IsodoseConfiguration);
            row.PercentText = "bad";
            vm.ApplyIsodosesCommand.Execute(null);
            TestAssert.Equal(accepted, JsonConvert.SerializeObject((object)vm.IsodoseConfiguration), "Invalid drafts must leave last applied palette unchanged.");
            TestAssert.True(((string)vm.IsodoseEditorMessage).Length > 0);
            var snapshot = ClearPlan.Core.Simulation.SyntheticScenarioFactory.Create("baseline-pass");
            source.PlanKey = snapshot.ActivePlanKey; snapshot.PlanImages = new System.Collections.Generic.List<ReviewPlanImage> { source };
            dynamic dynamicSnapshot = snapshot;
            dynamicSnapshot.IsodoseDisplay = vm.IsodoseConfiguration;
            var report = new ClearPlan.Reporting.ReviewSnapshotReportMapper().Map(snapshot);
            TestAssert.Equal("#123456", report.PlanImages.Single().Overlays.Single(o => o.DoseGy == 49).ColorHex);
            TestAssert.True(report.PlanImages.Single().DosePlane != source.DosePlane);
            TestAssert.True(source.Overlays.Count == 0);
            // Same-plan reload preserves deliberate edits, while Loaded defaults
            // adopts the newly read revision instead of silently overwriting edits.
            dynamic reloaded = new ClearPlan.Presentation.ViewModels.PlanImagesViewModel(new[] { source }, source.PlanKey);
            dynamic newDefaults = Defaults();
            foreach (dynamic level in (IEnumerable)newDefaults.Levels)
                if ((double)level.Percent == 95) level.ColorHex = "#654321";
            reloaded.SetIsodoseConfiguration(newDefaults, "new revision", true);
            reloaded.SetIsodoseConfiguration(vm.IsodoseConfiguration, "session restored", false);
            TestAssert.Equal(accepted, JsonConvert.SerializeObject((object)reloaded.IsodoseConfiguration));
            reloaded.ResetIsodosesCommand.Execute(null);
            TestAssert.True(((IEnumerable)reloaded.IsodoseConfiguration.Levels).Cast<dynamic>()
                .Any(l => (double)l.Percent == 95 && (string)l.ColorHex == "#654321"));
            int count = reloaded.IsodoseEditorRows.Count;
            reloaded.AddIsodoseCommand.Execute(null);
            TestAssert.Equal(count + 1, (int)reloaded.IsodoseEditorRows.Count);
            dynamic added = reloaded.SelectedIsodose;
            added.PercentText = "91"; added.ColorHex = "#556677"; added.Enabled = false;
            reloaded.ApplyIsodosesCommand.Execute(null);
            TestAssert.False(((System.Collections.Generic.List<ReviewPlanImage>)reloaded.ApplyToReport(new[] { source })).Single().Overlays.Any(o => o.PrescriptionPercent == 91));
            added.Enabled = true; reloaded.ApplyIsodosesCommand.Execute(null);
            TestAssert.True(((System.Collections.Generic.List<ReviewPlanImage>)reloaded.ApplyToReport(new[] { source })).Single().Overlays.Any(o => o.PrescriptionPercent == 91 && o.ColorHex == "#556677"));
            reloaded.RemoveIsodoseCommand.Execute(null); reloaded.ApplyIsodosesCommand.Execute(null);
            TestAssert.Equal(count, (int)reloaded.IsodoseEditorRows.Count);
            TestAssert.False(((System.Collections.Generic.List<ReviewPlanImage>)reloaded.ApplyToReport(new[] { source })).Single().Overlays.Any(o => o.PrescriptionPercent == 91));
        }
        public static void SettingsPathAndValidation()
        {
            var settings = new ClearPlan.Core.Settings.ClearPlanSettingsModel();
            ClearPlan.Core.Settings.PathSettingsIni.Apply(settings, "[Paths]\nIsodoseDisplayJsonPath=Configuration/palette.json");
            TestAssert.True(ClearPlan.Core.Settings.PathSettingsIni.Serialize(settings).Contains("IsodoseDisplayJsonPath = Configuration/palette.json"));
            ClearPlan.ConfigurationContentValidator.Validate("isodose-display", System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject((object)Defaults())), null);
            TestAssert.Throws<FormatException>(() => ClearPlan.ConfigurationContentValidator.Validate("isodose-display", System.Text.Encoding.UTF8.GetBytes("{'Levels':[{'Percent':-1}]}"), null));
        }
        public static void RenderedScaleShowsOnlyPresentLevels()
        {
            var type = typeof(ClearPlan.Rendering.PlanImageRenderer);
            var method = type.GetMethod("IsodoseScale");
            TestAssert.NotNull(method, "Shared image renderer must expose the exact displayed isodose scale.");
            var source = Apply(Plane(), (object)Defaults());
            var scale = ((IEnumerable)method.Invoke(null, new object[] { source })).Cast<ReviewImageOverlay>().ToList();
            TestAssert.True(scale.Any(l => l.DoseGy == 52.5));
            TestAssert.False(scale.Any(l => l.DoseGy >= 53.5));
            TestAssert.NotNull(typeof(ClearPlan.Presentation.ViewModels.PlanImagesViewModel).GetProperty("IsodoseScaleItems"),
                "GUI scale must be native readable text, not only shrinking raster labels.");
            var second = Apply(Plane(), (object)Defaults()); second.Kind = "coronal";
            var vm = new ClearPlan.Presentation.ViewModels.PlanImagesViewModel(new[] { source, second }, "p");
            TestAssert.Equal(scale.Count, vm.IsodoseScaleItems.Count(), "Shared scale must deduplicate identical levels across planes.");
            vm.ShowDose = false;
            TestAssert.False(vm.HasIsodoseScale, "Hidden isodoses must not leave a stale common scale.");
            using (var stream = new System.IO.MemoryStream(ClearPlan.Rendering.PlanImageRenderer.Render(source, true, true, true)))
            using (var bitmap = new System.Drawing.Bitmap(stream))
            {
                TestAssert.True(bitmap.Height > 1440, "Scale belongs below CT, not over anatomy.");
                int green = 0;
                for (int y = 1440; y < bitmap.Height; y++)
                for (int x = 0; x < bitmap.Width; x++)
                    if (bitmap.GetPixel(x, y).ToArgb() == System.Drawing.ColorTranslator.FromHtml("#38D46A").ToArgb()) green++;
                TestAssert.True(green > 40, "95% green scale swatch must be rendered.");
            }
        }
        public static void PlaneLevelsAreRealAndDetached()
        {
            var source = Plane(); var image = Apply(source, (object)Defaults());
            var dose = image.Overlays.Where(o => o.Kind == "isodose").ToList();
            TestAssert.True(dose.Any(o => o.DoseGy == 52.5 && o.ColorHex == "#FFD84A"));
            TestAssert.False(dose.Any(o => o.DoseGy >= 53.5), "Absent high-dose levels must not appear.");
            TestAssert.Equal(0, source.Overlays.Count, "Display edits must not mutate the capture.");
            image.GrayscalePixels[0] = 99;
            TestAssert.Equal((byte)0, source.GrayscalePixels[0]);
            dynamic changed = Parse("{'SchemaVersion':1,'Levels':[{'Percent':98,'ColorHex':'#123456','Enabled':true}]}");
            var edited = Apply(source, (object)changed);
            TestAssert.Equal(1, edited.Overlays.Count);
            TestAssert.Equal(49.0, edited.Overlays[0].DoseGy.Value);
            TestAssert.Equal("#123456", edited.Overlays[0].ColorHex);
            var noDose = Apply(new ReviewPlanImage(), (object)Defaults());
            TestAssert.False(noDose.Overlays.Any(o => o.Kind == "isodose" && o.SourceStatus == "available"));
            var highDose = Plane();
            highDose.DosePlane.SamplesGy = new[] { 0.0,70,0,0,70,0,0,70,0 };
            var high = Apply(highDose, (object)Defaults());
            TestAssert.True(high.Overlays.Any(o => o.PrescriptionPercent == 110 && o.ColorHex == "#F44336" && o.Paths.Count > 0));
            TestAssert.True(high.Overlays.Any(o => o.PrescriptionPercent == 125 && o.ColorHex == "#B56CFF" && o.Paths.Count > 0));
        }
    }
}
