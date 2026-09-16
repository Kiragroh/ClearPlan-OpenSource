using System;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class ReviewLanguageTests
    {
        public static void ExplicitBilingualLabelScopes()
        {
            string original = ClearPlan.Core.Localization.ReviewLanguage.Code;
            try
            {
                ClearPlan.Core.Localization.ReviewLanguage.Code = "de";
                TestAssert.Equal("Meldung", ClearPlan.Core.Localization.ReviewLanguage.Label("Meldung", "Message"));
                using (ClearPlan.Core.Localization.ReviewLanguage.Scope("en-US"))
                {
                    TestAssert.Equal("Message", ClearPlan.Core.Localization.ReviewLanguage.Label("Meldung", "Message"));
                    using (ClearPlan.Core.Localization.ReviewLanguage.Scope("de-DE"))
                        TestAssert.Equal("Meldung", ClearPlan.Core.Localization.ReviewLanguage.Label("Meldung", "Message"));
                    TestAssert.Equal("Message", ClearPlan.Core.Localization.ReviewLanguage.Label("Meldung", "Message"));
                }
                TestAssert.Equal("de", ClearPlan.Core.Localization.ReviewLanguage.Code);
            }
            finally { ClearPlan.Core.Localization.ReviewLanguage.Code = original; }
        }

        public static void GuiSelectorRoundTripPreservesIdentity()
        {
            // The real selector persists its choice. Restore the user's exact preference afterward.
            string preference = ClearPlan.Core.Localization.ReviewLanguage.PreferencePath;
            bool existed = System.IO.File.Exists(preference);
            byte[] saved = existed ? System.IO.File.ReadAllBytes(preference) : null;
            var language = ClearPlan.Presentation.Localization.DisplayLanguage.Instance;
            string original = ClearPlan.Core.Localization.ReviewLanguage.Code;
            try
            {
                var view = new ClearPlan.Presentation.Views.ReviewWorkspaceView();
                var selector = (System.Windows.Controls.ComboBox)view.FindName("ReviewLanguageSelector");
                TestAssert.NotNull(selector, "The review workspace must expose the real language selector.");
                var label = (System.Windows.Controls.TextBlock)System.Windows.Markup.XamlReader.Parse(
                    "<TextBlock xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' xmlns:l='clr-namespace:ClearPlan.Presentation.Localization;assembly=ClearPlan.Presentation' Text='{l:Translate Value=Schnittbilder}'/>");
                const string clinicalId = "PTV_Strukturen_ä_01";
                var identity = new System.Windows.Controls.TextBlock();
                identity.SetBinding(System.Windows.Controls.TextBlock.TextProperty, new System.Windows.Data.Binding { Source = clinicalId });
                foreach (string code in new[] { "en", "de", "en", "de" })
                {
                    selector.SelectedItem = code;
                    selector.GetBindingExpression(System.Windows.Controls.Primitives.Selector.SelectedItemProperty).UpdateSource();
                    System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(System.Windows.Threading.DispatcherPriority.DataBind, new Action(() => { }));
                    TestAssert.Equal(code, ClearPlan.Core.Localization.ReviewLanguage.Code);
                    TestAssert.Equal(code == "en" ? "CT views" : "Schnittbilder", label.Text);
                    TestAssert.Equal(clinicalId, identity.Text, "Language selection must not change identity bindings.");
                }
            }
            finally
            {
                ClearPlan.Core.Localization.ReviewLanguage.Code = original;
                if (existed) System.IO.File.WriteAllBytes(preference, saved);
                else if (System.IO.File.Exists(preference)) System.IO.File.Delete(preference);
            }
        }

        public static void HtmlAndPdfLanguageRoundTripPreservesClinicalValues()
        {
            string original = ClearPlan.Core.Localization.ReviewLanguage.Code;
            var report = LanguageReportFixture();
            try
            {
                foreach (string code in new[] { "en", "de", "en" })
                {
                    // The captured report language must win over an opposite current GUI language.
                    string guiCode = code == "en" ? "de" : "en";
                    ClearPlan.Core.Localization.ReviewLanguage.Code = guiCode;
                    report.LanguageCode = code;
                    string html = System.Net.WebUtility.HtmlDecode(new ClearPlan.Reporting.MigraDoc.HtmlReviewReportRenderer().Render(report));
                    TestAssert.True(html.Contains("<html lang=\"" + code + "\">"));
                    TestAssert.True(html.Contains("<h1>Plan Quality Report</h1>"));
                    TestAssert.True(html.Contains("<h2>" + (code == "en" ? "Clinical goals" : "Klinische Ziele") + "</h2>"));
                    TestAssert.True(html.Contains("<th scope=\"col\">" + (code == "en" ? "Message" : "Meldung") + "</th>"));
                    TestAssert.True(html.Contains("class=\"status pass\">" + (code == "en" ? "Met" : "Erfüllt") + "</span>"));
                    TestAssert.True(html.Contains("class=\"status fail\">" + (code == "en" ? "Not met" : "Nicht erfüllt") + "</span>"));
                    AssertReportLanguageIdentities(html);
                    TestAssert.Equal(guiCode, ClearPlan.Core.Localization.ReviewLanguage.Code, "HTML scope must restore the GUI language.");

                    var create = typeof(ClearPlan.Reporting.MigraDoc.ReportPdf).GetMethod("CreateReviewReport", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                    var document = (MigraDoc.DocumentObjectModel.Document)create.Invoke(new ClearPlan.Reporting.MigraDoc.ReportPdf(), new object[] { report });
                    string ddl = MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToString(document);
                    TestAssert.Equal("Plan Quality Report", document.Info.Title);
                    TestAssert.True(ddl.Contains(code == "en" ? "Delivery and modulation summary" : "Applikations- und Modulationsübersicht"));
                    TestAssert.True(ddl.Contains(code == "en" ? "Field identifiers and names" : "Feldkennungen und -namen"));
                    TestAssert.True(ddl.Contains(code == "en" ? "Message" : "Meldung"));
                    TestAssert.True(ddl.Contains(code == "en" ? "Not met (Error)" : "Nicht erfüllt (Fehler)"));
                    TestAssert.True(ddl.Contains("DEMO_Feld_Strukturen_ä_01"));
                    AssertReportLanguageIdentities(ddl);
                    TestAssert.Equal(guiCode, ClearPlan.Core.Localization.ReviewLanguage.Code, "PDF scope must restore the GUI language.");
                    TestAssert.Equal("pass", report.PlanCheckRows[0].Status);
                    TestAssert.Equal("fail", report.PlanCheckRows[1].Status);
                    TestAssert.Equal("error", report.PlanCheckRows[1].Severity);
                    TestAssert.Equal(95d, report.PqmRows[0].Goal.Value);
                    TestAssert.Equal("Strukturen", report.PqmRows[0].ResolvedStructureId);
                }
            }
            finally { ClearPlan.Core.Localization.ReviewLanguage.Code = original; }
        }

        private static void AssertReportLanguageIdentities(string rendered)
        {
            foreach (string exact in new[] { "DEMO_Patient_Strukturen_ä_01", "DEMO_Plan_Strukturen_ä_01", "Strukturen", "Lokale Notiz für Strukturen: ORIGINAL_XYZ" })
                TestAssert.True(rendered.Contains(exact), "Report language must preserve identity/free diagnostic text: " + exact);
            TestAssert.False(rendered.Contains("DEMO_Patient_Structures") || rendered.Contains("DEMO_Plan_Structures"));
        }

        private static ClearPlan.Reporting.ReviewReportDocument LanguageReportFixture()
        {
            var report = new ClearPlan.Reporting.ReviewReportDocument {
                Synthetic = false, LanguageCode = "en", ActivePlanKey = "DEMO_KEY", IncludeBeamEyeViews = false,
                PatientDisplayLabel = "DEMO_Patient_Strukturen_ä_01", PlanDisplayLabel = "DEMO_Plan_Strukturen_ä_01",
                ModeLabel = "TEST FIXTURE - NOT CLINICAL DATA", Watermark = "TEST FIXTURE - NOT CLINICAL DATA",
                GeneratedUtc = new DateTimeOffset(2026, 9, 16, 12, 0, 0, TimeSpan.Zero)
            };
            report.Plans.Add(new ClearPlan.Reporting.ReviewReportPlanRow { PlanKey = "DEMO_KEY", DisplayLabel = report.PlanDisplayLabel, TargetDisplayLabel = "Strukturen", FractionCount = 5, TotalDoseGy = 25, DosePerFractionGy = 5 });
            report.PqmRows.Add(new ClearPlan.Reporting.ReviewReportPqmRow { TemplateStructure = "Strukturen", ResolvedStructureId = "Strukturen", Objective = "D98%", Comparator = ">", Goal = 95, AchievedValue = 96, Unit = "%", Status = "pass", Severity = "none" });
            report.PlanCheckRows.Add(new ClearPlan.Reporting.ReviewReportCheckRow { CheckCode = "DEMO_CHECK_1", Category = "Default", Status = "pass", Severity = "none", Message = "Lokale Notiz für Strukturen: ORIGINAL_XYZ" });
            report.PlanCheckRows.Add(new ClearPlan.Reporting.ReviewReportCheckRow { CheckCode = "DEMO_CHECK_2", Category = "Default", Status = "fail", Severity = "error", Message = "DEMO deliberate failure" });
            report.FieldRows.Add(new ClearPlan.Reporting.ReviewReportFieldRow { CurrentId = "DEMO_Feld_Strukturen_ä_01", CurrentName = "DEMO_Name_Strukturen_ä_01", IdStatus = "pass", NameStatus = "fail" });
            return report;
        }

        public static void AllViewsAndPlotIdentity()
        {
            string original=ClearPlan.Core.Localization.ReviewLanguage.Code;
            try {
                foreach(var code in new[]{"de","en"}) {
                    ClearPlan.Core.Localization.ReviewLanguage.Code=code;
                    foreach(var view in new System.Windows.FrameworkElement[]{
                        new ClearPlan.Presentation.Views.ReviewWorkspaceView(),
                        new ClearPlan.Presentation.Views.PlanImagesView(),
                        new ClearPlan.Presentation.Views.PlanParametersView(),
                        new ClearPlan.Presentation.Views.PlanComparisonView(),
                        new ClearPlan.Presentation.Views.BeamEyeView(),
                        new ClearPlan.Presentation.Views.CollisionView() }) {
                        view.Measure(new System.Windows.Size(1300,850));
                        view.Arrange(new System.Windows.Rect(0,0,1300,850));
                        view.UpdateLayout();
                    }
                }
                var row=new ClearPlan.Presentation.ViewModels.ReviewDvhSeriesViewModel(new ReviewDvhSeries {
                    StableId="patient-structure-key", StructureId="PTV_Strukturen", DisplayName="PTV_Strukturen", Selected=true });
                var plot=ClearPlan.Presentation.Plot.ReviewPlotFactory.Create(new[]{row});
                var line=(OxyPlot.Series.LineSeries)plot.Series[0];
                var points=line.Points;
                line.IsVisible=false;
                ClearPlan.Core.Localization.ReviewLanguage.Code="de";
                ClearPlan.Presentation.Plot.ReviewPlotFactory.RefreshLanguage(plot);
                TestAssert.Equal("Dosis [Gy]",plot.Axes[0].Title);
                ClearPlan.Core.Localization.ReviewLanguage.Code="en";
                ClearPlan.Presentation.Plot.ReviewPlotFactory.RefreshLanguage(plot);
                TestAssert.Equal("Dose [Gy]",plot.Axes[0].Title);
                TestAssert.Equal("PTV_Strukturen",line.Title);
                TestAssert.True(object.ReferenceEquals(points,line.Points));
                TestAssert.True(!line.IsVisible);
                TestAssert.Equal("PTV_Strukturen",ClearPlan.Core.Localization.ReviewLanguage.Display("PTV_Strukturen"));
            } finally { ClearPlan.Core.Localization.ReviewLanguage.Code=original; }
        }
        public static void LiveBindingsAndReportScope()
        {
            var language=ClearPlan.Presentation.Localization.DisplayLanguage.Instance;
            string original=ClearPlan.Core.Localization.ReviewLanguage.Code;
            try {
                ClearPlan.Core.Localization.ReviewLanguage.Code="de";
                var block=(System.Windows.Controls.TextBlock)System.Windows.Markup.XamlReader.Parse(
                  "<TextBlock xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' xmlns:l='clr-namespace:ClearPlan.Presentation.Localization;assembly=ClearPlan.Presentation' Text='{l:Translate Value=Schnittbilder}'/>");
                block.UpdateLayout();
                TestAssert.Equal("Schnittbilder",block.Text);
                ClearPlan.Core.Localization.ReviewLanguage.Code="en";
                System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(System.Windows.Threading.DispatcherPriority.DataBind,new Action(()=>{}));
                TestAssert.Equal("CT views",block.Text);
                var report=new ClearPlan.Reporting.ReviewReportDocument();
                TestAssert.Equal("en",report.LanguageCode);
                ClearPlan.Core.Localization.ReviewLanguage.Code="de";
                using(ClearPlan.Core.Localization.ReviewLanguage.Scope(report.LanguageCode))
                    TestAssert.Equal("CT views",ClearPlan.Core.Localization.ReviewLanguage.Text("Schnittbilder"));
                TestAssert.Equal("de",ClearPlan.Core.Localization.ReviewLanguage.Code);
                System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(System.Windows.Threading.DispatcherPriority.DataBind,new Action(()=>{}));
                TestAssert.Equal("Schnittbilder",block.Text);
                TestAssert.Throws<ArgumentException>(()=>ClearPlan.Core.Localization.ReviewLanguage.Code="xx");
            } finally { ClearPlan.Core.Localization.ReviewLanguage.Code=original; }
        }
        public static void EnglishGermanRoundTrip()
        {
            var type = typeof(ReviewSnapshot).Assembly.GetType("ClearPlan.Core.Localization.ReviewLanguage");
            TestAssert.NotNull(type, "Review display language service is missing.");
            var language = type.GetProperty("Code");
            var text = type.GetMethod("Text", new[] {typeof(string)});
            string original = (string)language.GetValue(null);
            try {
                language.SetValue(null, "en");
                TestAssert.Equal("CT views", (string)text.Invoke(null,new object[] {"Schnittbilder"}));
                TestAssert.Equal("PTV_L_ä_01", (string)text.Invoke(null,new object[] {"PTV_L_ä_01"}));
                language.SetValue(null,"de");
                TestAssert.Equal("Schnittbilder", (string)text.Invoke(null,new object[] {"Schnittbilder"}));
            } finally { language.SetValue(null,original); }
        }
    }
}
