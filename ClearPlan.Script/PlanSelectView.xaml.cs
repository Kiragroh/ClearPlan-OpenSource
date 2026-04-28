using ClearPlan.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using VMS.TPS.Common.Model.API;

namespace ClearPlan
{
    public partial class PlanSelectView : UserControl
    {
        private readonly PlanSelectViewModel _psvm;

        public PlanSelectView(PlanSelectViewModel planSelectViewModel)
        {
            _psvm = planSelectViewModel;
            InitializeComponent();
            DataContext = _psvm;
        }

        private void ExitButton_OnClick(object sender, RoutedEventArgs e)
        {
            var window = Window.GetWindow(this);
            if (window != null)
            {
                window.Topmost = false;
                window.Close();
            }
        }

        private void RunButton_OnClick(object sender, RoutedEventArgs e)
        {
            var selectedItem = (PlanSelectDetailViewModel)planningItemSummariesDataGrid.SelectedItem;
            if (selectedItem == null)
            {
                MessageBox.Show("Please select a plan before pressing 'Run'.", "OOpsi...", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _psvm.ActivePlanningItem = selectedItem.ActivePlanningItem;
            var mainViewModel = new MainViewModel(
                _psvm.User,
                _psvm.Patient,
                _psvm.ScriptVersion,
                _psvm.PlanningItemList,
                _psvm.ActivePlanningItem,
                selectedItem.ActivePlanningItem.PlanningItemObject);

            _psvm.MainWindow.Title = mainViewModel.Title;
            _psvm.MainWindow.Content = new MainView(mainViewModel);
            _psvm.MainWindow.Height = SystemParameters.PrimaryScreenHeight * 0.85;
            _psvm.MainWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;

            LogPlanSelection(selectedItem);

            var performSort = typeof(DataGrid).GetMethod("PerformSort", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            performSort?.Invoke(planningItemSummariesDataGrid, new[] { planningItemSummariesDataGrid.Columns[7] });
            performSort?.Invoke(planningItemSummariesDataGrid, new[] { planningItemSummariesDataGrid.Columns[7] });
        }

        private void LogPlanSelection(PlanSelectDetailViewModel selectedItem)
        {
            string userLogPath = ClearPlanSettings.Load().GetUsageLogFile();
            var userLogCsvContent = new StringBuilder();

            if (!File.Exists(userLogPath))
            {
                var dataHeaderList = new List<string>
                {
                    "User",
                    "PC",
                    "Script",
                    "Version",
                    "Date",
                    "DayOfWeek",
                    "Time",
                    "PatientID",
                    "StructureSetId",
                    "PlanID",
                    "CourseID"
                };

                userLogCsvContent.AppendLine(string.Join(",", dataHeaderList.ToArray()));
            }

            string planId;
            string courseId;
            string structureSetId;

            try
            {
                if (selectedItem.ActivePlanningItem.PlanningItemObject is PlanSetup plan)
                {
                    planId = plan.Id.Replace(",", "");
                    courseId = plan.Course.Id.Replace(",", "");
                    structureSetId = plan.StructureSet.Id.Replace(",", "");
                }
                else
                {
                    var planSum = (PlanSum)selectedItem.ActivePlanningItem.PlanningItemObject;
                    planId = planSum.Id.Replace(",", "");
                    courseId = planSum.Course.Id.Replace(",", "");
                    structureSetId = planSum.StructureSet.Id.Replace(",", "");
                }
            }
            catch
            {
                planId = "error";
                courseId = "error";
                structureSetId = "error";
            }

            var culture = new System.Globalization.CultureInfo("de-DE");
            string dayOfWeek = culture.DateTimeFormat.GetDayName(DateTime.Today.DayOfWeek);
            string version = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).FileVersion;
            string userId = _psvm.User.Id.Replace(@"oncology\", "").Replace(",", "") +
                            "-" +
                            _psvm.User.Name.Replace(@"oncology\", "").Replace(",", "");

            var userStatsList = new List<object>
            {
                userId,
                Environment.MachineName,
                "ClearPlan",
                version,
                DateTime.Now.ToString("yyyy-MM-dd"),
                dayOfWeek,
                DateTime.Now.ToString("HH:mm"),
                _psvm.Patient.Id.Replace(",", ""),
                structureSetId,
                planId,
                courseId
            };

            userLogCsvContent.AppendLine(string.Join(",", userStatsList.ToArray()));
            File.AppendAllText(userLogPath, userLogCsvContent.ToString(), Encoding.UTF8);
        }
    }
}
