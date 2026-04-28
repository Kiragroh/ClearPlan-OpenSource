using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using ClearPlan;
using ClearPlan.Helpers;
using EsapiEssentials.Plugin;
using NLog;
using VMS.TPS.Common.Model.API;
using MessageBox = System.Windows.MessageBox;

namespace VMS.TPS
{
    public class Script : ScriptBase
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public override void Run(PluginScriptContext context)
        {
            var settings = ClearPlanSettings.Load();
            var mainWindow = new Window();

            PlanSetup planSetup = context.PlanSetup;
            PlanSum planSum = context.PlanSum;
            User user = context.CurrentUser;
            Patient patient = context.Patient;

            var planSetupsInScope = patient.Courses
                .SelectMany(course => course.PlanSetups)
                .ToList();

            var planSumsInScope = patient.Courses
                .SelectMany(course => course.PlanSums)
                .ToList();

            string scriptVersion = GetScriptVersion();
            DateTime lastCompileTime = File.GetLastWriteTime(System.Reflection.Assembly.GetExecutingAssembly().Location);
            TimeSpan daysSinceLastCompile = DateTime.Now - lastCompileTime;

            ShowVersionNotesIfNeeded(context, scriptVersion, settings);

            if ((planSum == null && planSetup == null) ||
                (planSum != null && !planSum.IsDoseValid()) ||
                (planSetup != null && !planSetup.IsDoseValid()))
            {
                var planSelectViewModel = new PlanSelectViewModel(
                    user,
                    patient,
                    scriptVersion,
                    planSetup,
                    planSetupsInScope,
                    planSumsInScope,
                    mainWindow);

                var planSelectView = new PlanSelectView(planSelectViewModel);
                mainWindow.Title = "Select a plan and constraint template";
                mainWindow.Content = planSelectView;
                mainWindow.Height = SystemParameters.PrimaryScreenHeight * 0.85;
                mainWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
            else
            {
                PlanningItem selectedPlanningItem = planSum ?? (PlanningItem)planSetup;
                var planningItemList = PlanningItemListViewModel.GetPlanningItemList(planSetupsInScope, planSumsInScope);
                var mainViewModel = new MainViewModel(
                    user,
                    patient,
                    scriptVersion,
                    planningItemList,
                    new PlanningItemViewModel(selectedPlanningItem),
                    selectedPlanningItem);

                mainWindow.Title = mainViewModel.Title + string.Format(" last updated {0} days before", (int)daysSinceLastCompile.TotalDays);
                mainWindow.Content = new MainView(mainViewModel);
                mainWindow.Height = SystemParameters.PrimaryScreenHeight * 0.85;
                mainWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;

                LogUserActivity(selectedPlanningItem, scriptVersion, user, patient, settings);
            }

            mainWindow.ShowInTaskbar = true;
            mainWindow.ShowActivated = true;

            try
            {
                mainWindow.Topmost = true;
            }
            catch
            {
            }

            mainWindow.ShowDialog();
        }

        private static string GetScriptVersion()
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            return FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion;
        }

        private static void ShowVersionNotesIfNeeded(PluginScriptContext context, string scriptVersion, ClearPlanSettings settings)
        {
            string seenVersionsFile = settings.GetVersionSeenUsersFile();
            string userKey = string.IsNullOrWhiteSpace(context.CurrentUser.Id)
                ? context.CurrentUser.Name
                : context.CurrentUser.Id;
            string marker = string.Format("{0};{1}", userKey, scriptVersion);

            var seenVersions = File.Exists(seenVersionsFile)
                ? new HashSet<string>(File.ReadAllLines(seenVersionsFile), StringComparer.OrdinalIgnoreCase)
                : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (seenVersions.Contains(marker))
            {
                return;
            }

            File.AppendAllLines(seenVersionsFile, new[] { marker });
            MessageBox.Show(
                BuildVersionMessage(context.CurrentUser.Name, scriptVersion, settings),
                "What's new in ClearPlan " + scriptVersion);
        }

        private static string BuildVersionMessage(string userName, string scriptVersion, ClearPlanSettings settings)
        {
            string notes = DocumentationHelper.GetLatestSection(settings.GetChangeLogFile(), scriptVersion);
            if (string.IsNullOrWhiteSpace(notes))
            {
                notes = "- This release focuses on easier onboarding, local defaults, and simpler starter examples.";
            }

            return string.Format(
                "Hello {0},\n\nYou are opening ClearPlan {1} for the first time on this workstation.\n\n{2}\n\nOpen Help > Change Log for the full notes.",
                userName,
                scriptVersion,
                notes);
        }

        private void LogUserActivity(PlanningItem selectedPlanningItem, string scriptVersion, User user, Patient patient, ClearPlanSettings settings)
        {
            var activityLogger = new CustomLog("ActivityLog.csv");

            try
            {
                string planId = "N/A";
                string courseId = "N/A";
                string structureSetId = "N/A";

                if (selectedPlanningItem is PlanSetup plan)
                {
                    planId = plan.Id;
                    courseId = plan.Course.Id;
                    structureSetId = plan.StructureSet.Id;
                }
                else if (selectedPlanningItem is PlanSum planSum)
                {
                    planId = planSum.Id;
                    courseId = planSum.Course.Id;
                    structureSetId = planSum.StructureSet.Id;
                }

                string userId = user.Id.Replace(@"oncology\", "").Replace(",", "") +
                                "-" +
                                user.Name.Replace(@"oncology\", "").Replace(",", "");
                string date = DateTime.Now.ToString("yyyy-MM-dd");
                string time = DateTime.Now.ToString("HH:mm");
                string dayOfWeek = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetDayName(DateTime.Now.DayOfWeek);

                string logEntry = string.Format(
                    "{0};{1};{2};{3};{4};ClearPlan;{5};{6};{7};{8};{9}",
                    date,
                    time,
                    dayOfWeek,
                    userId,
                    Environment.MachineName,
                    scriptVersion,
                    patient.Id.Replace(",", ""),
                    structureSetId,
                    planId,
                    courseId);

                string logFilePath = settings.GetActivityLogFile();
                bool fileExists = File.Exists(logFilePath);

                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    if (!fileExists)
                    {
                        writer.WriteLine("Date;Time;DayOfWeek;User;PC;Script;Version;PatientID;StructureSetId;PlanID;CourseID");
                    }

                    writer.WriteLine(logEntry);
                }
            }
            catch (Exception ex)
            {
                activityLogger.Error(ex, "Error logging user activity.");
                logger.Error(ex, "Error logging user activity.");
            }
        }
    }
}
