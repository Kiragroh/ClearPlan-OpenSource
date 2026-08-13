using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;
using ClearPlan.Core.Settings;

namespace ClearPlan.Runner
{
    internal static class RunnerFailureReporter
    {
        public static bool IsPathFailure(Exception exception)
        {
            return PathFailureClassifier.IsPathFailure(exception);
        }

        public static void ReportAndShow(
            string context,
            Exception exception)
        {
            string logPath = Report(context, exception);
            MessageBox.Show(
                context + Environment.NewLine + Environment.NewLine +
                PathFailureClassifier.Unwrap(exception).Message + Environment.NewLine +
                Environment.NewLine +
                "Diagnoseprotokoll: " + logPath,
                "ClearPlan Runner",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        public static string Report(
            string context,
            Exception exception)
        {
            string logPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Logs",
                "RunnerErrors.log");
            string entry = string.Format(
                "{0:o}{1}{2}{1}{3}{1}{1}",
                DateTime.Now,
                Environment.NewLine,
                context,
                PathFailureClassifier.Unwrap(exception));
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
                File.AppendAllText(
                    logPath,
                    entry,
                    new UTF8Encoding(false));
            }
            catch
            {
                logPath = Path.Combine(
                    Path.GetTempPath(),
                    "ClearPlan.RunnerErrors.log");
                try
                {
                    File.AppendAllText(
                        logPath,
                        entry,
                        new UTF8Encoding(false));
                }
                catch
                {
                }
            }

            return logPath;
        }
    }
}
