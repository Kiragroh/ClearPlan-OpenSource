using System;
using System.IO;
using System.Linq;
using System.Windows;

namespace ClearPlan.Simulator
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs eventArgs)
        {
            bool automatedModeRequested = (eventArgs.Args ?? new string[0]).Any(
                value =>
                    value.Equals(
                        "--capture",
                        StringComparison.OrdinalIgnoreCase) ||
                    value.Equals(
                        "--capture-all",
                        StringComparison.OrdinalIgnoreCase) ||
                    value.Equals(
                        "--layout-probe",
                        StringComparison.OrdinalIgnoreCase));

            // WPF otherwise silently returns transparent RenderTargetBitmaps when
            // this short-lived capture process has no attached display device.
            // This switch is process-local and capture-only: no registry, global
            // rendering policy or interactive/clinical application is changed.
            if (automatedModeRequested)
                AppContext.SetSwitch(
                    "Switch.System.Windows.Media.ShouldRenderEvenWhenNoDisplayDevicesAreAvailable",
                    true);
            base.OnStartup(eventArgs);

            try
            {
                SimulatorArguments arguments =
                    SimulatorArguments.Parse(eventArgs.Args);
                string executableDirectory =
                    AppDomain.CurrentDomain.BaseDirectory;
                var repository = new SimulatorScenarioRepository(
                    executableDirectory);
                var window = new MainWindow(
                    repository,
                    arguments,
                    new SyntheticReportService(),
                    new VisualCaptureService(),
                    new SyntheticDvhPngService(),
                    new ResponsiveLayoutProbeService());
                MainWindow = window;
                window.Show();
            }
            catch (Exception exception)
            {
                HandleFatal(exception, automatedModeRequested);
            }
        }

        internal static void HandleFatal(
            Exception exception,
            bool captureMode)
        {
            Exception failure = exception ??
                new InvalidOperationException(
                    "Unknown ClearPlan simulator failure.");
            try
            {
                Console.Error.WriteLine(failure);
            }
            catch
            {
            }

            if (captureMode)
            {
                try
                {
                    string diagnosticPath = Path.Combine(
                        Path.GetTempPath(),
                        "ClearPlan.Simulator.last-error.txt");
                    File.WriteAllText(diagnosticPath, failure.ToString());
                }
                catch
                {
                }
            }
            else
            {
                MessageBox.Show(
                    failure.Message,
                    "ClearPlan Simulator",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }

            Environment.ExitCode = 3;
            if (Current != null)
            {
                Current.Shutdown(3);
            }
        }
    }
}
