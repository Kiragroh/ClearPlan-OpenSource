using System.Windows;
using EsapiEssentials.PluginRunner;
using VMS.TPS;
using VMS.TPS.Common.Model.API;

namespace ClearPlan.Runner
{
    public partial class App : System.Windows.Application
    {
        private void App_OnStartup(object sender, StartupEventArgs e)
        {
            DispatcherUnhandledException += App_OnDispatcherUnhandledException;
            System.AppDomain.CurrentDomain.UnhandledException +=
                CurrentDomain_UnhandledException;
            RunnerWindowStyler.Register();
            try
            {
                ScriptRunner.Run(new Script());
            }
            catch (System.Exception exception)
            {
                RunnerFailureReporter.ReportAndShow(
                    "ClearPlan Runner konnte nicht gestartet werden.",
                    exception);
                Shutdown(1);
            }
        }

        private void App_OnDispatcherUnhandledException(
            object sender,
            System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            bool recoverablePathFailure =
                RunnerFailureReporter.IsPathFailure(e.Exception);
            RunnerFailureReporter.ReportAndShow(
                recoverablePathFailure
                    ? "Ein optionaler Datei- oder Netzwerkpfad war nicht erreichbar. ClearPlan bleibt geöffnet."
                    : "ClearPlan musste den aktuellen Vorgang abbrechen.",
                e.Exception);
            e.Handled = true;
            if (!recoverablePathFailure)
            {
                Shutdown(1);
            }
        }

        private static void CurrentDomain_UnhandledException(
            object sender,
            System.UnhandledExceptionEventArgs e)
        {
            RunnerFailureReporter.Report(
                "Unbehandelter Runner-Fehler",
                e.ExceptionObject as System.Exception ??
                new System.InvalidOperationException(
                    "Unknown non-Exception failure object."));
        }

        // Fix UnauthorizedScriptingAPIAccessException
        public void DoNothing(PlanSetup plan) { }
    }
}
