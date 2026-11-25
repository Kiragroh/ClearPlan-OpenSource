using System;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Security.Principal;
using ClearPlan;
using ClearPlan.Helpers;
using ClearPlan.Core.Configuration;
using ClearPlan.Core.Constraints;
using EsapiEssentials.Plugin;
using Newtonsoft.Json;
using VMS.TPS.Common.Model.API;

namespace ClearPlan.PhantomValidation
{
    public sealed class Target
    {
        public string PatientId, CourseId, PlanId, TableId;
        public int Fractions;
        public double TotalDoseGy;
    }
    public sealed class Manifest { public Target[] Targets; }
    public static class Selection
    {
        public static void Apply(MainViewModel vm, Patient patient, PlanSetup plan)
        {
            var root = ClearPlanSettings.Load().BaseDirectory;
            var manifest = JsonConvert.DeserializeObject<Manifest>(File.ReadAllText(Path.Combine(root, "phantom-validation.json")));
            var target = manifest.Targets.SingleOrDefault(t => t.PatientId == patient.Id && t.CourseId == plan.Course.Id && t.PlanId == plan.Id);
            if (target == null) throw new InvalidOperationException("Not an authorized phantom context.");
            double dose = plan.TotalDose.Unit == VMS.TPS.Common.Model.Types.DoseValue.DoseUnit.cGy ? plan.TotalDose.Dose / 100 : plan.TotalDose.Dose;
            if (plan.NumberOfFractions != target.Fractions || Math.Abs(dose - target.TotalDoseGy) > 0.001 || !plan.IsDoseValid)
                throw new InvalidOperationException("Phantom prescription or dose differs from test protocol.");
            var selected = vm.ConstraintComboBoxList.Single(t => t.ConstraintId == target.TableId);
            if (!selected.ConstraintName.StartsWith("Phantom validation", StringComparison.Ordinal))
                throw new InvalidOperationException("Only the explicitly authorized demonstration table may be confirmed.");
            vm.ActiveConstraintPath = selected;
            vm.ConfirmConstraintSelection(selected);
        }
    }
    // Fixed-purpose launcher. No arbitrary executable or command accepted from input.
    public sealed class MiningLauncher
    {
        public void Run(PluginScriptContext context)
        {
            if (context == null || context.Patient == null || context.PlanSetup == null)
                throw new InvalidOperationException("Native phantom plan context required.");
            var root = Path.GetDirectoryName(typeof(MiningLauncher).Assembly.Location);
            var manifest = JsonConvert.DeserializeObject<Manifest>(File.ReadAllText(Path.Combine(root,"phantom-validation.json")));
            if (!manifest.Targets.Any(t => t.PatientId == context.Patient.Id && t.CourseId == context.PlanSetup.Course.Id && t.PlanId == context.PlanSetup.Id))
                throw new InvalidOperationException("Unauthorized context.");
            var psi = new ProcessStartInfo(Path.Combine(root,"ClearPlan.RetrospectiveReviewMining.exe"),
                "--input \"" + Path.Combine(root,"phantom-plans.csv") + "\"") {
                WorkingDirectory=root, UseShellExecute=false, CreateNoWindow=true,
                RedirectStandardOutput=true, RedirectStandardError=true
            };
            using(var child = Process.Start(psi))
            {
                var stdout = child.StandardOutput.ReadToEndAsync();
                var stderr = child.StandardError.ReadToEndAsync();
                // Do not terminate unknown/native state on timeout; retain its process identity.
                if (!child.WaitForExit(480000)) throw new TimeoutException("Mining child still running; inspect before retry.");
                string prefix=Path.Combine(ClearPlanSettings.Load().GetReportsDirectory(),"mining-"+DateTime.UtcNow.ToString("yyyyMMdd-HHmmss"));
                File.WriteAllText(prefix+".stdout.txt",stdout.GetAwaiter().GetResult());
                File.WriteAllText(prefix+".stderr.txt",stderr.GetAwaiter().GetResult());
                File.WriteAllText(prefix+".result.json",JsonConvert.SerializeObject(new { ExitCode=child.ExitCode, NativeHost=Environment.MachineName, FinishedUtc=DateTime.UtcNow },Formatting.Indented));
                if(child.ExitCode!=0) throw new InvalidOperationException("Native retrospective child did not complete successfully.");
            }
        }
    }
    public static class ConfigurationSetup
    {
        public static int Main(string[] args)
        {
            var store=new ConfigurationHistoryStore(args[0]);
            string actor=WindowsIdentity.GetCurrent().Name;
            var first=store.Import("phantom-goals","default",".json",File.ReadAllBytes(args[1]),actor,"Physical phantom validation baseline only");
            var second=store.Save("phantom-goals",File.ReadAllBytes(args[2]),actor,"Change only two phantom complementary-volume thresholds",first.RevisionNumber);
            Console.WriteLine(first.RevisionNumber+","+second.RevisionNumber);
            return 0;
        }
    }
}
