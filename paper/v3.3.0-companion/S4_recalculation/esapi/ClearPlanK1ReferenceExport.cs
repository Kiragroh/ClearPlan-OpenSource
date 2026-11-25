// K1 reference export for the ClearPlan phantom manuscript.
//
// Deliberately independent of every ClearPlan assembly: it only copies primitive
// ESAPI data (dose grid, structure contours, control points) plus the values that
// ESAPI itself reports, so that DVH metrics, CI/GI/HI, aperture areas and PAM can be
// recomputed outside ClearPlan by k1_reference.py. Read-only: no BeginModifications,
// no structure creation, no save. Runs only for contexts listed in the manifest
// beside this assembly, and writes only to the manifest's protected output folder.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly: ESAPIScript(IsWriteable = false)]

namespace VMS.TPS
{
    public class Script
    {
        // The host logs no exception messages, so a failure leaves a small report without identifiers.
        private string stage = "start";
        private readonly List<string> facts = new List<string>();

        public void Execute(ScriptContext context)
        {
            string root = Path.GetDirectoryName(typeof(Script).Assembly.Location);
            string failureDir = root;
            try
            {
                stage = "manifest";
                var manifest = Manifest.Load(Path.Combine(root, "k1-reference-targets.json"));
                failureDir = manifest.OutputDirectory;
                Run(context, manifest);
            }
            catch (Exception ex)
            {
                var text = new StringBuilder();
                text.AppendLine("stage=" + stage);
                text.AppendLine("exception=" + ex.GetType().FullName);
                // Only this script's own messages are copied; vendor messages can contain identifiers.
                if (ex is K1Exception) text.AppendLine("message=" + ex.Message);
                foreach (var f in facts) text.AppendLine(f);
                try
                {
                    Directory.CreateDirectory(failureDir);
                    File.WriteAllText(Path.Combine(failureDir, "k1-failure-" +
                        DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".txt"), text.ToString());
                }
                catch (Exception) { }
                throw;
            }
        }

        private void Run(ScriptContext context, Manifest manifest)
        {
            stage = "context";
            if (context == null || context.Patient == null || context.PlanSetup == null)
                throw new K1Exception("An explicit patient and plan context is required.");
            var plan = context.PlanSetup;
            var target = manifest.Find(context.Patient.Id, plan.Course.Id, plan.Id);
            if (target == null) throw new K1Exception("Context is not listed as an authorized phantom plan.");
            stage = "prescription";
            facts.Add("total_dose_unit=" + plan.TotalDose.Unit + " total_dose=" + plan.TotalDose.Dose.ToString(CultureInfo.InvariantCulture));
            facts.Add("fractions=" + plan.NumberOfFractions + " dose_valid=" + plan.IsDoseValid + " dose_present=" + (plan.Dose != null));
            double? rx = Gy(plan.TotalDose);
            if (!plan.IsDoseValid || plan.Dose == null || plan.NumberOfFractions != target.Fractions
                || !rx.HasValue || Math.Abs(rx.Value - target.TotalDoseGy) > 0.001)
                throw new K1Exception("Prescription or dose differs from the K1 protocol.");

            stage = "output";
            Directory.CreateDirectory(manifest.OutputDirectory);
            string stem = Path.Combine(manifest.OutputDirectory,
                "k1-" + target.Label + "-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture));
            var w = new Json();
            w.Open();
            w.Str("schema", "clearplan.k1_reference_export.v1");
            w.Str("label", target.Label);
            w.Str("context_sha256", Sha256(context.Patient.Id + "|" + plan.Course.Id + "|" + plan.Id));
            w.Str("exported_utc", DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
            w.Str("host", Environment.MachineName);
            w.Num("rx_gy", rx.Value);
            w.Num("fractions", plan.NumberOfFractions ?? 0);
            w.Str("treatment_orientation", plan.TreatmentOrientation.ToString());
            stage = "dose";
            WriteDose(w, plan, stem + ".dose.bin");
            stage = "structures";
            WriteStructures(w, plan, rx.Value);
            stage = "beams";
            WriteBeams(w, plan);
            w.Close();
            stage = "write";
            File.WriteAllText(stem + ".json", w.ToString(), new UTF8Encoding(false));
        }

        private void WriteDose(Json w, PlanSetup plan, string binPath)
        {
            var dose = plan.Dose;
            // VoxelToDoseValue is linear in the raw value; three samples verify that.
            var v0 = dose.VoxelToDoseValue(0);
            var v1 = dose.VoxelToDoseValue(1000000);
            var v2 = dose.VoxelToDoseValue(2000000);
            facts.Add("voxel_unit=" + v1.Unit + " voxel_1e6=" + v1.Dose.ToString("R", CultureInfo.InvariantCulture));
            // With relative presentation Eclipse reports voxels in percent of the plan's total dose.
            // The presentation is not switched here; k1_reference.py checks the conversion against
            // the absolute DVH maximum.
            double rx = Gy(plan.TotalDose).Value;
            Func<DoseValue, double?> toGy = v => v.Unit == DoseValue.DoseUnit.Percent ? v.Dose / 100.0 * rx : Gy(v);
            if (!toGy(v0).HasValue || !toGy(v1).HasValue || !toGy(v2).HasValue)
                throw new K1Exception("Dose voxels are not in Gy, cGy or percent.");
            double d0 = toGy(v0).Value, d1 = toGy(v1).Value, d2 = toGy(v2).Value;
            double scale = (d1 - d0) / 1000000.0;
            if (Math.Abs(d2 - (d0 + scale * 2000000.0)) > 1e-6 * Math.Max(1, Math.Abs(d2)))
                throw new K1Exception("Dose voxel scaling is not linear.");
            var buffer = new int[dose.XSize, dose.YSize];
            using (var bin = new BinaryWriter(File.Create(binPath)))
            {
                for (int k = 0; k < dose.ZSize; k++)
                {
                    dose.GetVoxels(k, buffer);
                    for (int j = 0; j < dose.YSize; j++)
                        for (int i = 0; i < dose.XSize; i++) bin.Write(buffer[i, j]);
                }
            }
            w.Key("dose"); w.Open();
            w.Str("file", Path.GetFileName(binPath));
            w.Str("layout", "int32 little-endian, z slowest, then y, x fastest");
            w.Num("gy_per_raw", scale); w.Num("gy_offset", d0);
            w.Str("voxel_unit", v1.Unit.ToString());
            w.Ints("size", new[] { dose.XSize, dose.YSize, dose.ZSize });
            w.Nums("res_mm", new[] { dose.XRes, dose.YRes, dose.ZRes });
            w.Vec("origin", dose.Origin);
            w.Vec("x_dir", dose.XDirection); w.Vec("y_dir", dose.YDirection); w.Vec("z_dir", dose.ZDirection);
            w.Close();
        }

        private static void WriteStructures(Json w, PlanSetup plan, double rx)
        {
            var image = plan.StructureSet.Image;
            w.Key("image"); w.Open();
            w.Num("z_res_mm", image.ZRes); w.Num("z_size", image.ZSize);
            w.Vec("origin", image.Origin); w.Vec("z_dir", image.ZDirection);
            w.Close();
            w.Key("structures"); w.OpenArray();
            foreach (var s in plan.StructureSet.Structures.Where(s => !s.IsEmpty && s.HasSegment).OrderBy(s => s.Id).Take(80))
            {
                w.Open();
                w.Str("id", s.Id); w.Str("dicom_type", s.DicomType);
                w.Num("esapi_volume_cm3", s.Volume);
                WriteEsapiDvh(w, plan, s, rx);
                w.Key("contours"); w.OpenArray();
                for (int z = 0; z < image.ZSize; z++)
                    foreach (var contour in s.GetContoursOnImagePlane(z))
                    {
                        if (contour.Length < 3) continue;
                        w.Open();
                        w.Num("plane", z);
                        w.Key("xyz"); w.OpenArray();
                        foreach (var p in contour) w.RawNums(p.x, p.y, p.z);
                        w.CloseArray();
                        w.Close();
                    }
                w.CloseArray();
                w.Close();
            }
            w.CloseArray();
        }

        // ESAPI's own numbers: same vendor DVH ClearPlan reads, different code path.
        private static void WriteEsapiDvh(Json w, PlanSetup plan, Structure s, double rx)
        {
            w.Key("esapi"); w.Open();
            try
            {
                var dvh = plan.GetDVHCumulativeData(s, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.001);
                if (dvh != null)
                {
                    w.Num("dvh_volume_cm3", dvh.Volume);
                    w.Num("mean_gy", Gy(dvh.MeanDose)); w.Num("max_gy", Gy(dvh.MaxDose)); w.Num("min_gy", Gy(dvh.MinDose));
                    w.Num("coverage", dvh.Coverage); w.Num("sampling_coverage", dvh.SamplingCoverage);
                }
                w.Num("d98_gy", Gy(plan.GetDoseAtVolume(s, 98, VolumePresentation.Relative, DoseValuePresentation.Absolute)));
                w.Num("d2_gy", Gy(plan.GetDoseAtVolume(s, 2, VolumePresentation.Relative, DoseValuePresentation.Absolute)));
                w.Num("v_rx_cm3", plan.GetVolumeAtDose(s, new DoseValue(rx, DoseValue.DoseUnit.Gy), VolumePresentation.AbsoluteCm3));
                w.Num("v_half_rx_cm3", plan.GetVolumeAtDose(s, new DoseValue(rx * 0.5, DoseValue.DoseUnit.Gy), VolumePresentation.AbsoluteCm3));
                w.Num("v20gy_cm3", plan.GetVolumeAtDose(s, new DoseValue(20, DoseValue.DoseUnit.Gy), VolumePresentation.AbsoluteCm3));
            }
            catch (Exception ex) { w.Str("error", ex.GetType().Name); }
            w.Close();
        }

        private static void WriteBeams(Json w, PlanSetup plan)
        {
            var targets = plan.StructureSet.Structures.Where(s => !s.IsEmpty && s.HasSegment &&
                (string.Equals(s.DicomType, "PTV", StringComparison.OrdinalIgnoreCase) ||
                 s.Id.StartsWith("PTV", StringComparison.OrdinalIgnoreCase))).ToList();
            w.Key("beams"); w.OpenArray();
            // Same beam set and native order as ClearPlan's analysis; the comparison binds beams by position.
            foreach (var beam in plan.Beams.Where(b => !b.IsSetupField && !b.IsImagingTreatmentField))
            {
                w.Open();
                w.Str("id", beam.Id);
                w.Str("machine", beam.TreatmentUnit == null ? null : beam.TreatmentUnit.Id);
                w.Str("mlc_model", beam.MLC == null ? null : beam.MLC.Model);
                w.Num("meterset_mu", beam.Meterset.Unit == DosimeterUnit.MU ? (double?)beam.Meterset.Value : null);
                w.Num("weight_factor", beam.WeightFactor);
                w.Str("gantry_direction", beam.GantryDirection.ToString());
                w.Vec("isocenter", beam.IsocenterPosition);
                w.Key("control_points"); w.OpenArray();
                foreach (var cp in beam.ControlPoints)
                {
                    w.Open();
                    w.Num("meterset_weight", cp.MetersetWeight);
                    w.Num("gantry", cp.GantryAngle); w.Num("collimator", cp.CollimatorAngle); w.Num("couch", cp.PatientSupportAngle);
                    var j = cp.JawPositions;
                    w.Nums("jaws_x1_y1_x2_y2", new[] { j.X1, j.Y1, j.X2, j.Y2 });
                    w.Vec("source", beam.GetSourceLocation(cp.GantryAngle));
                    var leaves = cp.LeafPositions;
                    w.Key("leaves"); w.OpenArray();
                    if (leaves != null)
                        for (int b = 0; b < leaves.GetLength(0); b++)
                        {
                            var bank = new double[leaves.GetLength(1)];
                            for (int i = 0; i < bank.Length; i++) bank[i] = leaves[b, i];
                            w.RawNums(bank);
                        }
                    w.CloseArray();
                    w.Close();
                }
                w.CloseArray();
                // Native BEV silhouette at the first control point, used only to check
                // the orientation convention of the independent projection.
                w.Key("native_bev_outlines"); w.OpenArray();
                foreach (var t in targets)
                {
                    w.Open(); w.Str("structure", t.Id);
                    w.Key("polygons"); w.OpenArray();
                    try
                    {
                        foreach (var poly in beam.GetStructureOutlines(t, true))
                            w.RawNums(poly.SelectMany(p => new[] { p.X, p.Y }).ToArray());
                    }
                    catch (Exception) { }
                    w.CloseArray(); w.Close();
                }
                w.CloseArray();
                w.Close();
            }
            w.CloseArray();
        }

        private static double? Gy(DoseValue v)
        {
            if (v.Unit == DoseValue.DoseUnit.Gy) return v.Dose;
            if (v.Unit == DoseValue.DoseUnit.cGy) return v.Dose / 100.0;
            return null;
        }

        private static string Sha256(string text)
        {
            using (var sha = SHA256.Create())
                return string.Concat(sha.ComputeHash(Encoding.UTF8.GetBytes(text)).Select(b => b.ToString("x2")));
        }
    }

    internal sealed class K1Exception : InvalidOperationException
    {
        public K1Exception(string message) : base(message) { }
    }

    internal sealed class ManifestTarget
    {
        public string PatientId, CourseId, PlanId, Label;
        public int Fractions;
        public double TotalDoseGy;
    }

    // Minimal parser for the fixed manifest shape; avoids shipping a JSON library.
    internal sealed class Manifest
    {
        public string OutputDirectory;
        public List<ManifestTarget> Targets = new List<ManifestTarget>();

        public static Manifest Load(string path)
        {
            var m = new Manifest();
            var root = (Dictionary<string, object>)new System.Web.Script.Serialization.JavaScriptSerializer()
                .DeserializeObject(File.ReadAllText(path));
            m.OutputDirectory = (string)root["OutputDirectory"];
            foreach (Dictionary<string, object> t in (object[])root["Targets"])
                m.Targets.Add(new ManifestTarget {
                    PatientId = (string)t["PatientId"], CourseId = (string)t["CourseId"], PlanId = (string)t["PlanId"],
                    Label = (string)t["Label"], Fractions = Convert.ToInt32(t["Fractions"], CultureInfo.InvariantCulture),
                    TotalDoseGy = Convert.ToDouble(t["TotalDoseGy"], CultureInfo.InvariantCulture) });
            if (string.IsNullOrWhiteSpace(m.OutputDirectory) || m.Targets.Any(t => !System.Text.RegularExpressions.Regex.IsMatch(t.Label ?? "", "^[A-Za-z0-9_-]{1,32}$")))
                throw new K1Exception("Manifest is incomplete.");
            return m;
        }

        public ManifestTarget Find(string patient, string course, string plan)
        {
            var hits = Targets.Where(t => t.PatientId == patient && t.CourseId == course && t.PlanId == plan).ToList();
            return hits.Count == 1 ? hits[0] : null;
        }
    }

    internal sealed class Json
    {
        private readonly StringBuilder sb = new StringBuilder();
        private readonly Stack<bool> first = new Stack<bool>();
        private bool afterKey;

        // Writes the comma before the next member or element; a value right after its key needs none.
        private void Sep()
        {
            if (afterKey) { afterKey = false; return; }
            if (first.Count == 0) return;
            if (!first.Peek()) sb.Append(',');
            first.Pop(); first.Push(false);
        }
        public void Open() { Sep(); sb.Append('{'); first.Push(true); }
        public void Close() { sb.Append('}'); first.Pop(); }
        public void OpenArray() { Sep(); sb.Append('['); first.Push(true); }
        public void CloseArray() { sb.Append(']'); first.Pop(); }
        public void Key(string k) { Sep(); Quote(k); sb.Append(':'); afterKey = true; }

        public void Str(string k, string v) { Key(k); Sep(); if (v == null) sb.Append("null"); else Quote(v); }
        public void Num(string k, double? v) { Key(k); Sep(); sb.Append(Fmt(v)); }
        public void Nums(string k, double[] v) { Key(k); Sep(); sb.Append('[').Append(string.Join(",", v.Select(x => Fmt(x)))).Append(']'); }
        public void Ints(string k, int[] v) { Key(k); Sep(); sb.Append('[').Append(string.Join(",", v)).Append(']'); }
        public void Vec(string k, VVector v) { Nums(k, new[] { v.x, v.y, v.z }); }
        public void RawNums(params double[] v) { Sep(); sb.Append('[').Append(string.Join(",", v.Select(x => Fmt(x)))).Append(']'); }

        private static string Fmt(double? v)
        {
            return v.HasValue && !double.IsNaN(v.Value) && !double.IsInfinity(v.Value)
                ? v.Value.ToString("R", CultureInfo.InvariantCulture) : "null";
        }
        private void Quote(string s)
        {
            sb.Append('"');
            foreach (char c in s)
            {
                if (c == '"' || c == '\\') sb.Append('\\').Append(c);
                else if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("x4"));
                else sb.Append(c);
            }
            sb.Append('"');
        }
        public override string ToString() { return sb.ToString(); }
    }
}
