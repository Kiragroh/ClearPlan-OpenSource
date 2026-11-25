using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using VMS.TPS.Common.Model.API;

namespace ClearPlan.Headless
{
    public static class NativeMlcDiagnostics
    {
        // Aggregate machine/geometry numbers only. The owner reads every CP; no identifiers,
        // patient geometry, vendor objects or native arrays escape this method.
        public static object Capture(PlanSetup plan)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Native diagnostics require the ESAPI owner STA.");
            var beams = new List<NativeMlcBeamDiagnostic>();
            var timer = Stopwatch.StartNew();
            string failure = null;
            try
            {
                if (plan == null) throw new ArgumentNullException("plan");
                var nativeBeams = plan.Beams.Take(129).ToList();
                if (nativeBeams.Count > 128) failure = "beam_limit";
                else
                {
                    foreach (var beam in nativeBeams)
                    {
                        if (timer.Elapsed.TotalSeconds > 10) { failure = "time_limit"; break; }
                        if (beam.IsSetupField || beam.MLC == null) continue;
                        var diagnostic = new NativeMlcBeamDiagnostic { BeamOrdinal = beams.Count + 1 };
                        beams.Add(diagnostic);
                        try
                        {
                            string model = beam.MLC.Model;
                            diagnostic.Model = string.IsNullOrWhiteSpace(model) ? null
                                : new string(model.Take(128).Select(c => char.IsControl(c) ? '?' : c).ToArray());
                            var controlPoints = beam.ControlPoints.Take(20001).ToList();
                            if (controlPoints.Count == 0 || controlPoints.Count > 20000)
                            { diagnostic.FailureCode = "control_point_limit"; continue; }
                            foreach (var point in controlPoints)
                            {
                                if (timer.Elapsed.TotalSeconds > 10) { diagnostic.FailureCode = "time_limit"; break; }
                                var jaws = point.JawPositions;
                                diagnostic.AddSample(point.LeafPositions, jaws.X1, jaws.Y1, jaws.X2, jaws.Y2);
                            }
                            diagnostic.ReadComplete = diagnostic.SampledControlPoints == controlPoints.Count;
                        }
                        catch (Exception) { diagnostic.FailureCode = "native_read"; }
                    }
                }
            }
            catch (Exception) { failure = "native_read"; }
            return new {
                Source = "Read-only ESAPI aggregate diagnostics; zero-based SX2 0-27 distal, 28-56 proximal per cited reference profile. Not commissioning.",
                JawRangeOrder = "X1,Y1,X2,Y2; native coordinates only, not an assertion of physical jaws",
                ReadComplete = failure == null && beams.All(b => b.ReadComplete),
                FailureCode = failure,
                Beams = beams
            };
        }
    }
    public sealed class NativeMlcRange
    {
        public double? MinimumMm { get; internal set; }
        public double? MaximumMm { get; internal set; }
        internal void Add(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value)) return;
            MinimumMm = MinimumMm.HasValue ? Math.Min(MinimumMm.Value, value) : value;
            MaximumMm = MaximumMm.HasValue ? Math.Max(MaximumMm.Value, value) : value;
        }
    }
    public sealed class NativeMlcLayerDiagnostic
    {
        public string Label { get; internal set; }
        public int LeafPairCount { get; internal set; }
        public NativeMlcRange Bank0 { get; internal set; }
        public NativeMlcRange Bank1 { get; internal set; }
    }
    public sealed class NativeMlcBeamDiagnostic
    {
        public NativeMlcBeamDiagnostic()
        {
            ShapeCounts = new SortedDictionary<string, int>(StringComparer.Ordinal);
            Layers = new List<NativeMlcLayerDiagnostic>();
            JawRanges = Enumerable.Range(0, 4).Select(i => new NativeMlcRange()).ToArray();
        }
        public int BeamOrdinal { get; internal set; }
        public string Model { get; set; }
        public bool ReadComplete { get; internal set; }
        public string FailureCode { get; internal set; }
        public int SampledControlPoints { get; internal set; }
        public SortedDictionary<string, int> ShapeCounts { get; private set; }
        public int InvalidShapeSampleCount { get; internal set; }
        public int InvalidLeafPairCount { get; internal set; }
        public int InvalidJawSampleCount { get; internal set; }
        public NativeMlcRange[] JawRanges { get; private set; }
        public List<NativeMlcLayerDiagnostic> Layers { get; private set; }
        public bool JawPositionsInvariant
        {
            get { return SampledControlPoints > 0 && InvalidJawSampleCount == 0 &&
                JawRanges.All(r => r.MinimumMm.HasValue && r.MaximumMm.HasValue &&
                    Math.Abs(r.MaximumMm.Value - r.MinimumMm.Value) <= 1e-6); }
        }
        internal void AddSample(float[,] positions, double x1, double y1, double x2, double y2)
        {
            SampledControlPoints++;
            var jaws = new[] { x1, y1, x2, y2 };
            if (jaws.Any(v => !Finite(v)) || x1 > x2 || y1 > y2) InvalidJawSampleCount++;
            for (int index = 0; index < 4; index++) JawRanges[index].Add(jaws[index]);
            int banks = positions == null ? 0 : positions.GetLength(0);
            int pairs = positions == null ? 0 : positions.GetLength(1);
            string shape = banks + "x" + pairs;
            if (!ShapeCounts.ContainsKey(shape)) ShapeCounts.Add(shape, 0);
            ShapeCounts[shape]++;
            if (banks != 2 || pairs < 1 || pairs > 2000 || Model == "SX2" && pairs != 57)
            { InvalidShapeSampleCount++; return; }
            if (Model == "SX2" && Layers.Count == 0)
            {
                Layers.Add(new NativeMlcLayerDiagnostic { Label = "Distal", LeafPairCount = 28,
                    Bank0 = new NativeMlcRange(), Bank1 = new NativeMlcRange() });
                Layers.Add(new NativeMlcLayerDiagnostic { Label = "Proximal", LeafPairCount = 29,
                    Bank0 = new NativeMlcRange(), Bank1 = new NativeMlcRange() });
            }
            for (int index = 0; index < pairs; index++)
            {
                double bank0 = positions[0, index], bank1 = positions[1, index];
                if (!Finite(bank0) || !Finite(bank1) || bank0 > bank1) InvalidLeafPairCount++;
                if (Model == "SX2")
                {
                    var layer = Layers[index < 28 ? 0 : 1];
                    layer.Bank0.Add(bank0); layer.Bank1.Add(bank1);
                }
            }
        }
        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
    }
}
