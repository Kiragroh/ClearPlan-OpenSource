using System;

namespace ClearPlan.Core.PlanAnalysis
{
    public static class DoseRateEstimator
    {
        /// <summary>
        /// Estimates interval time from the slower of configured constant-speed gantry travel and
        /// nominal-rate MU delivery. This detached model is not an ESAPI time-resolved trajectory.
        /// All results are published atomically after validating every interval.
        /// </summary>
        public static void Apply(ReviewBeamAnalysis beam, DoseRateEstimationProfile profile)
        {
            if (beam == null) throw new ArgumentNullException("beam");
            Clear(beam);
            string profileError = DoseRateEstimationProfileCatalog.ValidateProfile(profile);
            if (profileError != null) { Unavailable(beam, profileError); return; }
            if (!DoseRateEstimationProfileCatalog.Matches(profile, beam))
            { Unavailable(beam, "The selected profile does not match the exact machine metadata."); return; }
            if (!Is(beam.Technique, "ARC") && !Is(beam.Technique, "SRS ARC") && !Is(beam.Technique, "VMAT") && !Is(beam.Technique, "RapidArc"))
            { Unavailable(beam, "Only explicitly identified ARC, SRS ARC, VMAT or RapidArc beams are supported."); return; }
            bool clockwise = Is(beam.GantryDirection, "Clockwise");
            if (!clockwise && !Is(beam.GantryDirection, "CounterClockwise"))
            { Unavailable(beam, "A known Clockwise or CounterClockwise gantry direction is required."); return; }
            if (!Positive(beam.MetersetMu) || !Positive(beam.NominalDoseRateMuPerMin))
            { Unavailable(beam, "Finite positive beam MU and nominal dose rate in MU/min are required."); return; }
            var points = beam.ControlPoints;
            if (points == null || points.Count < 2)
            { Unavailable(beam, "At least two complete control points are required."); return; }
            for (int index = 0; index < points.Count; index++)
            {
                var point = points[index];
                if (point == null || point.Index != index)
                { Unavailable(beam, "Control-point indices must be complete, ordered and contiguous from zero."); return; }
                if (!Finite(point.GantryAngleDegrees) || point.GantryAngleDegrees < 0 || point.GantryAngleDegrees > 360)
                { Unavailable(beam, "All gantry angles must be finite values from 0 to 360 degrees."); return; }
                if (!Finite(point.CumulativeMetersetWeight) || (index == 0 && point.CumulativeMetersetWeight != 0) ||
                    (index > 0 && point.CumulativeMetersetWeight < points[index - 1].CumulativeMetersetWeight))
                { Unavailable(beam, "Cumulative meterset weights must be finite, start at zero and never decrease."); return; }
            }
            double finalWeight = points[points.Count - 1].CumulativeMetersetWeight;
            if (finalWeight <= 0)
            { Unavailable(beam, "The final cumulative meterset weight must be positive."); return; }
            var durations = new double?[points.Count];
            var rates = new double?[points.Count];
            double totalSeconds = 0, nominalRate = beam.NominalDoseRateMuPerMin.Value;
            for (int index = 1; index < points.Count; index++)
            {
                double from = points[index - 1].GantryAngleDegrees % 360;
                double to = points[index].GantryAngleDegrees % 360;
                double directedAngle = clockwise ? to - from : from - to;
                if (directedAngle < 0) directedAngle += 360;
                // A reversal must never become a fictitious almost-full revolution.
                if (directedAngle >= 180)
                { Unavailable(beam, "A gantry interval reverses direction or spans an ambiguous 180 degrees or more."); return; }
                double deltaWeight = points[index].CumulativeMetersetWeight - points[index - 1].CumulativeMetersetWeight;
                // Normalize before multiplying to avoid intermediate beamMU * deltaWeight overflow.
                double deltaMu = beam.MetersetMu.Value * (deltaWeight / finalWeight);
                double gantrySeconds = directedAngle / profile.MaxGantrySpeedDegreesPerSecond;
                double muSeconds = (deltaMu / nominalRate) * 60;
                double duration = Math.Max(gantrySeconds, muSeconds);
                if (!Finite(deltaMu) || !Finite(duration) ||
                    (deltaWeight > 0 && (deltaMu <= 0 || muSeconds <= 0)) ||
                    (directedAngle > 0 && gantrySeconds <= 0))
                { Unavailable(beam, "Interval arithmetic is non-finite or outside representable precision."); return; }
                durations[index] = duration;
                if (duration > 0)
                {
                    double rate = (deltaMu / duration) * 60;
                    if (!Finite(rate) || (deltaMu > 0 && rate <= 0))
                    { Unavailable(beam, "Estimated dose-rate arithmetic is non-finite or outside representable precision."); return; }
                    // Round-off alone must not place an estimate above its explicit nominal-rate bound.
                    rates[index] = Math.Min(nominalRate, rate);
                }
                totalSeconds += duration;
                if (!Finite(totalSeconds))
                { Unavailable(beam, "Estimated beam duration exceeds representable precision."); return; }
            }
            for (int index = 0; index < points.Count; index++)
            {
                points[index].EstimatedDoseRateMuPerMin = rates[index];
                points[index].EstimatedSegmentDurationSeconds = durations[index];
            }
            beam.DoseRateEstimateStatus = "Estimated";
            beam.DoseRateEstimateProfile = profile.Id;
            beam.DoseRateEstimateMaxGantrySpeedDegreesPerSecond = profile.MaxGantrySpeedDegreesPerSecond;
            beam.EstimatedBeamDurationSeconds = totalSeconds;
            beam.DoseRateEstimateReason = "PlanCheck-style ESTIMATE from normalized cumulative MU, gantry angles, configured speed and nominal MU/min; " +
                "not time-resolved ESAPI or measured delivery. Ignores acceleration, MLC/jaw dynamics, dose ramping and holds. " +
                "Configured assumption: " + profile.Assumption;
        }

        public static void Apply(ReviewBeamAnalysis beam, DoseRateEstimationProfileCatalog catalog)
        {
            if (beam == null) throw new ArgumentNullException("beam");
            Clear(beam);
            if (catalog == null)
            { Unavailable(beam, "No dose-rate estimation profile catalog is configured."); return; }
            string reason;
            var profile = catalog.Resolve(beam, out reason);
            if (profile == null) { Unavailable(beam, reason); return; }
            Apply(beam, profile);
        }

        private static bool Is(string value, string expected)
        { return string.Equals(value, expected, StringComparison.OrdinalIgnoreCase); }
        private static bool Finite(double value) { return DoseRateEstimationProfileCatalog.Finite(value); }
        private static bool Positive(double? value) { return value.HasValue && Finite(value.Value) && value.Value > 0; }
        private static void Unavailable(ReviewBeamAnalysis beam, string reason)
        { beam.DoseRateEstimateReason = "ESTIMATE unavailable: " + reason; }
        private static void Clear(ReviewBeamAnalysis beam)
        {
            beam.DoseRateEstimateStatus = "Unavailable";
            beam.DoseRateEstimateReason = null;
            beam.DoseRateEstimateProfile = null;
            beam.EstimatedBeamDurationSeconds = null;
            beam.DoseRateEstimateMaxGantrySpeedDegreesPerSecond = null;
            if (beam.ControlPoints == null) return;
            foreach (var point in beam.ControlPoints)
            {
                if (point == null) continue;
                point.EstimatedDoseRateMuPerMin = null;
                point.EstimatedSegmentDurationSeconds = null;
            }
        }
    }
}
