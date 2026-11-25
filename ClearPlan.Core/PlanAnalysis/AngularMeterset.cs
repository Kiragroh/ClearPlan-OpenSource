using System;

namespace ClearPlan.Core.PlanAnalysis
{
    // Angular meterset density, not time-resolved dose rate or a complexity score.
    public static class AngularMeterset
    {
        public static double? Travel(ReviewBeamAnalysis beam, ReviewControlPointSample a, ReviewControlPointSample b)
        {
            if (beam == null || a == null || b == null || !Finite(a.GantryAngleDegrees) || !Finite(b.GantryAngleDegrees)) return null;
            int sign = string.Equals(beam.GantryDirection, "Clockwise", StringComparison.OrdinalIgnoreCase) ? 1 :
                string.Equals(beam.GantryDirection, "CounterClockwise", StringComparison.OrdinalIgnoreCase) ? -1 : 0;
            if (sign == 0) return null;
            double delta = ((b.GantryAngleDegrees - a.GantryAngleDegrees) * sign % 360 + 360) % 360;
            // Repeated angles do not prove a full revolution; never invent one.
            return delta < 1e-6 ? (double?)null : sign * delta;
        }

        public static double? Density(ReviewBeamAnalysis beam, ReviewControlPointSample a, ReviewControlPointSample b)
        {
            double? travel = Travel(beam, a, b);
            if (!travel.HasValue || beam.ControlPoints == null || beam.ControlPoints.Count < 2 ||
                !beam.MetersetMu.HasValue || !Finite(beam.MetersetMu.Value) || beam.MetersetMu.Value < 0) return null;
            double final = beam.ControlPoints[beam.ControlPoints.Count - 1].CumulativeMetersetWeight;
            double delta = b.CumulativeMetersetWeight - a.CumulativeMetersetWeight;
            if (!Finite(final) || final <= 0 || !Finite(delta) || delta < 0 ||
                a.CumulativeMetersetWeight < 0 || b.CumulativeMetersetWeight > final + 1e-8) return null;
            double value = beam.MetersetMu.Value * delta / final / Math.Abs(travel.Value);
            return Finite(value) ? (double?)value : null;
        }
        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
    }
}
