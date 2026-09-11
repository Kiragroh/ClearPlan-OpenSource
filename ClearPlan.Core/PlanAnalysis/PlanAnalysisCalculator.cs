using System;
using System.Collections.Generic;
using System.Linq;

namespace ClearPlan.Core.PlanAnalysis
{
    /// <summary>
    /// Research geometry, not a delivery reconstruction. Each beam's PAM is the
    /// cumulative-meterset-weighted blocked target BEV fraction with endpoint sampling.
    /// Native plan aggregation uses PlanCheck's Beam.WeightFactor convention; imported
    /// and synthetic inputs retain their explicitly declared meterset-MU weighting.
    /// </summary>
    public static class PlanAnalysisCalculator
    {
        public const string PamDefinition = "PAM = sum(MU-weight * blocked target BEV area / total target BEV area) / sum(MU-weight). DOI:10.1002/mp.70144. Endpoint trapezoidal sampling; no clinical threshold.";
        public const string SamplingDefinition = "All control points; each endpoint receives half the MU of each adjacent interval. Geometry in isocenter-plane mm; areas in cm2. Small-aperture fraction is the MU-weighted proportion of endpoint apertures below the stated area, not a clinical limit.";
        public const string NativePamDefinition = "PAM per beam: blocked target BEV fraction with trapezoidal cumulative-meterset endpoint weights. Plan PAM: sum(beam PAM * Beam.WeightFactor) / sum(Beam.WeightFactor), matching the PlanCheck Metrics.PAM weighting convention. No clinical threshold.";

        public static ReviewPlanAnalysis Calculate(ReviewPlanAnalysis plan)
        {
            if (plan == null) throw new ArgumentNullException("plan");
            plan.Beams = plan.Beams ?? new List<ReviewBeamAnalysis>();
            plan.Warnings = plan.Warnings ?? new List<string>();
            plan.TotalMetersetMu = null;
            plan.MuPerGy = null;
            plan.Pam = null;
            plan.MeanApertureAreaCm2 = null;
            plan.SmallApertureFraction = null;
            if (!Finite(plan.SmallApertureThresholdCm2) || plan.SmallApertureThresholdCm2 <= 0)
                throw new ArgumentOutOfRangeException("SmallApertureThresholdCm2");
            foreach (var beam in plan.Beams) CalculateBeam(beam, plan.SmallApertureThresholdCm2);
            if (plan.Beams.Count > 0 && plan.Beams.All(b => ValidNonnegative(b.MetersetMu)))
            {
                double total = plan.Beams.Sum(b => b.MetersetMu.Value);
                if(Finite(total)) plan.TotalMetersetMu = total;
            }
            if (ValidNonnegative(plan.TotalMetersetMu) && ValidPositive(plan.DosePerFractionGy))
            {
                double ratio = plan.TotalMetersetMu.Value / plan.DosePerFractionGy.Value;
                if(Finite(ratio)) plan.MuPerGy = ratio;
            }
            string pamWeightReason = null;
            if (ValidPositive(plan.TotalMetersetMu))
            {
                if (plan.PamWeightingMode == "BeamWeightFactor")
                    plan.Pam = NativePamWeighted(plan.Beams, out pamWeightReason);
                else if (plan.PamWeightingMode == "MetersetMu")
                    plan.Pam = PlanWeighted(plan.Beams, b => b.Pam, plan.TotalMetersetMu.Value);
                else pamWeightReason = "Unknown PAM beam-weighting mode; no weighting basis is inferred.";
                plan.MeanApertureAreaCm2 = PlanWeighted(plan.Beams, b => b.MeanApertureAreaCm2, plan.TotalMetersetMu.Value);
                plan.SmallApertureFraction = PlanWeighted(plan.Beams, b => b.SmallApertureFraction, plan.TotalMetersetMu.Value);
            }
            if (string.IsNullOrWhiteSpace(plan.TargetStructureId)) plan.Pam = null;
            plan.PamStatus = plan.Pam.HasValue ? "available" : "unavailable";
            if (plan.Pam.HasValue) plan.PamReason = plan.PamWeightingMode == "BeamWeightFactor" ? NativePamDefinition : PamDefinition;
            else if (string.IsNullOrWhiteSpace(plan.TargetStructureId))
                plan.PamReason = "No explicit, unambiguous target structure selected. PAM is structure-specific.";
            else if (!ValidPositive(plan.TotalMetersetMu))
                plan.PamReason = "Complete positive treatment-beam MU is unavailable; no partial-plan PAM is reported.";
            else if (pamWeightReason != null) plan.PamReason = pamWeightReason;
            else plan.PamReason = string.Join(" ", plan.Beams.Where(b => b.MetersetMu > 0 && !b.Pam.HasValue)
                .Select(b => b.PamReason ?? "Incomplete target projection or aperture geometry.").Distinct());
            if (!string.IsNullOrWhiteSpace(plan.TargetSelectionProvenance))
                plan.PamReason += " Target selection: " + plan.TargetSelectionProvenance;
            return plan;
        }

        private static double? NativePamWeighted(List<ReviewBeamAnalysis> beams, out string reason)
        {
            reason = null;
            if (beams.Count == 0 || !beams.All(b => ValidNonnegative(b.PamBeamWeightFactor)))
            { reason = "Complete finite nonnegative Beam.WeightFactor values are required for native PlanCheck PAM; no MU fallback is used."; return null; }
            double total = beams.Sum(b => b.PamBeamWeightFactor.Value);
            if (!Finite(total) || total <= 0)
            { reason = "A finite positive sum of Beam.WeightFactor is required for native PlanCheck PAM."; return null; }
            var active = beams.Where(b => b.PamBeamWeightFactor > 0).ToList();
            if (!active.All(b => b.Pam.HasValue))
            { reason = string.Join(" ", active.Where(b => !b.Pam.HasValue).Select(b => b.PamReason ?? "Incomplete beam PAM.").Distinct()); return null; }
            return active.Sum(b => b.Pam.Value * (b.PamBeamWeightFactor.Value / total));
        }

        private static double? PlanWeighted(IEnumerable<ReviewBeamAnalysis> beams,
            Func<ReviewBeamAnalysis,double?> value, double total)
        {
            var active = beams.Where(b => b.MetersetMu > 0).ToList();
            return active.All(b => value(b).HasValue)
                ? (double?)(active.Sum(b => (b.MetersetMu.Value/total) * value(b).Value)) : null;
        }

        private static void CalculateBeam(ReviewBeamAnalysis beam, double smallArea)
        {
            beam.Pam = beam.MeanApertureAreaCm2 = beam.SmallApertureFraction = null;
            beam.MinimumApertureAreaCm2 = beam.MaximumApertureAreaCm2 = null;
            beam.ControlPoints = beam.ControlPoints ?? new List<ReviewControlPointSample>();
            foreach (var cp in beam.ControlPoints)
            {
                cp.IncrementalMetersetMu = cp.MetricMetersetWeightMu = null;
                cp.ApertureAreaCm2 = cp.TargetAreaCm2 = cp.BlockedTargetFraction = null;
                bool hasProjectedStrips = cp.TargetProjectionStrips != null && cp.TargetProjectionStrips.Count > 0;
                if (hasProjectedStrips)
                {
                    try
                    {
                        foreach (var strip in cp.TargetProjectionStrips) ValidateRectangle(strip);
                        cp.TargetAreaCm2 = MeshTargetProjector.AreaMm2(cp.TargetProjectionStrips)/100;
                        if(!Finite(cp.TargetAreaCm2.Value)) throw new ArgumentException("Invalid projected area.");
                    }
                    catch (ArgumentException) { hasProjectedStrips = false; cp.TargetAreaCm2 = null; }
                }
                if (cp.Aperture == null) continue;
                try
                {
                    cp.Aperture.EffectiveOpenings = BuildOpenings(cp.Aperture);
                    cp.ApertureAreaCm2 = cp.Aperture.EffectiveOpenings.Sum(Area) / 100;
                    if(!Finite(cp.ApertureAreaCm2.Value)) throw new ArgumentException("Invalid aperture area.");
                    cp.GeometryReason = null;
                    double target, exposed;
                    if (hasProjectedStrips)
                    {
                        target = cp.TargetAreaCm2.Value*100;
                        exposed = MeshTargetProjector.IntersectionAreaMm2(cp.TargetProjectionStrips,cp.Aperture.EffectiveOpenings);
                    }
                    else
                    {
                        if (cp.TargetOutlines == null || cp.TargetOutlines.Count == 0) continue;
                        TargetAreas(cp.TargetOutlines, cp.Aperture.EffectiveOpenings, out target, out exposed);
                    }
                    if(!Finite(target)||!Finite(exposed)) throw new ArgumentException("Invalid target area.");
                    if (target <= 1e-9)
                    { cp.TargetProjectionReason = "Target projection has zero area."; continue; }
                    cp.TargetAreaCm2 = target / 100;
                    cp.BlockedTargetFraction = Math.Max(0, Math.Min(1, 1 - exposed / target));
                }
                catch (ArgumentException)
                {
                    cp.GeometryReason = "Incomplete or invalid physical aperture/target geometry; no inferred leaf layout.";
                    cp.Aperture.EffectiveOpenings.Clear();
                    cp.ApertureAreaCm2 = cp.TargetAreaCm2 = cp.BlockedTargetFraction = null;
                }
            }
            var known = beam.ControlPoints.Where(c => c.Aperture != null).ToList();
            if (known.Count > 0)
            {
                beam.MlcLayerCount = known.Max(c => c.Aperture.Layers == null ? 0 : c.Aperture.Layers.Count);
                beam.HasJaws = known.Any(c => c.Aperture.Jaws != null);
            }
            if (!SetMuWeights(beam))
            {
                beam.PamReason = "Missing or invalid beam MU/control-point meterset sequence.";
                beam.GeometryStatus = "unavailable";
                return;
            }
            var weighted = beam.ControlPoints.Where(c => c.MetricMetersetWeightMu > 0).ToList();
            if (weighted.Count == 0)
            { beam.PamReason = "Beam has no positive-MU intervals."; return; }
            beam.MeanApertureAreaCm2 = BeamWeighted(weighted, c => c.ApertureAreaCm2, beam.MetersetMu.Value);
            beam.Pam = BeamWeighted(weighted, c => c.BlockedTargetFraction, beam.MetersetMu.Value);
            if (beam.MeanApertureAreaCm2.HasValue)
            {
                beam.SmallApertureFraction = weighted.Sum(c => c.ApertureAreaCm2 < smallArea ? c.MetricMetersetWeightMu.Value/beam.MetersetMu.Value : 0);
                beam.MinimumApertureAreaCm2 = weighted.Min(c => c.ApertureAreaCm2.Value);
                beam.MaximumApertureAreaCm2 = weighted.Max(c => c.ApertureAreaCm2.Value);
                beam.GeometryStatus = "available";
                beam.GeometryReason = SamplingDefinition;
            }
            else
            {
                beam.GeometryStatus = "unavailable";
                beam.GeometryReason = weighted.Where(c => !c.ApertureAreaCm2.HasValue)
                    .Select(c => c.GeometryReason).FirstOrDefault(r => !string.IsNullOrWhiteSpace(r))
                    ?? "Physical leaf boundaries and every MLC layer are required; no fabricated jaws or leaf widths.";
            }
            beam.PamReason = beam.Pam.HasValue ? PamDefinition : weighted.Where(c => !c.BlockedTargetFraction.HasValue)
                .Select(c => c.TargetProjectionReason ?? c.GeometryReason).FirstOrDefault(r => !string.IsNullOrWhiteSpace(r))
                ?? "Aperture and target BEV projection are required at every MU-weighted control point.";
        }

        private static double? BeamWeighted(List<ReviewControlPointSample> points,
            Func<ReviewControlPointSample,double?> value, double total)
        {
            return points.All(p => value(p).HasValue)
                ? (double?)(points.Sum(p => value(p).Value * (p.MetricMetersetWeightMu.Value/total))) : null;
        }

        private static bool SetMuWeights(ReviewBeamAnalysis beam)
        {
            var points = beam.ControlPoints;
            if (!ValidNonnegative(beam.MetersetMu) || points.Count < 2 ||
                !points.All(p => Finite(p.CumulativeMetersetWeight)) ||
                Math.Abs(points[0].CumulativeMetersetWeight) > 1e-9 ||
                points[points.Count-1].CumulativeMetersetWeight <= 0) return false;
            for (int i = 1; i < points.Count; i++)
                if (points[i].CumulativeMetersetWeight < points[i-1].CumulativeMetersetWeight) return false;
            foreach (var point in points) point.MetricMetersetWeightMu = 0;
            points[0].IncrementalMetersetMu = 0;
            double final = points[points.Count-1].CumulativeMetersetWeight;
            for (int i = 1; i < points.Count; i++)
            {
                double mu = beam.MetersetMu.Value * ((points[i].CumulativeMetersetWeight - points[i-1].CumulativeMetersetWeight) / final);
                points[i].IncrementalMetersetMu = mu;
                points[i].MetricMetersetWeightMu += mu / 2;
                points[i-1].MetricMetersetWeightMu += mu / 2;
            }
            return true;
        }

        // Each physical layer is a union of disjoint leaf strips. The effective aperture
        // is their intersection, also intersected with real jaws only when supplied.
        public static List<ApertureRectangle> BuildOpenings(ApertureGeometry aperture)
        {
            if (aperture == null) throw new ArgumentNullException("aperture");
            List<ApertureRectangle> result = null;
            if (aperture.Jaws != null)
            {
                ValidateRectangle(aperture.Jaws);
                result = new List<ApertureRectangle> { aperture.Jaws };
            }
            if (aperture.FixedBoundingBox != null)
            {
                ValidateRectangle(aperture.FixedBoundingBox);
                var fixedBoundary = new List<ApertureRectangle> { aperture.FixedBoundingBox };
                result = result == null ? fixedBoundary : Intersect(result,fixedBoundary);
            }
            foreach (var layer in aperture.Layers ?? new List<ApertureLayer>())
            {
                ValidateLayer(layer);
                var strips = new List<ApertureRectangle>();
                for (int i = 0; i < layer.Bank1PositionsMm.Length; i++)
                {
                    if (layer.Bank2PositionsMm[i] <= layer.Bank1PositionsMm[i]) continue;
                    strips.Add(layer.LeafTravelAxis == "Y"
                        ? new ApertureRectangle(layer.LeafBoundariesMm[i], layer.Bank1PositionsMm[i], layer.LeafBoundariesMm[i+1], layer.Bank2PositionsMm[i])
                        : new ApertureRectangle(layer.Bank1PositionsMm[i], layer.LeafBoundariesMm[i], layer.Bank2PositionsMm[i], layer.LeafBoundariesMm[i+1]));
                }
                result = result == null ? strips : Intersect(result, strips);
            }
            if (result == null) throw new ArgumentException("No physical aperture definition.");
            return result.Where(r => Area(r) > 0).ToList();
        }

        private static List<ApertureRectangle> Intersect(List<ApertureRectangle> first, List<ApertureRectangle> second)
        {
            var result = new List<ApertureRectangle>();
            foreach (var a in first) foreach (var b in second)
            {
                double x1 = Math.Max(a.X1,b.X1), x2 = Math.Min(a.X2,b.X2);
                double y1 = Math.Max(a.Y1,b.Y1), y2 = Math.Min(a.Y2,b.Y2);
                if (x2 > x1 && y2 > y1) result.Add(new ApertureRectangle(x1,y1,x2,y2));
            }
            return result;
        }

        private static void ValidateLayer(ApertureLayer layer)
        {
            if (layer == null || layer.LeafBoundariesMm == null || layer.Bank1PositionsMm == null ||
                layer.Bank2PositionsMm == null || layer.Bank1PositionsMm.Length == 0 ||
                layer.LeafBoundariesMm.Length != layer.Bank1PositionsMm.Length + 1 ||
                layer.Bank1PositionsMm.Length != layer.Bank2PositionsMm.Length ||
                !(layer.LeafTravelAxis == "X" || layer.LeafTravelAxis == "Y") ||
                !layer.LeafBoundariesMm.All(Finite) || !layer.Bank1PositionsMm.All(Finite) ||
                !layer.Bank2PositionsMm.All(Finite)) throw new ArgumentException("Invalid leaf geometry.");
            for (int i = 1; i < layer.LeafBoundariesMm.Length; i++)
                if (layer.LeafBoundariesMm[i] <= layer.LeafBoundariesMm[i-1])
                    throw new ArgumentException("Leaf boundaries must be strictly increasing.");
        }

        private static void ValidateRectangle(ApertureRectangle rect)
        {
            if (!Finite(rect.X1) || !Finite(rect.X2) || !Finite(rect.Y1) || !Finite(rect.Y2) ||
                rect.X2 < rect.X1 || rect.Y2 < rect.Y1) throw new ArgumentException("Invalid jaw rectangle.");
        }

        // Exact piecewise-linear scan-line area for compound closed outlines (even-odd fill,
        // including holes and disconnected components). Breaks at vertices, edge crossings,
        // and aperture boundaries, so midpoint integration is exact within each strip.
        private static void TargetAreas(List<List<BeamPoint>> outlines, List<ApertureRectangle> openings,
            out double target, out double exposed)
        {
            var edges = new List<Edge>();
            var ys = new List<double>();
            foreach (var loop in outlines)
            {
                if (loop == null || loop.Count < 3 || loop.Any(p => p == null || !Finite(p.X) || !Finite(p.Y)))
                    throw new ArgumentException("Invalid target outline.");
                for (int i = 0; i < loop.Count; i++)
                {
                    var a = loop[i]; var b = loop[(i+1)%loop.Count];
                    ys.Add(a.Y);
                    if (Math.Abs(a.Y-b.Y)>1e-12) edges.Add(new Edge(a,b));
                }
            }
            foreach (var rect in openings)
            {
                ys.Add(rect.Y1); ys.Add(rect.Y2);
                foreach (var edge in edges)
                { edge.AddCrossingX(rect.X1,ys); edge.AddCrossingX(rect.X2,ys); }
            }
            for (int i = 0; i < edges.Count; i++) for (int j = i+1; j < edges.Count; j++)
            {
                var a = edges[i]; var b = edges[j];
                if (Math.Abs(a.Slope-b.Slope)<1e-12) continue;
                double y = (b.Intercept-a.Intercept)/(a.Slope-b.Slope);
                if (y > Math.Max(a.LowY,b.LowY) && y < Math.Min(a.HighY,b.HighY)) ys.Add(y);
            }
            ys = ys.Distinct().OrderBy(y => y).ToList();
            target = exposed = 0;
            for (int i = 1; i < ys.Count; i++)
            {
                double height = ys[i]-ys[i-1]; if (height <= 1e-12) continue;
                double y = (ys[i]+ys[i-1])/2;
                var xs = edges.Where(e => y>e.LowY && y<e.HighY).Select(e => e.X(y)).OrderBy(x => x).ToList();
                if (xs.Count%2 != 0) throw new ArgumentException("Target outline topology is invalid.");
                for (int j = 0; j < xs.Count; j+=2)
                {
                    target += (xs[j+1]-xs[j])*height;
                    foreach (var rect in openings.Where(r => y>r.Y1 && y<r.Y2))
                        exposed += Math.Max(0,Math.Min(xs[j+1],rect.X2)-Math.Max(xs[j],rect.X1))*height;
                }
            }
        }

        private sealed class Edge
        {
            internal readonly double Slope, Intercept, LowY, HighY;
            internal Edge(BeamPoint a, BeamPoint b)
            { Slope = (b.X-a.X)/(b.Y-a.Y); Intercept = a.X-Slope*a.Y; LowY = Math.Min(a.Y,b.Y); HighY = Math.Max(a.Y,b.Y); }
            internal double X(double y) { return Slope*y+Intercept; }
            internal void AddCrossingX(double x,List<double> ys)
            { if (Math.Abs(Slope)<1e-12) return; double y = (x-Intercept)/Slope; if (y>LowY && y<HighY) ys.Add(y); }
        }

        private static double Area(ApertureRectangle r)
        {
            double area=Math.Max(0,r.X2-r.X1)*Math.Max(0,r.Y2-r.Y1);
            if(!Finite(area)) throw new ArgumentException("Invalid aperture area.");
            return area;
        }
        internal static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static bool ValidPositive(double? value) { return value.HasValue && Finite(value.Value) && value.Value > 0; }
        private static bool ValidNonnegative(double? value) { return value.HasValue && Finite(value.Value) && value.Value >= 0; }
    }
}
