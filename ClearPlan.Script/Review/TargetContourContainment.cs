using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;

namespace ClearPlan.Review
{
    /// <summary>Read-only planar polygon comparison; full polygon area, not only vertices or bounds.</summary>
    public static class TargetContourContainment
    {
        public static bool IsContained(IEnumerable<Point[]> innerContours, IEnumerable<Point[]> outerContours)
        {
            Geometry inner = PolygonGeometry(innerContours);
            Geometry outer = PolygonGeometry(outerContours);
            if (inner == null || outer == null) return false;
            Geometry outside = Geometry.Combine(inner, outer, GeometryCombineMode.Exclude,
                null, 0.000001, ToleranceType.Absolute);
            double area = outside.GetArea(0.000001, ToleranceType.Absolute);
            return !double.IsNaN(area) && !double.IsInfinity(area) && area <= 0.000001;
        }

        private static Geometry PolygonGeometry(IEnumerable<Point[]> contours)
        {
            var loops = (contours ?? Enumerable.Empty<Point[]>()).ToList();
            if (loops.Count == 0 || loops.Any(loop => loop == null || loop.Length < 3 ||
                loop.Any(point => double.IsNaN(point.X) || double.IsInfinity(point.X) ||
                    double.IsNaN(point.Y) || double.IsInfinity(point.Y)))) return null;
            var geometry = new StreamGeometry { FillRule = FillRule.EvenOdd };
            using (StreamGeometryContext context = geometry.Open())
            {
                foreach (Point[] loop in loops)
                {
                    context.BeginFigure(loop[0], true, true);
                    context.PolyLineTo(loop.Skip(1).ToArray(), true, false);
                }
            }
            geometry.Freeze();
            return geometry.GetArea(0.000001, ToleranceType.Absolute) > 0 ? geometry : null;
        }
    }
}
