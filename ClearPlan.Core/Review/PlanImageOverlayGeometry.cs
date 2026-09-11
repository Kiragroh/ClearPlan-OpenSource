using System;
using System.Collections.Generic;
using System.Globalization;

namespace ClearPlan.Core.Review
{
    /// <summary>Axis-aligned DICOM plane: left-to-right +X/+Y; coronal/sagittal top-to-bottom -Z.</summary>
    public sealed class PlanImagePlane
    {
        private readonly double[] origin;
        public int Width { get; private set; }
        public int Height { get; private set; }
        public int NormalAxis { get; private set; }
        public int HorizontalAxis { get; private set; }
        public int VerticalAxis { get; private set; }
        public double VerticalSign { get; private set; }
        public double SpacingX { get; private set; }
        public double SpacingY { get; private set; }
        public double PlaneCoordinate { get { return origin[NormalAxis]; } }

        public PlanImagePlane(string kind, int width, int height, double x, double y, double z, double dx, double dy)
        {
            if (width < 2 || height < 2 || width > 2048 || height > 4096 || dx <= 0 || dy <= 0 ||
                !Finite(x) || !Finite(y) || !Finite(z) || !Finite(dx) || !Finite(dy)) throw new ArgumentException("Invalid CT plane.");
            if (kind != "transversal" && kind != "coronal" && kind != "sagittal") throw new ArgumentException("Unsupported CT plane.");
            Width = width; Height = height; origin = new[] { x, y, z }; SpacingX = dx; SpacingY = dy;
            HorizontalAxis = kind == "sagittal" ? 1 : 0;
            VerticalAxis = kind == "transversal" ? 1 : 2;
            NormalAxis = kind == "transversal" ? 2 : kind == "coronal" ? 1 : 0;
            VerticalSign = kind == "transversal" ? 1 : -1;
        }

        public ReviewImagePoint ToPixel(double[] world)
        {
            ValidatePoint(world);
            return new ReviewImagePoint { X = (world[HorizontalAxis] - origin[HorizontalAxis]) / SpacingX,
                Y = (world[VerticalAxis] - origin[VerticalAxis]) / (SpacingY * VerticalSign) };
        }

        public double[] ToWorld(double pixelX, double pixelY)
        {
            if (!Finite(pixelX) || !Finite(pixelY)) throw new ArgumentException("Invalid pixel coordinate.");
            var result = (double[])origin.Clone();
            result[HorizontalAxis] += pixelX * SpacingX; result[VerticalAxis] += pixelY * SpacingY * VerticalSign;
            return result;
        }

        internal static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        internal static void ValidatePoint(double[] point)
        {
            if (point == null || point.Length != 3 || !Finite(point[0]) || !Finite(point[1]) || !Finite(point[2]))
                throw new ArgumentException("Invalid native geometry point.");
        }
    }

    /// <summary>Native dose-grid center support in DICOM coordinates; outside samples never become zero dose.</summary>
    public sealed class PlanImageDoseGrid
    {
        private readonly double[] origin;
        private readonly double[][] directions;
        private readonly double[] spacing;
        private readonly int[] sizes;

        public PlanImageDoseGrid(double[] origin, double[][] directions, double[] spacing, int[] sizes)
        {
            PlanImagePlane.ValidatePoint(origin);
            if (directions == null || directions.Length != 3 || spacing == null || spacing.Length != 3 || sizes == null || sizes.Length != 3)
                throw new ArgumentException("Invalid native dose grid.");
            for (int axis = 0; axis < 3; axis++)
            {
                PlanImagePlane.ValidatePoint(directions[axis]);
                if (!PlanImagePlane.Finite(spacing[axis]) || spacing[axis] <= 0 || sizes[axis] < 2 || sizes[axis] > 4096)
                    throw new ArgumentException("Invalid dose grid size or resolution.");
            }
            for (int a = 0; a < 3; a++)
            for (int b = 0; b < 3; b++)
            {
                double dot = directions[a][0] * directions[b][0] + directions[a][1] * directions[b][1] + directions[a][2] * directions[b][2];
                if (Math.Abs(dot - (a == b ? 1 : 0)) > 0.0001) throw new ArgumentException("Dose directions are not orthonormal.");
            }
            this.origin = (double[])origin.Clone(); this.spacing = (double[])spacing.Clone(); this.sizes = (int[])sizes.Clone();
            this.directions = new[] { (double[])directions[0].Clone(), (double[])directions[1].Clone(), (double[])directions[2].Clone() };
        }

        public bool Contains(double[] point)
        {
            PlanImagePlane.ValidatePoint(point);
            for (int axis = 0; axis < 3; axis++)
            {
                double index = 0;
                for (int coordinate = 0; coordinate < 3; coordinate++) index += (point[coordinate] - origin[coordinate]) * directions[axis][coordinate] / spacing[axis];
                if (index < -0.000001 || index > sizes[axis] - 1 + 0.000001) return false;
            }
            return true;
        }
    }

    /// <summary>Vendor-free numeric processing. No synthetic fallback, inferred contour correspondence, or native object access.</summary>
    public static class PlanImageOverlayGeometry
    {
        public const int MaxMeshVertices = 200000;
        public const int MaxMeshTriangles = 300000;
        private const int MaxPaths = 50000;
        private const double Epsilon = 0.000001;

        public static double ToGy(double value, string unit)
        {
            if (!PlanImagePlane.Finite(value) || value < 0) throw new ArgumentException("Invalid absolute dose.");
            if (unit == "Gy") return value;
            if (unit == "cGy") return value / 100.0;
            throw new ArgumentException("Dose must expose Gy or cGy; relative/unknown units are not inferred.");
        }

        public static List<ReviewImagePath> ProjectContours(PlanImagePlane plane, double[][][] contours)
        {
            if (plane == null || contours == null) throw new ArgumentException("Missing native contours.");
            var result = new List<ReviewImagePath>(); int count = 0;
            foreach (var contour in contours)
            {
                if (contour == null || contour.Length < 3) continue;
                count += contour.Length;
                if (count > MaxMeshVertices) throw new ArgumentException("Native contour point limit exceeded.");
                for (int i = 0; i < contour.Length; i++)
                {
                    PlanImagePlane.ValidatePoint(contour[i]);
                    if (Math.Abs(contour[i][plane.NormalAxis] - plane.PlaneCoordinate) > 0.1)
                        throw new ArgumentException("Native contour is not on the captured CT plane.");
                    AddClipped(result, plane.ToPixel(contour[i]), plane.ToPixel(contour[(i + 1) % contour.Length]), plane.Width, plane.Height, MaxPaths);
                }
            }
            return result;
        }

        public static List<ReviewImagePath> IntersectMesh(PlanImagePlane plane, double[][] vertices, int[] triangleIndices)
        {
            if (plane == null || vertices == null || triangleIndices == null || triangleIndices.Length % 3 != 0 ||
                vertices.Length > MaxMeshVertices || triangleIndices.Length / 3 > MaxMeshTriangles) throw new ArgumentException("Invalid or oversized native mesh.");
            foreach (var point in vertices) PlanImagePlane.ValidatePoint(point);
            foreach (int index in triangleIndices) if (index < 0 || index >= vertices.Length) throw new ArgumentException("Invalid native mesh index.");
            var result = new List<ReviewImagePath>();
            var coplanarEdges = new Dictionary<string, int>();
            var coplanarVertices = new Dictionary<string, int[]>();
            var unique = new HashSet<string>();
            for (int triangle = 0; triangle < triangleIndices.Length; triangle += 3)
            {
                int[] ids = { triangleIndices[triangle], triangleIndices[triangle + 1], triangleIndices[triangle + 2] };
                double[] distances = { vertices[ids[0]][plane.NormalAxis] - plane.PlaneCoordinate,
                    vertices[ids[1]][plane.NormalAxis] - plane.PlaneCoordinate, vertices[ids[2]][plane.NormalAxis] - plane.PlaneCoordinate };
                if (Math.Abs(distances[0]) <= Epsilon && Math.Abs(distances[1]) <= Epsilon && Math.Abs(distances[2]) <= Epsilon)
                {
                    for (int edge = 0; edge < 3; edge++)
                    {
                        int a = ids[edge], b = ids[(edge + 1) % 3];
                        string pa = PointKey(plane.ToPixel(vertices[a])), pb = PointKey(plane.ToPixel(vertices[b]));
                        string key = string.CompareOrdinal(pa, pb) < 0 ? pa + ":" + pb : pb + ":" + pa;
                        int count; coplanarEdges.TryGetValue(key, out count); coplanarEdges[key] = count + 1; coplanarVertices[key] = new[] { a, b };
                    }
                    continue;
                }
                var cuts = new List<ReviewImagePoint>();
                for (int edge = 0; edge < 3; edge++)
                {
                    int next = (edge + 1) % 3; double da = distances[edge], db = distances[next];
                    if (Math.Abs(da) <= Epsilon) AddUniquePoint(cuts, plane.ToPixel(vertices[ids[edge]]));
                    if (da < -Epsilon && db > Epsilon || da > Epsilon && db < -Epsilon)
                    {
                        var a = vertices[ids[edge]]; var b = vertices[ids[next]]; double t = da / (da - db);
                        AddUniquePoint(cuts, plane.ToPixel(new[] { a[0] + t * (b[0] - a[0]), a[1] + t * (b[1] - a[1]), a[2] + t * (b[2] - a[2]) }));
                    }
                }
                if (cuts.Count == 2) AddUniqueSegment(result, unique, cuts[0], cuts[1], plane);
            }
            foreach (var edge in coplanarEdges)
            {
                if (edge.Value != 1) continue;
                var ids = coplanarVertices[edge.Key];
                AddUniqueSegment(result, unique, plane.ToPixel(vertices[ids[0]]), plane.ToPixel(vertices[ids[1]]), plane);
            }
            return result;
        }

        /// <summary>Marching squares with bilinear asymptotic saddle decision. Nonfinite cells are excluded, never zero-filled.</summary>
        public static List<ReviewImagePath> TraceIsodose(double[] samplesGy, int columns, int rows, double levelGy, int imageWidth, int imageHeight, int maximumSegments)
        {
            if (samplesGy == null || columns < 2 || rows < 2 || columns > 512 || rows > 512 || samplesGy.Length != columns * rows ||
                !PlanImagePlane.Finite(levelGy) || levelGy <= 0 || imageWidth < 2 || imageHeight < 2 || maximumSegments < 1 || maximumSegments > MaxPaths)
                throw new ArgumentException("Invalid isodose sampling grid.");
            var result = new List<ReviewImagePath>();
            double sx = (imageWidth - 1.0) / (columns - 1), sy = (imageHeight - 1.0) / (rows - 1);
            for (int row = 0; row < rows - 1; row++)
            for (int column = 0; column < columns - 1; column++)
            {
                double[] v = { samplesGy[row * columns + column] - levelGy, samplesGy[row * columns + column + 1] - levelGy,
                    samplesGy[(row + 1) * columns + column + 1] - levelGy, samplesGy[(row + 1) * columns + column] - levelGy };
                if (!PlanImagePlane.Finite(v[0]) || !PlanImagePlane.Finite(v[1]) || !PlanImagePlane.Finite(v[2]) || !PlanImagePlane.Finite(v[3])) continue;
                double[] cx = { column * sx, (column + 1) * sx, (column + 1) * sx, column * sx };
                double[] cy = { row * sy, row * sy, (row + 1) * sy, (row + 1) * sy };
                var cuts = new List<ReviewImagePoint>();
                for (int edge = 0; edge < 4; edge++)
                {
                    int next = (edge + 1) % 4;
                    if ((v[edge] >= 0) == (v[next] >= 0)) continue;
                    double t = v[edge] / (v[edge] - v[next]);
                    cuts.Add(new ReviewImagePoint { X = cx[edge] + t * (cx[next] - cx[edge]), Y = cy[edge] + t * (cy[next] - cy[edge]) });
                }
                if (cuts.Count == 2) AddClipped(result, cuts[0], cuts[1], imageWidth, imageHeight, maximumSegments);
                else if (cuts.Count == 4)
                {
                    double determinant = v[0] * v[2] - v[1] * v[3];
                    if (determinant == 0)
                    {
                        AddClipped(result, cuts[0], cuts[2], imageWidth, imageHeight, maximumSegments);
                        AddClipped(result, cuts[1], cuts[3], imageWidth, imageHeight, maximumSegments);
                    }
                    else
                    {
                        int first = determinant > 0 ? 1 : 3, other = determinant > 0 ? 3 : 1;
                        AddClipped(result, cuts[0], cuts[first], imageWidth, imageHeight, maximumSegments);
                        AddClipped(result, cuts[2], cuts[other], imageWidth, imageHeight, maximumSegments);
                    }
                }
            }
            return result;
        }

        private static void AddUniquePoint(List<ReviewImagePoint> points, ReviewImagePoint point)
        {
            foreach (var existing in points) if (Math.Abs(existing.X - point.X) < Epsilon && Math.Abs(existing.Y - point.Y) < Epsilon) return;
            points.Add(point);
        }

        private static void AddUniqueSegment(List<ReviewImagePath> paths, HashSet<string> unique, ReviewImagePoint a, ReviewImagePoint b, PlanImagePlane plane)
        {
            string pa = PointKey(a), pb = PointKey(b);
            string key = string.CompareOrdinal(pa, pb) < 0 ? pa + ":" + pb : pb + ":" + pa;
            if (unique.Add(key)) AddClipped(paths, a, b, plane.Width, plane.Height, MaxPaths);
        }

        private static string PointKey(ReviewImagePoint point)
        { return point.X.ToString("F6", CultureInfo.InvariantCulture) + "," + point.Y.ToString("F6", CultureInfo.InvariantCulture); }

        // Liang-Barsky clipping prevents off-grid segments and spurious connectors along CT boundaries.
        private static void AddClipped(List<ReviewImagePath> paths, ReviewImagePoint a, ReviewImagePoint b, int width, int height, int maximum)
        {
            double dx = b.X - a.X, dy = b.Y - a.Y, start = 0, stop = 1;
            if (!Clip(-dx, a.X, ref start, ref stop) || !Clip(dx, width - 1 - a.X, ref start, ref stop) ||
                !Clip(-dy, a.Y, ref start, ref stop) || !Clip(dy, height - 1 - a.Y, ref start, ref stop)) return;
            if (Math.Abs(dx * (stop - start)) < Epsilon && Math.Abs(dy * (stop - start)) < Epsilon) return;
            if (paths.Count >= maximum) throw new InvalidOperationException("Overlay segment limit exceeded; partial geometry is withheld.");
            paths.Add(new ReviewImagePath { Points = new List<ReviewImagePoint> {
                new ReviewImagePoint { X = a.X + start * dx, Y = a.Y + start * dy },
                new ReviewImagePoint { X = a.X + stop * dx, Y = a.Y + stop * dy } } });
        }

        private static bool Clip(double direction, double distance, ref double start, ref double stop)
        {
            if (Math.Abs(direction) < Epsilon) return distance >= 0;
            double t = distance / direction;
            if (direction < 0) { if (t > stop) return false; if (t > start) start = t; }
            else { if (t < start) return false; if (t < stop) stop = t; }
            return true;
        }
    }
}
