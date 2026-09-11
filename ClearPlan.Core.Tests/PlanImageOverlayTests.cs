using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Reflection;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    public static class PlanImageOverlayTests
    {
#if PLAN_IMAGE_OVERLAY_TEST_MAIN
        public static int Main()
        {
            try { RunAll(); return 0; }
            catch (Exception ex) { Console.WriteLine(ex.Message); return 1; }
        }
#endif
        public static void RunAll()
        {
            var tests = new Action[] { DetachedContract, CoordinateMapping, NativeContoursAreClipped,
                MeshIntersection, CoplanarMeshBoundary, DuplicatedMeshVerticesDoNotDrawDiagonals, InvalidMeshRejected, DoseUnits,
                LinearIsodose, MissingDoseDoesNotCreateContours, SaddleIsDeterministic,
                IsodoseComplexityIsBounded, DoseGridExcludesOutsideSamples, NativeCaptureContract };
            int failed = 0;
            foreach (var test in tests)
            {
                try { test(); Console.WriteLine("PASS " + test.Method.Name); }
                catch (Exception ex) { failed++; Console.WriteLine("FAIL " + test.Method.Name + ": " + ex.GetBaseException().Message); }
            }
            if (failed != 0) throw new Exception(failed + " of " + tests.Length + " CT overlay tests failed.");
        }

        public static void DetachedContract()
        {
            Require(typeof(ReviewPlanImage).GetProperty("Overlays") != null, "Missing detached Overlays contract.");
            foreach (string name in new[] { "ReviewImageOverlay", "ReviewImagePath", "ReviewImagePoint" })
            {
                Type type = TypeOf(name);
                Require(type.GetProperties().All(p => !p.PropertyType.FullName.StartsWith("VMS.")), "Live vendor object in DTO.");
            }
        }

        public static void CoordinateMapping()
        {
            dynamic transverse = Plane("transversal", 11, 21, -10, -20, 30, 2, 2);
            dynamic coronal = Plane("coronal", 11, 21, -10, -20, 30, 2, 2);
            dynamic sagittal = Plane("sagittal", 11, 21, -10, -20, 30, 2, 2);
            dynamic p = transverse.ToPixel(new double[] { 0, 0, 30 }); Close(5, p.X); Close(10, p.Y);
            p = coronal.ToPixel(new double[] { 0, -20, 10 }); Close(5, p.X); Close(10, p.Y);
            p = sagittal.ToPixel(new double[] { -10, -10, 10 }); Close(5, p.X); Close(10, p.Y);
            double[] world = coronal.ToWorld(5.0, 10.0); Close(0, world[0]); Close(-20, world[1]); Close(10, world[2]);
        }

        public static void NativeContoursAreClipped()
        {
            dynamic plane = Plane("transversal", 11, 11, 0, 0, 0, 1, 1);
            var contours = new[] { new[] { new double[] { -5, 5, 0 }, new double[] { 5, 5, 0 }, new double[] { 5, 8, 0 } } };
            IList paths = (IList)Call("ProjectContours", plane, contours);
            Require(paths.Count == 3, "Closed contour must preserve three clipped edges without closing across the image boundary.");
            foreach (dynamic path in paths)
            foreach (dynamic point in path.Points) Require(point.X >= 0 && point.X <= 10 && point.Y >= 0 && point.Y <= 10, "Unclipped point.");
        }

        public static void MeshIntersection()
        {
            dynamic plane = Plane("coronal", 11, 11, 0, 0, 10, 1, 1);
            var vertices = new[] { new double[] { 2, -1, 8 }, new double[] { 8, 1, 8 }, new double[] { 2, 1, 2 } };
            IList paths = (IList)Call("IntersectMesh", plane, vertices, new int[] { 0, 1, 2 });
            Require(paths.Count == 1, "Triangle crossing must produce one segment.");
            dynamic path = paths[0];
            Require(path.Points.Count == 2, "No invented contour correspondence.");
            var coords = ((IEnumerable)path.Points).Cast<dynamic>().Select(p => ((double)p.X).ToString("0.0", System.Globalization.CultureInfo.InvariantCulture) + "/" + ((double)p.Y).ToString("0.0", System.Globalization.CultureInfo.InvariantCulture)).ToArray();
            Require(coords.Contains("5.0/2.0") && coords.Contains("2.0/5.0"), "Incorrect native mesh-plane intersection.");
        }

        public static void CoplanarMeshBoundary()
        {
            dynamic plane = Plane("transversal", 11, 11, 0, 0, 0, 1, 1);
            var v = new[] { new double[] { 2, 2, 0 }, new double[] { 8, 2, 0 }, new double[] { 8, 8, 0 }, new double[] { 2, 8, 0 } };
            IList paths = (IList)Call("IntersectMesh", plane, v, new int[] { 0, 1, 2, 0, 2, 3 });
            Require(paths.Count == 4, "Coplanar triangulation diagonal must not appear as a contour.");
        }

        public static void InvalidMeshRejected()
        {
            dynamic plane = Plane("transversal", 11, 11, 0, 0, 0, 1, 1);
            Reject(() => Call("IntersectMesh", plane, new[] { new double[] { 0, 0, 0 } }, new int[] { 0, 1, 2 }));
            Reject(() => Plane("oblique", 11, 11, 0, 0, 0, 1, 1));
            Reject(() => Plane("transversal", 11, 11, double.NaN, 0, 0, 1, 1));
        }

        public static void DuplicatedMeshVerticesDoNotDrawDiagonals()
        {
            dynamic plane = Plane("transversal", 11, 11, 0, 0, 0, 1, 1);
            var v = new[] { new double[] { 2, 2, 0 }, new double[] { 8, 2, 0 }, new double[] { 8, 8, 0 },
                new double[] { 2, 2, 0 }, new double[] { 8, 8, 0 }, new double[] { 2, 8, 0 } };
            IList paths = (IList)Call("IntersectMesh", plane, v, new int[] { 0, 1, 2, 3, 4, 5 });
            Require(paths.Count == 4, "Geometrically shared coplanar edges must cancel even with duplicated native mesh vertices.");
        }

        public static void DoseUnits()
        {
            Close(2, (double)Call("ToGy", 200.0, "cGy")); Close(2, (double)Call("ToGy", 2.0, "Gy"));
            Reject(() => Call("ToGy", 100.0, "Percent")); Reject(() => Call("ToGy", 2.0, "Unknown"));
            Reject(() => Call("ToGy", double.PositiveInfinity, "Gy"));
        }

        public static void LinearIsodose()
        {
            var samples = new double[] { 0, 10, 0, 10 };
            IList paths = (IList)Call("TraceIsodose", samples, 2, 2, 5.0, 101, 51, 100);
            Require(paths.Count == 1, "Expected single 5 Gy line.");
            dynamic p = paths[0]; Close(50, p.Points[0].X); Close(50, p.Points[1].X);
            Require(Math.Abs((double)p.Points[0].Y - (double)p.Points[1].Y) == 50, "Wrong dose-to-image scaling.");
        }

        public static void MissingDoseDoesNotCreateContours()
        {
            IList paths = (IList)Call("TraceIsodose", new double[] { 0, 10, double.NaN, 10 }, 2, 2, 5.0, 11, 11, 100);
            Require(paths.Count == 0, "NaN/outside-dose cells must stay empty, never zero-filled.");
        }

        public static void SaddleIsDeterministic()
        {
            IList paths = (IList)Call("TraceIsodose", new double[] { 10, 0, 0, 10 }, 2, 2, 5.0, 11, 11, 100);
            Require(paths.Count == 2, "An exact bilinear saddle has two crossing branches.");
            foreach (dynamic p in paths)
            {
                Close(5, ((double)p.Points[0].X + (double)p.Points[1].X) / 2);
                Close(5, ((double)p.Points[0].Y + (double)p.Points[1].Y) / 2);
            }
        }

        public static void IsodoseComplexityIsBounded()
        {
            Reject(() => Call("TraceIsodose", new double[] { 10, 0, 0, 10 }, 2, 2, 5.0, 11, 11, 1));
        }

        public static void NativeCaptureContract()
        {
            string source = File.ReadAllText(Path.Combine("ClearPlan.Script", "Review", "EsapiPlanImageBuilder.cs"));
            Require(source.Contains("IsDoseValid") && source.Contains("GetDoseProfile("), "Native valid-dose row capture missing.");
            Require(source.Contains("finally") && source.Contains("plan.DoseValuePresentation = previousPresentation.Value"), "Native dose display must be restored after absolute-unit sampling.");
            Require(source.Contains("GetContoursOnImagePlane(") && source.Contains("MeshGeometry"), "Native contour/mesh sources missing.");
            Require(source.Contains("Stopwatch") && source.Contains("MaxDoseSamples") && source.Contains("MaxStructures"), "Capture limits missing.");
            Require(source.Contains("lookup.Count >= 65536"), "Native HU conversion count must be bounded.");
            Require(!source.Contains("GetDoseToPoint(") && !source.Contains("Task.Run"), "Unbounded point-call or off-thread native work.");
            Require(source.Contains("doseGrid.Contains(expected)"), "Native zero-valued out-of-grid samples must be excluded explicitly.");
        }

        public static void DoseGridExcludesOutsideSamples()
        {
            var dirs = new[] { new double[] { 0, 1, 0 }, new double[] { -1, 0, 0 }, new double[] { 0, 0, -1 } };
            dynamic grid = Activator.CreateInstance(TypeOf("PlanImageDoseGrid"), new double[] { 10, 20, 30 }, dirs, new double[] { 2, 3, 4 }, new int[] { 3, 4, 5 });
            Require(grid.Contains(new double[] { 4, 22, 18 }), "Rotated dose basis must use physical coordinates.");
            Require(!grid.Contains(new double[] { 11, 22, 18 }), "Outside dose origin must not become valid zero dose.");
            Require(!grid.Contains(new double[] { 4, 25, 18 }), "Outside final dose voxel must be excluded.");
            Reject(() => Activator.CreateInstance(TypeOf("PlanImageDoseGrid"), new double[] { 0, 0, 0 }, new[] { new double[] { 1, 0, 0 }, new double[] { 1, 0, 0 }, new double[] { 0, 0, 1 } }, new double[] { 1, 1, 1 }, new int[] { 2, 2, 2 }));
        }

        private static Type TypeOf(string name)
        {
            Type type = typeof(ReviewPlanImage).Assembly.GetType("ClearPlan.Core.Review." + name);
            Require(type != null, "Missing " + name + "."); return type;
        }
        private static dynamic Plane(string kind, int width, int height, double x, double y, double z, double dx, double dy)
        { return Activator.CreateInstance(TypeOf("PlanImagePlane"), kind, width, height, x, y, z, dx, dy); }
        private static object Call(string method, params object[] arguments)
        { return TypeOf("PlanImageOverlayGeometry").GetMethod(method).Invoke(null, arguments); }
        private static void Require(bool value, string message) { if (!value) throw new Exception(message); }
        private static void Close(double expected, double actual) { Require(Math.Abs(expected - actual) < 0.000001, "Expected " + expected + "; got " + actual); }
        private static void Reject(Action action)
        {
            try { action(); }
            catch (Exception ex) { if (ex.GetBaseException() is ArgumentException || ex.GetBaseException() is InvalidOperationException) return; throw; }
            throw new Exception("Expected a safe rejection.");
        }
    }
}
