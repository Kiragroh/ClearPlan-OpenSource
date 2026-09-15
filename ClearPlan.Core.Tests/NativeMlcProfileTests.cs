using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;
using Newtonsoft.Json;

namespace ClearPlan.Core.Tests
{
    internal static class NativeMlcProfileTests
    {
        public static void RunAll()
        {
            ExplicitNativeMappingIsDetached();
            ExactModelAndShapeAreRequired();
            UnverifiedDualLayerLayoutIsNeverGuessed();
            ExplicitStaggeredLayersUseIndependentIndices();
            InvalidProfilesRejectSafely();
            InvalidNativeValuesStayUnavailable();
            PhysicalJawsAndFixedLimitsStayDistinct();
            PublicReferenceProfilesHaveExpectedGeometry();
            Sx2ReferenceKeepsDistinctNativeLayers();
            ConfiguredMaximumFieldIsDetachedValidatedAndClipsTheOpening();
            Sx2StaggeredOpeningsIntersectBothLayers();
            EsapiNativeAdapterHasReadOnlyAtomicContract();
        }

        private static void ExplicitNativeMappingIsDetached()
        {
            var catalog = Catalog(); var positions = Positions(2);
            var jaws = new ApertureRectangle(-5, -20, 5, 20);
            var result = NativeMlcGeometryMapper.Map(catalog, "Synthetic MLC", positions, jaws);
            TestAssert.Equal("available", result.Code); TestAssert.NotNull(result.Geometry);
            TestAssert.Equal(2, result.NativeBankCount); TestAssert.Equal(2, result.NativeLeafCount);
            TestAssert.Equal(-10d, result.Geometry.Layers[0].Bank1PositionsMm[0]);
            TestAssert.Equal(10d, result.Geometry.Layers[0].Bank2PositionsMm[0]);
            positions[0, 0] = -99; jaws.X1 = -99;
            catalog.Profiles[0].Layers[0].LeafBoundariesMm[0] = -99;
            TestAssert.Equal(-10d, result.Geometry.Layers[0].Bank1PositionsMm[0]);
            TestAssert.Equal(-5d, result.Geometry.Jaws.X1);
            TestAssert.Equal(-10d, result.Geometry.Layers[0].LeafBoundariesMm[0]);
            TestAssert.False(JsonConvert.SerializeObject(result).Contains("Bank1PositionsMm"), "Native clinical geometry must not serialize through diagnostics.");
        }
        private static void ExactModelAndShapeAreRequired()
        {
            var catalog = Catalog();
            TestAssert.Equal("profile_missing", Map(catalog, "synthetic mlc").Code);
            TestAssert.Equal("profile_missing", Map(catalog, "Synthetic MLC ").Code);
            TestAssert.Equal("profile_missing", Map(catalog, null).Code);
            TestAssert.Equal("profile_missing", Map(null).Code);
            var badShape = NativeMlcGeometryMapper.Map(catalog, "Synthetic MLC", new float[1, 2], Jaws());
            TestAssert.Equal("native_shape", badShape.Code); TestAssert.Equal(1, badShape.NativeBankCount);
            TestAssert.Equal("native_shape", NativeMlcGeometryMapper.Map(catalog, "Synthetic MLC", Positions(3), Jaws()).Code);
            TestAssert.Equal("native_shape", NativeMlcGeometryMapper.Map(catalog, "Synthetic MLC", null, Jaws()).Code);
        }
        private static void UnverifiedDualLayerLayoutIsNeverGuessed()
        {
            var result = NativeMlcGeometryMapper.Map(Catalog(), "Synthetic unknown dual-layer MLC", Positions(57), Jaws());
            TestAssert.Equal("profile_missing", result.Code); TestAssert.Equal(57, result.NativeLeafCount);
            TestAssert.True(result.Geometry == null);
            TestAssert.False(result.Reason.Contains("DICOM") || result.Reason.Contains("RTPLAN"), "Native diagnostics must not require an export fallback.");
            TestAssert.True(result.Reason.Contains("2x57"), "Observed native array shape must be visible.");
        }
        private static void ExplicitStaggeredLayersUseIndependentIndices()
        {
            var profile = Profile(); profile.NativeLeafCount = 4; profile.JawMode = "None";
            profile.Layers = new List<NativeMlcGeometryLayer>
            {
                Layer("Proximal synthetic", new[] { 2, 0 }, new[] { -15d, -5d, 5d }),
                Layer("Distal synthetic", new[] { 3, 1 }, new[] { -10d, 0d, 10d })
            };
            var catalog = Parse(profile); var positions = Positions(4); positions[0, 2] = -3; positions[1, 3] = 4;
            var result = NativeMlcGeometryMapper.Map(catalog, profile.Model, positions, null);
            TestAssert.NotNull(result.Geometry); TestAssert.Equal(2, result.Geometry.Layers.Count);
            TestAssert.Equal(-3d, result.Geometry.Layers[0].Bank1PositionsMm[0]);
            TestAssert.Equal(4d, result.Geometry.Layers[1].Bank2PositionsMm[0]);
            TestAssert.True(result.Geometry.Jaws == null && result.Geometry.FixedBoundingBox == null);
        }
        private static void InvalidProfilesRejectSafely()
        {
            var mutations = new Action<NativeMlcGeometryProfile>[]
            {
                p => p.NativeLeafCount = 0,
                p => p.Model = " ",
                p => p.JawMode = "guess",
                p => p.Layers = null,
                p => p.Layers[0].LeafTravelAxis = "Z",
                p => p.Layers[0].LeafBoundariesMm = new[] { -10d, 0d },
                p => p.Layers[0].LeafBoundariesMm = new[] { -10d, 0d, 0d },
                p => p.Layers[0].LeafBoundariesMm = new[] { -10d, double.NaN, 10d },
                p => p.Layers[0].SourceLeafIndices = new[] { 0, 0 },
                p => p.Layers[0].SourceLeafIndices = new[] { 0, 2 },
                p => p.Layers[0].SourceLeafIndices = new[] { -1, 1 },
                p => p.NativeLeafCount = 3,
                p => p.Layers.Add(Layer("Repeated source", new[] { 0 }, new[] { -1d, 1d }))
            };
            foreach (var mutate in mutations)
            {
                var profile = Profile(); mutate(profile);
                var error = TestAssert.Throws<NativeMlcProfileException>(() => Parse(profile));
                TestAssert.Equal("profile_invalid", error.Code); TestAssert.True(error.InnerException == null);
            }
            var duplicate = new NativeMlcProfileCatalog { SchemaVersion = 1, Profiles = new List<NativeMlcGeometryProfile> { Profile(), Profile() } };
            TestAssert.Throws<NativeMlcProfileException>(() => NativeMlcProfileCatalog.Parse(JsonConvert.SerializeObject(duplicate)));
            foreach (string json in new[] { "bad-json-private-marker", "{}", "{\"SchemaVersion\":2,\"Profiles\":[]}", "{\"SchemaVersion\":1,\"Profiles\":[],\"Typo\":true}" })
            {
                var error = TestAssert.Throws<NativeMlcProfileException>(() => NativeMlcProfileCatalog.Parse(json));
                TestAssert.False(error.ToString().Contains("private-marker")); TestAssert.True(error.InnerException == null);
            }
            var mutable = Catalog(); mutable.Profiles[0].Layers[0].SourceLeafIndices[1] = 0;
            TestAssert.Equal("profile_invalid", Map(mutable).Code);
        }
        private static void InvalidNativeValuesStayUnavailable()
        {
            foreach (float invalid in new[] { float.NaN, float.PositiveInfinity, float.NegativeInfinity, 11f })
            {
                var positions = Positions(2); positions[0, 0] = invalid;
                var result = NativeMlcGeometryMapper.Map(Catalog(), "Synthetic MLC", positions, Jaws());
                TestAssert.Equal("native_positions", result.Code); TestAssert.True(result.Geometry == null);
            }
            var closed = Positions(2); closed[0, 0] = closed[1, 0];
            TestAssert.NotNull(NativeMlcGeometryMapper.Map(Catalog(), "Synthetic MLC", closed, Jaws()).Geometry);
        }
        private static void PhysicalJawsAndFixedLimitsStayDistinct()
        {
            var fixedProfile = Profile(); fixedProfile.JawMode = "FixedLimits";
            var result = Map(Parse(fixedProfile));
            TestAssert.NotNull(result.Geometry.FixedBoundingBox); TestAssert.True(result.Geometry.Jaws == null);
            foreach (var jaws in new[] { null, new ApertureRectangle(double.NaN, -1, 1, 1), new ApertureRectangle(2, -1, 1, 1) })
            {
                TestAssert.Equal("native_jaws", NativeMlcGeometryMapper.Map(Catalog(), "Synthetic MLC", Positions(2), jaws).Code);
            }
        }
        private static void PublicReferenceProfilesHaveExpectedGeometry()
        {
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "ClearPlan.sln"))) dir = dir.Parent;
            TestAssert.NotNull(dir);
            string path = Path.Combine(dir.FullName, "ClearPlan.Script", "Distribution", "MachineGeometry", "MlcGeometryProfiles.example.json");
            var catalog = NativeMlcProfileCatalog.Parse(File.ReadAllText(path));
            TestAssert.Equal(3, catalog.Profiles.Count);
            var hd = catalog.Profiles.Single(p => p.Model == "Varian High Definition 120");
            CheckWidths(hd, -110d, 110d, 2.5d, 32, 5d, 28);
            var millennium = catalog.Profiles.Single(p => p.Model == "native-model-name-to-confirm");
            CheckWidths(millennium, -200d, 200d, 5d, 40, 10d, 20);
            TestAssert.False(catalog.Profiles.Any(p => p.Model.IndexOf("Halcyon", StringComparison.OrdinalIgnoreCase) >= 0));
        }
        private static void Sx2ReferenceKeepsDistinctNativeLayers()
        {
            var catalog = ReferenceCatalog();
            var profile = catalog.Profiles.Single(p => p.Model == "SX2");
            TestAssert.Equal(57, profile.NativeLeafCount);
            TestAssert.Equal("None", profile.JawMode);
            TestAssert.Equal(2, profile.Layers.Count);
            TestAssert.Equal("Distal physical MLC", profile.Layers[0].Label);
            TestAssert.Equal("Proximal physical MLC", profile.Layers[1].Label);
            var positions = Positions(57);
            for (int index = 0; index < 57; index++)
            {
                positions[0, index] = -100 + index;
                positions[1, index] = 20 + index;
            }
            var result = NativeMlcGeometryMapper.Map(catalog, "SX2", positions, new ApertureRectangle(-140, -140, 140, 140));
            TestAssert.Equal("available", result.Code);
            TestAssert.True(result.Geometry.Jaws == null, "SX2 has no movable physical jaws.");
            TestAssert.NotNull(result.Geometry.FixedBoundingBox, "The SX2 reference profile must carry its explicit nominal maximum field, independently of native jaw values.");
            TestAssert.Equal(-140d, result.Geometry.FixedBoundingBox.X1);
            TestAssert.Equal(140d, result.Geometry.FixedBoundingBox.X2);
            TestAssert.Equal(-140d, result.Geometry.FixedBoundingBox.Y1);
            TestAssert.Equal(140d, result.Geometry.FixedBoundingBox.Y2);
            TestAssert.Equal("available", NativeMlcGeometryMapper.Map(catalog, "SX2", positions, null).Code,
                "Jawless SX2 must not require or infer a native jaw rectangle.");
            for (int layerIndex = 0; layerIndex < 2; layerIndex++)
            {
                int count = layerIndex == 0 ? 28 : 29;
                int offset = layerIndex == 0 ? 0 : 28;
                double firstBoundary = layerIndex == 0 ? -140 : -145;
                var layer = result.Geometry.Layers[layerIndex];
                TestAssert.Equal("X", layer.LeafTravelAxis);
                TestAssert.Equal(count, layer.Bank1PositionsMm.Length);
                for (int index = 0; index < count; index++)
                {
                    TestAssert.Equal(offset + index, profile.Layers[layerIndex].SourceLeafIndices[index]);
                    TestAssert.Equal((double)positions[0, offset + index], layer.Bank1PositionsMm[index]);
                    TestAssert.Equal((double)positions[1, offset + index], layer.Bank2PositionsMm[index]);
                }
                for (int index = 0; index <= count; index++)
                    TestAssert.Equal(firstBoundary + 10 * index, layer.LeafBoundariesMm[index]);
            }
            foreach (string otherModel in new[] { "SX1", "sx2", "SX2 ", "Halcyon" })
                TestAssert.Equal("profile_missing", NativeMlcGeometryMapper.Map(catalog, otherModel, positions, Jaws()).Code);
            TestAssert.Equal("native_shape", NativeMlcGeometryMapper.Map(catalog, "SX2", Positions(56), Jaws()).Code);
        }
        private static void Sx2StaggeredOpeningsIntersectBothLayers()
        {
            // Synthetic leaves isolate the two 5-mm overlaps around the central proximal leaf.
            var positions = new float[2, 57];
            positions[0, 13] = -20; positions[1, 13] = 20; // Distal y=-10..0.
            positions[0, 14] = -10; positions[1, 14] = 30; // Distal y=0..10.
            positions[0, 42] = -15; positions[1, 42] = 15; // Proximal y=-5..5.
            var mapped = NativeMlcGeometryMapper.Map(ReferenceCatalog(), "SX2", positions,
                new ApertureRectangle(-140, -140, 140, 140));
            TestAssert.NotNull(mapped.Geometry);
            var openings = PlanAnalysisCalculator.BuildOpenings(mapped.Geometry).OrderBy(r => r.Y1).ToList();
            TestAssert.Equal(2, openings.Count);
            TestAssert.Equal(-15d, openings[0].X1); TestAssert.Equal(15d, openings[0].X2);
            TestAssert.Equal(-5d, openings[0].Y1); TestAssert.Equal(0d, openings[0].Y2);
            TestAssert.Equal(-10d, openings[1].X1); TestAssert.Equal(15d, openings[1].X2);
            TestAssert.Equal(0d, openings[1].Y1); TestAssert.Equal(5d, openings[1].Y2);
            TestAssert.Equal(275d, openings.Sum(r => (r.X2 - r.X1) * (r.Y2 - r.Y1)));
            positions[0, 42] = positions[1, 42];
            mapped = NativeMlcGeometryMapper.Map(ReferenceCatalog(), "SX2", positions,
                new ApertureRectangle(-140, -140, 140, 140));
            TestAssert.Equal(0, PlanAnalysisCalculator.BuildOpenings(mapped.Geometry).Count,
                "Closing the proximal layer must block the distal openings completely.");
        }
        private static void ConfiguredMaximumFieldIsDetachedValidatedAndClipsTheOpening()
        {
            var catalog = ReferenceCatalog();
            var positions = new float[2, 57];
            for (int i = 0; i < 57; i++) { positions[0, i] = -200; positions[1, i] = 200; }
            var mapped = NativeMlcGeometryMapper.Map(catalog, "SX2", positions, null);
            TestAssert.NotNull(mapped.Geometry);
            var opening = PlanAnalysisCalculator.BuildOpenings(mapped.Geometry);
            TestAssert.Equal(78400d, opening.Sum(r => (r.X2-r.X1)*(r.Y2-r.Y1)), "Nominal 28 x 28 cm field limits the geometric opening.");
            catalog.Profiles.Single(p => p.Model == "SX2").MaximumFieldOpeningMm.X1 = -100;
            TestAssert.Equal(-140d, mapped.Geometry.FixedBoundingBox.X1, "Detached geometry must not alias the configuration object.");
            foreach (Action<NativeMlcGeometryProfile> invalid in new Action<NativeMlcGeometryProfile>[] {
                p => p.MaximumFieldOpeningMm.X2 = p.MaximumFieldOpeningMm.X1,
                p => p.MaximumFieldOpeningMm.Y1 = double.NaN,
                p => p.MaximumFieldOpeningMm.Y2 = 3000,
                p => p.JawMode = "Physical",
                p => p.Evidence = ""
            }) {
                var bad = ReferenceCatalog(); invalid(bad.Profiles.Single(p => p.Model == "SX2"));
                TestAssert.Throws<NativeMlcProfileException>(() => NativeMlcProfileCatalog.Parse(JsonConvert.SerializeObject(bad)));
            }
        }
        private static NativeMlcProfileCatalog ReferenceCatalog()
        {
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "ClearPlan.sln"))) dir = dir.Parent;
            TestAssert.NotNull(dir);
            return NativeMlcProfileCatalog.Parse(File.ReadAllText(Path.Combine(dir.FullName,
                "ClearPlan.Script", "Distribution", "MachineGeometry", "MlcGeometryProfiles.example.json")));
        }
        private static void EsapiNativeAdapterHasReadOnlyAtomicContract()
        {
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "ClearPlan.sln"))) dir = dir.Parent;
            TestAssert.NotNull(dir);
            string path = Path.Combine(dir.FullName, "ClearPlan.Script", "Review", "EsapiNativeMlcAdapter.cs");
            TestAssert.True(File.Exists(path), "Missing native ESAPI MLC adapter.");
            string code = File.ReadAllText(path);
            TestAssert.True(code.Contains("ApplyToBeam") && code.Contains("NativeMlcGeometryMapper.Map") && code.Contains("cp.LeafPositions"));
            TestAssert.True(code.Contains("fixed_limits_changed") && code.Contains("CommitCompleteBeam"));
            TestAssert.False(code.Contains("BeginModifications(") || code.Contains("SaveModifications(") || code.Contains("ApplyParameters("));
            TestAssert.False(code.Contains("Task.Run") || code.Contains("RTPLAN"));
        }
        private static void CheckWidths(NativeMlcGeometryProfile p, double start, double stop, double inner, int innerCount, double outer, int outerCount)
        {
            TestAssert.Equal(60, p.NativeLeafCount); TestAssert.Equal(1, p.Layers.Count);
            var bounds = p.Layers[0].LeafBoundariesMm;
            TestAssert.Equal(start, bounds.First()); TestAssert.Equal(stop, bounds.Last());
            var widths = bounds.Skip(1).Select((value, index) => value - bounds[index]).ToList();
            TestAssert.Equal(innerCount, widths.Count(w => w == inner)); TestAssert.Equal(outerCount, widths.Count(w => w == outer));
        }
        private static NativeMlcGeometryResult Map(NativeMlcProfileCatalog catalog, string model = "Synthetic MLC")
        { return NativeMlcGeometryMapper.Map(catalog, model, Positions(2), Jaws()); }
        private static NativeMlcProfileCatalog Catalog() { return Parse(Profile()); }
        private static NativeMlcProfileCatalog Parse(NativeMlcGeometryProfile profile)
        { return NativeMlcProfileCatalog.Parse(JsonConvert.SerializeObject(new NativeMlcProfileCatalog { SchemaVersion = 1, Profiles = new List<NativeMlcGeometryProfile> { profile } })); }
        private static NativeMlcGeometryProfile Profile()
        { return new NativeMlcGeometryProfile { Model = "Synthetic MLC", NativeLeafCount = 2, JawMode = "Physical", Evidence = "Synthetic test only", Layers = new List<NativeMlcGeometryLayer> { Layer("Synthetic layer", new[] { 0, 1 }, new[] { -10d, 0d, 10d }) } }; }
        private static NativeMlcGeometryLayer Layer(string label, int[] indices, double[] bounds)
        { return new NativeMlcGeometryLayer { Label = label, LeafTravelAxis = "X", SourceLeafIndices = indices, LeafBoundariesMm = bounds }; }
        private static float[,] Positions(int count)
        { var positions = new float[2, count]; for (int i = 0; i < count; i++) { positions[0, i] = -10; positions[1, i] = 10; } return positions; }
        private static ApertureRectangle Jaws() { return new ApertureRectangle(-5, -20, 5, 20); }
    }
}
