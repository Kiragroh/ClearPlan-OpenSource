using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace ClearPlan.Core.PlanAnalysis
{
    public sealed class NativeMlcProfileException : Exception
    {
        public NativeMlcProfileException(string code, string message) : base(message) { Code = code; }
        public string Code { get; private set; }
    }
    public sealed class NativeMlcProfileCatalog
    {
        public int SchemaVersion { get; set; }
        public List<NativeMlcGeometryProfile> Profiles { get; set; }
        public static NativeMlcProfileCatalog Parse(string json)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(json) || json.Length > 1024 * 1024) Invalid();
                var catalog = JsonConvert.DeserializeObject<NativeMlcProfileCatalog>(json, new JsonSerializerSettings
                { TypeNameHandling = TypeNameHandling.None, MissingMemberHandling = MissingMemberHandling.Error, MaxDepth = 32 });
                if (catalog == null || catalog.SchemaVersion != 1 || catalog.Profiles == null ||
                    catalog.Profiles.Count == 0 || catalog.Profiles.Count > 128) Invalid();
                var models = new HashSet<string>(StringComparer.Ordinal);
                foreach (var profile in catalog.Profiles)
                {
                    ValidateProfile(profile);
                    if (!models.Add(profile.Model)) Invalid();
                }
                return catalog;
            }
            catch (NativeMlcProfileException) { throw; }
            catch (Exception)
            {
                // Parser messages may contain configuration paths or values. Keep errors detached.
                throw new NativeMlcProfileException("profile_invalid", "Native MLC profile JSON is malformed or unsupported.");
            }
        }

        internal static void ValidateProfile(NativeMlcGeometryProfile profile)
        {
            if (profile == null || string.IsNullOrWhiteSpace(profile.Model) || profile.Model.Length > 256 ||
                profile.Model != profile.Model.Trim() || profile.NativeLeafCount < 1 || profile.NativeLeafCount > 2000 ||
                !(profile.JawMode == "Physical" || profile.JawMode == "FixedLimits" || profile.JawMode == "None") ||
                profile.Layers == null || profile.Layers.Count == 0 || profile.Layers.Count > 8) Invalid();
            var usedIndices = new HashSet<int>();
            foreach (var layer in profile.Layers)
            {
                if (layer == null || string.IsNullOrWhiteSpace(layer.Label) ||
                    !(layer.LeafTravelAxis == "X" || layer.LeafTravelAxis == "Y") ||
                    layer.SourceLeafIndices == null || layer.SourceLeafIndices.Length == 0 ||
                    layer.SourceLeafIndices.Length > profile.NativeLeafCount || layer.LeafBoundariesMm == null ||
                    layer.LeafBoundariesMm.Length != layer.SourceLeafIndices.Length + 1 ||
                    layer.LeafBoundariesMm.Any(value => !Finite(value))) Invalid();
                for (int i = 1; i < layer.LeafBoundariesMm.Length; i++)
                    if (layer.LeafBoundariesMm[i] <= layer.LeafBoundariesMm[i - 1]) Invalid();
                foreach (int index in layer.SourceLeafIndices)
                    if (index < 0 || index >= profile.NativeLeafCount || !usedIndices.Add(index)) Invalid();
            }
            // Every native leaf must belong to exactly one explicitly configured physical layer.
            if (usedIndices.Count != profile.NativeLeafCount) Invalid();
        }
        internal static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static void Invalid()
        { throw new NativeMlcProfileException("profile_invalid", "Native MLC profile needs an exact model, valid layer boundaries, and complete non-overlapping native leaf indices."); }
    }
    public sealed class NativeMlcGeometryProfile
    {
        public string Model { get; set; }
        public int NativeLeafCount { get; set; }
        public string JawMode { get; set; }
        public string Evidence { get; set; }
        public List<NativeMlcGeometryLayer> Layers { get; set; }
    }
    public sealed class NativeMlcGeometryLayer
    {
        public string Label { get; set; }
        public string LeafTravelAxis { get; set; }
        public int[] SourceLeafIndices { get; set; }
        public double[] LeafBoundariesMm { get; set; }
    }
    public sealed class NativeMlcGeometryResult
    {
        [JsonIgnore] public ApertureGeometry Geometry { get; internal set; }
        public string Code { get; internal set; }
        public string Reason { get; internal set; }
        public int NativeBankCount { get; internal set; }
        public int NativeLeafCount { get; internal set; }
    }
    public static class NativeMlcGeometryMapper
    {
        /// <summary>
        /// Maps detached native bank data through an explicit machine profile, never by model-name
        /// similarity or guessed dual-layer layout. A beam-level caller must additionally verify
        /// that FixedLimits do not move between control points.
        /// </summary>
        public static NativeMlcGeometryResult Map(NativeMlcProfileCatalog catalog, string exactMlcModel,
            float[,] nativePositions, ApertureRectangle nativeJawPositions)
        {
            var result = new NativeMlcGeometryResult
            {
                NativeBankCount = nativePositions == null ? 0 : nativePositions.GetLength(0),
                NativeLeafCount = nativePositions == null ? 0 : nativePositions.GetLength(1)
            };
            string shape = result.NativeBankCount + "x" + result.NativeLeafCount;
            if (catalog == null || string.IsNullOrEmpty(exactMlcModel))
                return Unavailable(result, "profile_missing", "No exact native MLC geometry profile is selected; observed native array " + shape + ".");
            if (catalog.SchemaVersion != 1 || catalog.Profiles == null || catalog.Profiles.Count > 128)
                return Unavailable(result, "profile_invalid", "The native MLC geometry profile catalog is invalid.");
            var matches = catalog.Profiles.Where(p => p != null && string.Equals(p.Model, exactMlcModel, StringComparison.Ordinal)).ToList();
            if (matches.Count == 0)
                return Unavailable(result, "profile_missing", "No profile matches the exact native MLC model; observed native array " + shape + ". No physical layer layout is inferred.");
            if (matches.Count != 1)
                return Unavailable(result, "profile_invalid", "The exact native MLC model has more than one profile.");
            var profile = matches[0];
            try { NativeMlcProfileCatalog.ValidateProfile(profile); }
            catch (NativeMlcProfileException)
            { return Unavailable(result, "profile_invalid", "The native MLC geometry profile has invalid boundaries or index coverage."); }
            if (result.NativeBankCount != 2 || result.NativeLeafCount != profile.NativeLeafCount)
                return Unavailable(result, "native_shape", "The observed native array " + shape + " does not match the explicit two-bank profile.");
            for (int i = 0; i < result.NativeLeafCount; i++)
            {
                if (!NativeMlcProfileCatalog.Finite(nativePositions[0, i]) || !NativeMlcProfileCatalog.Finite(nativePositions[1, i]) ||
                    nativePositions[0, i] > nativePositions[1, i])
                    return Unavailable(result, "native_positions", "Native leaf-bank positions are non-finite or reversed; no partial geometry is used.");
            }
            var geometry = new ApertureGeometry();
            if (profile.JawMode != "None")
            {
                if (!ValidRectangle(nativeJawPositions))
                    return Unavailable(result, "native_jaws", "The profile requires a finite, ordered native jaw or fixed-limit rectangle.");
                var rectangle = new ApertureRectangle(nativeJawPositions.X1, nativeJawPositions.Y1, nativeJawPositions.X2, nativeJawPositions.Y2);
                if (profile.JawMode == "Physical") geometry.Jaws = rectangle;
                else geometry.FixedBoundingBox = rectangle;
            }
            foreach (var layer in profile.Layers)
            {
                geometry.Layers.Add(new ApertureLayer
                {
                    Label = layer.Label, LeafTravelAxis = layer.LeafTravelAxis,
                    LeafBoundariesMm = (double[])layer.LeafBoundariesMm.Clone(),
                    Bank1PositionsMm = layer.SourceLeafIndices.Select(index => (double)nativePositions[0, index]).ToArray(),
                    Bank2PositionsMm = layer.SourceLeafIndices.Select(index => (double)nativePositions[1, index]).ToArray()
                });
            }
            result.Geometry = geometry; result.Code = "available";
            result.Reason = "Native bank positions mapped using an explicit machine geometry profile; research geometry, not clinical commissioning.";
            return result;
        }
        private static bool ValidRectangle(ApertureRectangle value)
        {
            return value != null && NativeMlcProfileCatalog.Finite(value.X1) && NativeMlcProfileCatalog.Finite(value.Y1) &&
                NativeMlcProfileCatalog.Finite(value.X2) && NativeMlcProfileCatalog.Finite(value.Y2) && value.X1 <= value.X2 && value.Y1 <= value.Y2;
        }
        private static NativeMlcGeometryResult Unavailable(NativeMlcGeometryResult result, string code, string reason)
        { result.Code = code; result.Reason = reason; return result; }
    }
}
