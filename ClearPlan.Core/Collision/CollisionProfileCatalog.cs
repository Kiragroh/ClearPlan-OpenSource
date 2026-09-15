using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace ClearPlan.Core.Collision
{
    public sealed class CollisionProfileCatalog
    {
        public int SchemaVersion { get; set; }
        public List<CollisionProfile> Profiles { get; set; }
        public static CollisionProfileCatalog Parse(string json)
        {
            if (string.IsNullOrWhiteSpace(json) || json.Length > 1024 * 1024) Invalid();
            try
            {
                RejectDuplicateProperties(json);
                var catalog = JsonConvert.DeserializeObject<CollisionProfileCatalog>(json, new JsonSerializerSettings {
                    TypeNameHandling = TypeNameHandling.None, MissingMemberHandling = MissingMemberHandling.Error, MaxDepth = 16 });
                if (catalog == null || catalog.SchemaVersion != 1 || catalog.Profiles == null || catalog.Profiles.Count > 128) Invalid();
                var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var profile in catalog.Profiles)
                {
                    Validate(profile, profile != null && profile.SyntheticOnly);
                    if (!ids.Add(profile.MachineId)) Invalid();
                }
                return catalog;
            }
            catch (JsonException) { throw new ArgumentException("Collision profile JSON is malformed or unsupported."); }
        }

        public CollisionProfile Select(string machineId)
        {
            if (!Text(machineId, 256) || machineId != machineId.Trim()) return null;
            if (SchemaVersion != 1 || Profiles == null || Profiles.Count > 128) Invalid();
            var matches = Profiles.Where(p => p != null && string.Equals(p.MachineId, machineId, StringComparison.OrdinalIgnoreCase)).ToList();
            if (matches.Count == 0) return null;
            if (matches.Count != 1) Invalid();
            Validate(matches[0], matches[0].SyntheticOnly);
            return matches[0];
        }

        public static void Validate(CollisionProfile profile, bool synthetic)
        {
            if (profile == null || !Text(profile.MachineId, 256) || profile.MachineId != profile.MachineId.Trim() ||
                !Text(profile.Revision, 256) || !Text(profile.CommissionedBy, 512) || !Text(profile.Evidence, 4096) ||
                (profile.SyntheticOnly && !synthetic) || (!profile.Commissioned && !(profile.SyntheticOnly && synthetic)) ||
                !Range(profile.SafetyMarginMm, 0, 1000) ||
                !Range(profile.HeadCenterFromIsoMm, 0, 10000) || !Range(profile.HeadRadiusMm, 0, 10000) ||
                !Range(profile.BoreRadiusMm, 0, 10000) || !Range(profile.BoreHalfLengthMm, 0, 10000)) Invalid();
            if (profile.Kind == "CArmSphere")
            {
                if (profile.HeadCenterFromIsoMm <= 0 || profile.HeadRadiusMm <= 0) Invalid();
            }
            else if (profile.Kind == "RingBore")
            {
                if (profile.BoreRadiusMm <= 0 || profile.BoreHalfLengthMm <= 0) Invalid();
            }
            else Invalid();
        }

        private static void RejectDuplicateProperties(string json)
        {
            var objects = new Stack<HashSet<string>>();
            using (var reader = new JsonTextReader(new StringReader(json)) { MaxDepth = 16 })
            {
                while (reader.Read())
                {
                    if (reader.TokenType == JsonToken.StartObject) objects.Push(new HashSet<string>(StringComparer.OrdinalIgnoreCase));
                    else if (reader.TokenType == JsonToken.EndObject) objects.Pop();
                    else if (reader.TokenType == JsonToken.PropertyName && !objects.Peek().Add((string)reader.Value)) Invalid();
                }
            }
        }

        private static bool Text(string value, int maximum) { return !string.IsNullOrWhiteSpace(value) && value.Length <= maximum; }
        private static bool Range(double value, double low, double high) { return !double.IsNaN(value) && value >= low && value <= high; }
        private static void Invalid()
        { throw new ArgumentException("Collision profile requires an exact machine ID, bounded envelope, revision and commissioning evidence; synthetic profiles cannot evaluate patient scenes."); }
    }
}
