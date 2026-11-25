using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace ClearPlan.Core.Collision
{
    public sealed class SourceCollisionModelCatalog
    {
        public int SchemaVersion { get; set; }
        public string Units { get; set; }
        public List<SourceCollisionModel> Models { get; set; }
        public List<SourceCollisionMachineBinding> MachineBindings { get; set; }
        public static SourceCollisionModelCatalog Parse(string json)
        {
            if (string.IsNullOrWhiteSpace(json) || json.Length > 1024 * 1024) Invalid();
            try
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
                var catalog = JsonConvert.DeserializeObject<SourceCollisionModelCatalog>(json, new JsonSerializerSettings {
                    TypeNameHandling = TypeNameHandling.None, MissingMemberHandling = MissingMemberHandling.Error,
                    DateParseHandling = DateParseHandling.None, MaxDepth = 16 });
                if (catalog == null) Invalid();
                catalog.Validate();
                return catalog;
            }
            catch (JsonException) { throw new ArgumentException("Source collision model JSON is malformed or unsupported."); }
        }

        public SourceCollisionModel Select(string machineId)
        {
            Validate();
            if (!Text(machineId, 256)) return null;
            var binding = MachineBindings.SingleOrDefault(b => string.Equals(b.MachineId, machineId, StringComparison.OrdinalIgnoreCase));
            return binding == null ? null : Models.Single(m => string.Equals(m.ModelId, binding.ModelId, StringComparison.OrdinalIgnoreCase));
        }

        internal void Validate()
        {
            if (SchemaVersion != 1 || Units != "mm" || Models == null || Models.Count > 16 ||
                MachineBindings == null || MachineBindings.Count > 128) Invalid();
            var modelIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var model in Models)
            {
                if (model == null || !Text(model.ModelId, 256) || !modelIds.Add(model.ModelId) || !Text(model.SourcePath, 4096) ||
                    !Text(model.SourceSha256, 64) || !Regex.IsMatch(model.SourceSha256, "\\A[0-9a-fA-F]{64}\\z") ||
                    !Date(model.SourceSnapshotDate) || !Text(model.Evidence, 4096)) Invalid();
                if (model.EvidenceLevel == "source-recorded-measurement")
                { if (!Date(model.MeasurementDate)) Invalid(); }
                else if (model.EvidenceLevel != "source-constants" || !string.IsNullOrEmpty(model.MeasurementDate)) Invalid();

                if (model.Kind == "TrueBeamHeadSource")
                {
                    if (model.HeadObstacles == null || model.HeadObstacles.Count == 0 || model.HeadObstacles.Count > 16 ||
                        !Range(model.ReachRadiusMm, 1, 10000) || !Range(model.WarningMarginMm, 0, 1000) ||
                        !Range(model.SourceAngularMarginDegrees, 0, 90) || model.BoreWarningRadiusMm != 0 ||
                        model.BoreLimitRadiusMm != 0 || model.LeadingWarningFromIsoMm != 0 || model.LeadingLimitFromIsoMm != 0) Invalid();
                    foreach (var head in model.HeadObstacles)
                        if (head == null || !Range(head.FrontFaceFromIsoMm, 1, model.ReachRadiusMm) ||
                            !Range(head.RadiusMm, 1, model.ReachRadiusMm) || model.WarningMarginMm >= head.FrontFaceFromIsoMm) Invalid();
                }
                else if (model.Kind == "HalcyonBoreSource")
                {
                    if ((model.HeadObstacles != null && model.HeadObstacles.Count != 0) || model.ReachRadiusMm != 0 ||
                        model.WarningMarginMm != 0 || model.SourceAngularMarginDegrees != 0 ||
                        !Range(model.BoreWarningRadiusMm, 1, 10000) || !Range(model.BoreLimitRadiusMm, 1, 10000) ||
                        model.BoreWarningRadiusMm >= model.BoreLimitRadiusMm ||
                        !Range(model.LeadingWarningFromIsoMm, 1, 10000) || !Range(model.LeadingLimitFromIsoMm, 1, 10000) ||
                        model.LeadingWarningFromIsoMm >= model.LeadingLimitFromIsoMm) Invalid();
                }
                else Invalid();
            }
            var machineIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var binding in MachineBindings)
                if (binding == null || !Text(binding.MachineId, 256) || !machineIds.Add(binding.MachineId) ||
                    !Text(binding.ModelId, 256) || !modelIds.Contains(binding.ModelId)) Invalid();
        }

        internal static bool Text(string value, int maximum)
        { return !string.IsNullOrWhiteSpace(value) && value.Length <= maximum && value == value.Trim() && !value.Any(char.IsControl); }
        private static bool Date(string value)
        {
            DateTime parsed;
            return DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed);
        }
        private static bool Range(double value, double low, double high)
        { return !double.IsNaN(value) && value >= low && value <= high; }
        private static void Invalid()
        { throw new ArgumentException("Source collision models require bounded millimetre geometry, source provenance and exact, unambiguous machine bindings; they do not grant commissioning."); }
    }

    public sealed class SourceCollisionMachineBinding
    {
        public string MachineId { get; set; }
        public string ModelId { get; set; }
    }

    public sealed class SourceCollisionModel
    {
        public string ModelId { get; set; }
        public string Kind { get; set; }
        public string SourcePath { get; set; }
        public string SourceSha256 { get; set; }
        public string SourceSnapshotDate { get; set; }
        public string EvidenceLevel { get; set; }
        public string MeasurementDate { get; set; }
        public string Evidence { get; set; }
        public List<SourceHeadObstacle> HeadObstacles { get; set; }
        public double ReachRadiusMm { get; set; }
        public double WarningMarginMm { get; set; }
        public double SourceAngularMarginDegrees { get; set; }
        public double BoreWarningRadiusMm { get; set; }
        public double BoreLimitRadiusMm { get; set; }
        public double LeadingWarningFromIsoMm { get; set; }
        public double LeadingLimitFromIsoMm { get; set; }
    }

    public sealed class SourceHeadObstacle
    {
        public double FrontFaceFromIsoMm { get; set; }
        public double RadiusMm { get; set; }
    }
}
