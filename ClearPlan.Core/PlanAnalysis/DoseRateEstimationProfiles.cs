using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace ClearPlan.Core.PlanAnalysis
{
    public sealed class DoseRateEstimationProfileCatalog
    {
        public int SchemaVersion { get; set; }
        public List<DoseRateEstimationProfile> Profiles { get; set; }

        public static DoseRateEstimationProfileCatalog Parse(string json)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(json) || json.Length > 65536) throw new FormatException();
                var catalog = JsonConvert.DeserializeObject<DoseRateEstimationProfileCatalog>(json, new JsonSerializerSettings
                {
                    TypeNameHandling = TypeNameHandling.None,
                    MissingMemberHandling = MissingMemberHandling.Error,
                    MaxDepth = 16
                });
                if (catalog == null || catalog.Validate().Count != 0) throw new FormatException();
                return catalog;
            }
            catch (Exception error) when (error is JsonException || error is FormatException || error is OverflowException)
            {
                // Configuration can contain private paths/values; do not retain parser messages or inner exceptions.
                throw new FormatException("Dose-rate estimation profile JSON is malformed, invalid or unsupported.");
            }
        }

        public IList<string> Validate()
        {
            var messages = new List<string>();
            if (SchemaVersion != 1) messages.Add("Dose-rate profiles require SchemaVersion 1.");
            if (Profiles == null || Profiles.Count > 128)
            {
                messages.Add("Dose-rate profiles require a list of at most 128 profiles.");
                return messages;
            }
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var profile in Profiles)
            {
                string reason = ValidateProfile(profile);
                if (reason != null) messages.Add(reason);
                if (profile != null && !string.IsNullOrEmpty(profile.Id) && !ids.Add(profile.Id))
                    messages.Add("Dose-rate profile identifiers must be unique, ignoring case.");
            }
            return messages;
        }

        public DoseRateEstimationProfile Resolve(ReviewBeamAnalysis beam, out string reason)
        {
            if (Validate().Count != 0)
            { reason = "Dose-rate estimation profile catalog is invalid."; return null; }
            if (beam == null)
            { reason = "Dose-rate estimation requires detached beam metadata."; return null; }
            var matches = Profiles.Where(p => Matches(p, beam)).ToList();
            var exactIds = matches.Where(p => HasValues(p.MachineIds)).ToList();
            if (exactIds.Count != 0) matches = exactIds;
            if (matches.Count != 1)
            {
                reason = matches.Count == 0
                    ? "No dose-rate profile matches the exact machine metadata; no machine speed is inferred."
                    : "Dose-rate estimation profile match is ambiguous; no profile is selected.";
                return null;
            }
            reason = "Unique exact machine metadata match; speed is a configured model assumption.";
            return matches[0];
        }

        internal static string ValidateProfile(DoseRateEstimationProfile profile)
        {
            if (profile == null || !ValidText(profile.Id, 128) || !ValidText(profile.Assumption, 4096))
                return "Dose-rate profiles require a nonblank identifier and explicit assumption.";
            if (!Finite(profile.MaxGantrySpeedDegreesPerSecond) || profile.MaxGantrySpeedDegreesPerSecond <= 0)
                return "Dose-rate profiles require a finite positive maximum gantry speed in degrees/second.";
            if (!ValidList(profile.MachineIds) || !ValidList(profile.MachineModelNames) ||
                !ValidList(profile.MachineModels) || !ValidList(profile.MlcModels))
                return "Dose-rate profile matching lists must contain distinct, nonblank exact values without edge whitespace.";
            if (!HasValues(profile.MachineIds) && !HasValues(profile.MachineModelNames) && !HasValues(profile.MachineModels))
                return "Dose-rate profiles require exact machine IDs or machine models; MLC alone is insufficient.";
            return null;
        }

        internal static bool Matches(DoseRateEstimationProfile profile, ReviewBeamAnalysis beam)
        {
            return MatchList(profile.MachineIds, beam.MachineId) && MatchList(profile.MachineModelNames, beam.MachineModelName)
                && MatchList(profile.MachineModels, beam.MachineModel) && MatchList(profile.MlcModels, beam.MlcModel);
        }
        internal static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static bool HasValues(List<string> values) { return values != null && values.Count != 0; }
        private static bool MatchList(List<string> values, string actual)
        { return !HasValues(values) || values.Any(value => string.Equals(value, actual, StringComparison.OrdinalIgnoreCase)); }
        private static bool ValidText(string value, int maximum)
        { return !string.IsNullOrWhiteSpace(value) && value.Length <= maximum && value == value.Trim(); }
        private static bool ValidList(List<string> values)
        {
            if (values == null) return true;
            return values.Count <= 256 && values.All(value => ValidText(value, 256)) &&
                values.Distinct(StringComparer.OrdinalIgnoreCase).Count() == values.Count;
        }
    }

    public sealed class DoseRateEstimationProfile
    {
        public string Id { get; set; }
        public List<string> MachineIds { get; set; }
        public List<string> MachineModelNames { get; set; }
        public List<string> MachineModels { get; set; }
        public List<string> MlcModels { get; set; }
        public double MaxGantrySpeedDegreesPerSecond { get; set; }
        public string Assumption { get; set; }
    }
}
