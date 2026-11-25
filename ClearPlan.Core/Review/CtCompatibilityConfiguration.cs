using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ClearPlan.Core.Review
{
    /// <summary>Exact local approvals for detached manufacturer/model/serial/HU-calibration tuples.
    /// Only outer whitespace and case are normalized. No wildcards, aliases or site defaults.</summary>
    public sealed class CtCompatibilityConfiguration
    {
        private const int MaximumBytes = 131072;
        private readonly List<Combination> combinations = new List<Combination>();
        private static readonly Regex LegacyTuple = new Regex(
            @"\AImagingDevice-SerialNo \('(?<serial>[^'\r\n]{0,256})'\), -Model \('(?<model>[^'\r\n]{0,256})'\), -Manufacturer \('(?<manufacturer>[^'\r\n]{0,256})'\) and HU-Table \('(?<calibration>[^'\r\n]{0,256})'\) match\?\z",
            RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));

        public string Status { get; private set; }
        public string Message { get; private set; }

        public static CtCompatibilityConfiguration Parse(string json)
        {
            if (json == null) return Missing();
            if (Encoding.UTF8.GetByteCount(json) > MaximumBytes) return Invalid();
            try
            {
                var file = JObject.Parse(json, new JsonLoadSettings { DuplicatePropertyNameHandling = DuplicatePropertyNameHandling.Error });
                if (!HasOnly(file, "schemaVersion", "combinations") || file["schemaVersion"].Type != JTokenType.Integer ||
                    (int)file["schemaVersion"] != 1 || file["combinations"].Type != JTokenType.Array) return Invalid();
                var entries = (JArray)file["combinations"];
                if (entries.Count > 1000) return Invalid();
                var result = new CtCompatibilityConfiguration { Status = ReviewStatusCodes.Available,
                    Message = "CT compatibility configuration loaded; only explicit enabled exact tuples can pass." };
                var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (JToken item in entries)
                {
                    var entry = item as JObject;
                    if (!HasOnly(entry, "id", "enabled", "manufacturer", "model", "serialNumber", "calibration") ||
                        entry["enabled"].Type != JTokenType.Boolean) return Invalid();
                    foreach (string field in new[] { "id", "manufacturer", "model", "serialNumber", "calibration" })
                        if (entry[field].Type != JTokenType.String || !ValidValue((string)entry[field])) return Invalid();
                    var combination = new Combination { Id = ((string)entry["id"]).Trim(), Enabled = (bool)entry["enabled"],
                        Manufacturer = ((string)entry["manufacturer"]).Trim(), Model = ((string)entry["model"]).Trim(),
                        SerialNumber = ((string)entry["serialNumber"]).Trim(), Calibration = ((string)entry["calibration"]).Trim() };
                    if (!ids.Add(combination.Id) || result.combinations.Any(c => c.Matches(combination.Manufacturer, combination.Model,
                        combination.SerialNumber, combination.Calibration))) return Invalid();
                    result.combinations.Add(combination);
                }
                return result;
            }
            catch (Exception ex) when (ex is JsonException || ex is ArgumentException || ex is OverflowException) { return Invalid(); }
        }

        public static CtCompatibilityConfiguration Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return Missing();
            // Optional network configuration must not block the ESAPI owner thread indefinitely.
            // This worker reads configuration bytes only, never an ESAPI object or clinical snapshot.
            var read = Task.Run(() =>
            {
                try
                {
                    using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                    {
                        if (stream.Length > MaximumBytes) return Invalid();
                        var bytes = new byte[MaximumBytes + 1];
                        int count = 0, chunk;
                        while (count < bytes.Length && (chunk = stream.Read(bytes, count, bytes.Length - count)) > 0) count += chunk;
                        if (count > MaximumBytes) return Invalid();
                        using (var reader = new StreamReader(new MemoryStream(bytes, 0, count), new UTF8Encoding(false, true), true))
                            return Parse(reader.ReadToEnd());
                    }
                }
                catch (FileNotFoundException) { return Missing(); }
                catch (DirectoryNotFoundException) { return Missing(); }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException || ex is System.Security.SecurityException)
                { return Invalid(); }
            });
            try { if (read.Wait(TimeSpan.FromSeconds(2))) return read.Result; }
            catch (AggregateException) { return Invalid(); }
            read.ContinueWith(t => { var observed = t.Exception; }, TaskContinuationOptions.OnlyOnFaulted);
            return Invalid();
        }

        public ReviewCheckRow Evaluate(string manufacturer, string model, string serialNumber, string calibration)
        {
            var row = new ReviewCheckRow { CheckCode = "ct-compatibility", Category = "CT compatibility", Unit = ReviewUnitCodes.Text,
                Status = ReviewStatusCodes.NotEvaluated, Severity = ReviewSeverityCodes.Warning,
                ObservedValue = "Manufacturer: " + Display(manufacturer) + "; model: " + Display(model) +
                    "; serial: " + Display(serialNumber) + "; HU calibration: " + Display(calibration),
                ExpectedValue = "An enabled exact manufacturer/model/serial/HU-calibration tuple in the local configuration." };
            if (Status != ReviewStatusCodes.Available) { row.Message = Message; return row; }
            if (!new[] { manufacturer, model, serialNumber, calibration }.All(ValidValue))
            { row.Message = "CT compatibility not evaluated: required manufacturer, model, serial or HU calibration is missing or invalid."; return row; }
            if (!combinations.Any(c => c.Enabled))
            { row.Message = "CT compatibility not evaluated: no enabled approved combinations are configured."; return row; }
            var approved = combinations.FirstOrDefault(c => c.Enabled && c.Matches(manufacturer, model, serialNumber, calibration));
            if (approved == null)
            { row.Status = ReviewStatusCodes.Fail; row.Severity = ReviewSeverityCodes.Error;
                row.Message = "CT compatibility mismatch: the complete observed tuple is not approved in the local configuration."; }
            else
            { row.Status = ReviewStatusCodes.Pass; row.Severity = ReviewSeverityCodes.None;
                row.Message = "CT compatibility verified: the complete observed tuple matches an enabled local approval.";
                row.ExpectedValue = "Approved combination: " + approved.Id; }
            return row;
        }

        /// <summary>Narrow adapter for the private legacy fallback check. The code AND complete
        /// anchored four-field grammar are required. Never infer approval from a loose message match.</summary>
        public ReviewCheckRow EvaluateLegacy(string rawCheckCode, string originalMessage)
        {
            if (rawCheckCode != "19") return null;
            Match match = LegacyTuple.Match(originalMessage ?? string.Empty);
            ReviewCheckRow row = match.Success
                ? Evaluate(match.Groups["manufacturer"].Value, match.Groups["model"].Value,
                    match.Groups["serial"].Value, match.Groups["calibration"].Value)
                : Evaluate(null, null, null, null);
            if (!match.Success) row.Message = "CT compatibility not evaluated: the complete legacy CT tuple could not be read. No approval inferred.";
            // Keep original finding text as provenance without changing the upstream ErrorGrid.
            row.Message += " Legacy source: " + ClinicalReviewValueMapper.SanitizeClinicalLabel(originalMessage, "unavailable");
            return row;
        }

        private static bool HasOnly(JObject value, params string[] names)
        { return value != null && value.Properties().Count() == names.Length && names.All(name => value.Property(name) != null); }
        private static bool ValidValue(string value)
        { return !string.IsNullOrWhiteSpace(value) && value.Length <= 256 && !value.Any(char.IsControl) && value.IndexOfAny(new[] { '*', '?' }) < 0; }
        private static string Display(string value)
        { return ValidValue(value) ? ClinicalReviewValueMapper.SanitizeClinicalLabel(value.Trim(), "unavailable") : "unavailable"; }
        private static CtCompatibilityConfiguration Missing()
        { return new CtCompatibilityConfiguration { Status = ReviewStatusCodes.NotConfigured,
            Message = "CT compatibility not evaluated: no readable local approval configuration. No default approval is assumed." }; }
        private static CtCompatibilityConfiguration Invalid()
        { return new CtCompatibilityConfiguration { Status = ReviewStatusCodes.Unavailable,
            Message = "CT compatibility not evaluated: configuration is invalid, unreadable, oversized or timed out. Review local settings." }; }

        private sealed class Combination
        {
            public string Id, Manufacturer, Model, SerialNumber, Calibration;
            public bool Enabled;
            public bool Matches(string manufacturer, string model, string serial, string calibration)
            { return Same(Manufacturer, manufacturer) && Same(Model, model) && Same(SerialNumber, serial) && Same(Calibration, calibration); }
            private static bool Same(string left, string right)
            { return string.Equals(left, (right ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase); }
        }
    }
}
