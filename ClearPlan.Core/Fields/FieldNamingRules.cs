using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClearPlan.Core.Review;
using Newtonsoft.Json;

namespace ClearPlan.Core.Fields
{
    /// <summary>Editable name tokens/order; treatment ordering and the established field-ID algorithm stay unchanged.</summary>
    public sealed class FieldNamingRules
    {
        public int SchemaVersion { get; set; }
        public bool Enabled { get; set; }
        public List<string> ArcOrder { get; set; }
        public List<string> StaticOrder { get; set; }
        public string PartSeparator { get; set; }
        public string AngleSeparator { get; set; }
        public string TablePrefix { get; set; }
        public string ClockwiseToken { get; set; }
        public string CounterClockwiseToken { get; set; }
        public string StaticDuplicateToken { get; set; }

        // Explicit deterministic examples/simulator only. Clinical callers must load their external file.
        public static FieldNamingRules CreateDefault()
        {
            return new FieldNamingRules
            {
                SchemaVersion = 1, Enabled = true,
                ArcOrder = new List<string> { "angles", "table", "direction" },
                StaticOrder = new List<string> { "angles", "table" },
                PartSeparator = " ", AngleSeparator = "-", TablePrefix = "T",
                ClockwiseToken = "UZ", CounterClockwiseToken = "GUZ", StaticDuplicateToken = "UZ"
            };
        }

        public bool IsValid()
        {
            return SchemaVersion == 1 && ExactOrder(ArcOrder, "angles", "table", "direction") &&
                ExactOrder(StaticOrder, "angles", "table") &&
                Regex.IsMatch(PartSeparator ?? "", @"^[ _-]{1,3}$") &&
                Regex.IsMatch(AngleSeparator ?? "", @"^[A-Za-z0-9_:-]{1,8}$") &&
                ValidToken(TablePrefix) && ValidToken(ClockwiseToken) && ValidToken(CounterClockwiseToken) &&
                ValidToken(StaticDuplicateToken) &&
                !string.Equals(ClockwiseToken, CounterClockwiseToken, StringComparison.OrdinalIgnoreCase);
        }

        private static bool ExactOrder(IList<string> order, params string[] expected)
        {
            return order != null && order.Count == expected.Length && order.Distinct(StringComparer.Ordinal).Count() == expected.Length &&
                expected.All(token => order.Contains(token));
        }

        private static bool ValidToken(string value)
        {
            return Regex.IsMatch(value ?? "", @"^[\p{L}\p{Nd}_-]{1,16}$");
        }
    }

    /// <summary>Small read-only JSON configuration. Missing/deleted/invalid configurations never restore default suggestions.</summary>
    public sealed class FieldNamingRuleConfiguration
    {
        public FieldNamingRules Rules { get; private set; }
        public string Status { get; private set; }
        public string Message { get; private set; }

        public static FieldNamingRuleConfiguration Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return Result(ReviewStatusCodes.NotConfigured, "Default field-naming path is empty; preview not evaluated.");
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    if (stream.Length > 32768) return Invalid("Default field-naming JSON exceeds 32 KiB.");
                    using (var reader = new StreamReader(stream, new UTF8Encoding(false, true), true))
                    {
                        var buffer = new char[32769];
                        int count = reader.ReadBlock(buffer, 0, buffer.Length);
                        if (count > 32768) return Invalid("Default field-naming JSON exceeds the read limit.");
                        return Parse(new string(buffer, 0, count));
                    }
                }
            }
            catch (FileNotFoundException) { return Result(ReviewStatusCodes.NotConfigured, "Default field-naming file missing/deleted; no replacement created; preview not evaluated."); }
            catch (DirectoryNotFoundException) { return Result(ReviewStatusCodes.NotConfigured, "Default field-naming file missing/deleted; no replacement created; preview not evaluated."); }
            catch (Exception) { return Invalid("Default field-naming file could not be read."); }
        }

        public static FieldNamingRuleConfiguration Parse(string json)
        {
            try
            {
                if (json == null || json.Length > 32768) return Invalid("Default field-naming JSON is missing or too large.");
                var rules = JsonConvert.DeserializeObject<FieldNamingRules>(json, new JsonSerializerSettings
                { MaxDepth = 6, TypeNameHandling = TypeNameHandling.None, MissingMemberHandling = MissingMemberHandling.Error });
                if (rules == null || !rules.IsValid()) return Invalid("Default field-naming schema/tokens/order are invalid or ambiguous.");
                var result = Result(rules.Enabled ? ReviewStatusCodes.Available : ReviewStatusCodes.NotConfigured,
                    rules.Enabled ? "Default field-naming rules loaded from editable JSON; read-only suggestions, never TPS renaming. Field IDs retain the established prefix/sequence policy."
                        : "Default field-naming rules disabled; preview not evaluated.");
                result.Rules = rules;
                return result;
            }
            catch (Exception) { return Invalid("Default field-naming JSON is invalid or unsupported."); }
        }

        private static FieldNamingRuleConfiguration Invalid(string reason)
        {
            return Result(ReviewStatusCodes.Unavailable, reason + " No fallback suggestions; preview not evaluated.");
        }

        private static FieldNamingRuleConfiguration Result(string status, string message)
        {
            return new FieldNamingRuleConfiguration { Status = status, Message = message };
        }
    }
}
