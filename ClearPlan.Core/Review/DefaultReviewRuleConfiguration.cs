using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    public sealed class DefaultTargetReviewRule
    {
        public string Id { get; set; }
        public bool Enabled { get; set; }
        public List<string> TargetTypes { get; set; }
        public string Metric { get; set; }
        public string Comparator { get; set; }
        public string Reference { get; set; }
        public double PrescriptionPercent { get; set; }

        public bool TryGetVolumePercent(out double value)
        {
            value = 0;
            Match match = Regex.Match(Metric ?? "", @"^D(?<v>\d+(?:\.\d+)?)%$", RegexOptions.CultureInvariant);
            return match.Success && double.TryParse(match.Groups["v"].Value, NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture, out value) && value >= 0 && value <= 100;
        }
    }

    /// <summary>Bounded, read-only external Default rules. Missing/deleted/empty never creates replacements.</summary>
    public sealed class DefaultReviewRuleConfiguration
    {
        private sealed class RuleFile
        {
            public int SchemaVersion { get; set; }
            public List<DefaultTargetReviewRule> Rules { get; set; }
        }

        public List<DefaultTargetReviewRule> Rules { get; private set; } = new List<DefaultTargetReviewRule>();
        public string Status { get; private set; }
        public string Message { get; private set; }

        public static DefaultReviewRuleConfiguration Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return Result(ReviewStatusCodes.NotConfigured, "Default rule path is empty; no rules run.");
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    if (stream.Length > 131072) return Result(ReviewStatusCodes.Unavailable, "Default rule file exceeds the 128 KiB limit; no rules run.");
                    using (var reader = new StreamReader(stream, new UTF8Encoding(false, true), true))
                    {
                        var buffer = new char[131073];
                        int count = reader.ReadBlock(buffer, 0, buffer.Length);
                        if (count > 131072) return Result(ReviewStatusCodes.Unavailable, "Default rule file exceeds the read limit; no rules run.");
                        return Parse(new string(buffer, 0, count));
                    }
                }
            }
            catch (FileNotFoundException) { return Result(ReviewStatusCodes.NotConfigured, "Default rule file is missing/deleted; no rules are recreated or run."); }
            catch (DirectoryNotFoundException) { return Result(ReviewStatusCodes.NotConfigured, "Default rule file is missing/deleted; no rules are recreated or run."); }
            catch (Exception) { return Result(ReviewStatusCodes.Unavailable, "Default rule file is unreadable; no fallback rules run."); }
        }

        public static DefaultReviewRuleConfiguration Parse(string json)
        {
            try
            {
                var file = JsonConvert.DeserializeObject<RuleFile>(json, new JsonSerializerSettings
                { MaxDepth = 8, MissingMemberHandling = MissingMemberHandling.Error, TypeNameHandling = TypeNameHandling.None });
                if (file == null || file.SchemaVersion != 1 || file.Rules == null || file.Rules.Count > 100)
                    return Invalid();
                var ids = new HashSet<string>(StringComparer.Ordinal);
                foreach (var rule in file.Rules)
                {
                    double volume;
                    if (rule == null || !Regex.IsMatch(rule.Id ?? "", @"^[A-Za-z0-9_.-]{1,64}$") || !ids.Add(rule.Id) ||
                        rule.TargetTypes == null || rule.TargetTypes.Count == 0 ||
                        rule.TargetTypes.Any(type => !new[] { "PTV", "CTV", "GTV", "ITV" }.Contains(type)) ||
                        !rule.TryGetVolumePercent(out volume) ||
                        !new[] { ">", ">=", "<", "<=", "=" }.Contains(rule.Comparator) ||
                        rule.Reference != "target-prescription" || rule.PrescriptionPercent <= 0 ||
                        rule.PrescriptionPercent > 1000 || double.IsNaN(rule.PrescriptionPercent) || double.IsInfinity(rule.PrescriptionPercent))
                        return Invalid();
                }
                var result = Result(file.Rules.Any(rule => rule.Enabled) ? ReviewStatusCodes.Available : ReviewStatusCodes.NotConfigured,
                    "Default rules loaded from editable JSON; " + file.Rules.Count(rule => rule.Enabled) + " enabled. Deleted/disabled rules are not evaluated.");
                result.Rules = file.Rules.Where(rule => rule.Enabled).ToList();
                return result;
            }
            catch (Exception) { return Invalid(); }
        }

        private static DefaultReviewRuleConfiguration Invalid()
        {
            return Result(ReviewStatusCodes.Unavailable, "Default rule configuration is invalid or unsupported; no fallback rules run.");
        }

        private static DefaultReviewRuleConfiguration Result(string status, string message)
        {
            return new DefaultReviewRuleConfiguration { Status = status, Message = message };
        }
    }
}
