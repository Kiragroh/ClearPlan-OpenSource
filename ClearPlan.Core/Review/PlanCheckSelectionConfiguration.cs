using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class PlanCheckSelectionEntry
    {
        [JsonProperty("code")] public string Code { get; set; }
        [JsonProperty("enabled")] public bool? Enabled { get; set; }
    }

    /// <summary>Only controls inclusion of already-calculated legacy findings in review/report.
    /// It neither skips legacy calculation nor disables native Eclipse warnings.</summary>
    public sealed class PlanCheckSelectionConfiguration
    {
        private sealed class SelectionFile
        {
            public int SchemaVersion { get; set; }
            public List<PlanCheckSelectionEntry> Checks { get; set; }
        }
        public List<PlanCheckSelectionEntry> Checks { get; private set; } = new List<PlanCheckSelectionEntry>();
        public string Status { get; private set; }
        public string Message { get; private set; }

        public static PlanCheckSelectionConfiguration Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return Missing();
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    if (stream.Length > 131072) return Invalid();
                    using (var reader = new StreamReader(stream, new UTF8Encoding(false, true), true))
                    {
                        var buffer = new char[131073];
                        int count = reader.ReadBlock(buffer, 0, buffer.Length);
                        return count > 131072 ? Invalid() : Parse(new string(buffer, 0, count));
                    }
                }
            }
            catch (FileNotFoundException) { return Missing(); }
            catch (DirectoryNotFoundException) { return Missing(); }
            catch (Exception) { return Invalid(); }
        }

        public static PlanCheckSelectionConfiguration Parse(string json)
        {
            if (json == null || json.Length > 131072) return Invalid();
            try
            {
                var file = JsonConvert.DeserializeObject<SelectionFile>(json, new JsonSerializerSettings {
                    MaxDepth = 6, MissingMemberHandling = MissingMemberHandling.Error, TypeNameHandling = TypeNameHandling.None });
                if (file == null || file.SchemaVersion != 1 || file.Checks == null || file.Checks.Count > 1000) return Invalid();
                var codes = new HashSet<string>(StringComparer.Ordinal);
                if (file.Checks.Any(check => check == null || !check.Enabled.HasValue ||
                    string.IsNullOrWhiteSpace(check.Code) || check.Code != BaseCheckCode(check.Code) || !codes.Add(check.Code))) return Invalid();
                return new PlanCheckSelectionConfiguration { Status = ReviewStatusCodes.Available, Checks = file.Checks,
                    Message = "PlanCheck inclusion loaded. This filters existing legacy findings in review/report only; legacy calculations still run. Native Eclipse warnings are unaffected." };
            }
            catch (Exception) { return Invalid(); }
        }

        public static string Serialize(IEnumerable<PlanCheckSelectionEntry> checks)
        {
            string json = JsonConvert.SerializeObject(new { schemaVersion = 1,
                checks = (checks ?? Enumerable.Empty<PlanCheckSelectionEntry>()).OrderBy(c => c.Code, StringComparer.Ordinal).ToList() }, Formatting.Indented);
            if (Parse(json).Status == ReviewStatusCodes.Unavailable) throw new FormatException("Invalid PlanCheck inclusion entries.");
            return json;
        }

        public static string BaseCheckCode(string code)
        {
            // CreateUniqueStableId adds -2, -3, ... to repeat findings. Never persist
            // generated plan-check-N IDs or reference-point labels as algorithm identities.
            var match = Regex.Match(code ?? "", @"^(?<base>[0-9]{1,9}|PC[0-9]{1,8})(?:-(?:[2-9]|[1-9][0-9]{1,5}))?$", RegexOptions.CultureInvariant);
            return match.Success ? match.Groups["base"].Value : null;
        }

        public static bool CanConfigure(ReviewCheckRow row)
        {
            return row != null && BaseCheckCode(row.CheckCode) != null &&
                (row.Category == "Default" || row.Category == "PlanCheck" || row.CheckCode.StartsWith("PC", StringComparison.Ordinal));
        }

        public bool IsEnabled(string baseCode)
        {
            return Status != ReviewStatusCodes.Available || !Checks.Any(c => c.Code == baseCode && c.Enabled == false);
        }

        public List<ReviewCheckRow> Filter(IEnumerable<ReviewCheckRow> rows, out int disabledCount)
        {
            var result = new List<ReviewCheckRow>();
            disabledCount = 0;
            foreach (var row in rows ?? Enumerable.Empty<ReviewCheckRow>())
            {
                if (CanConfigure(row) && !IsEnabled(BaseCheckCode(row.CheckCode))) disabledCount++;
                else result.Add(row);
            }
            return result;
        }

        private static PlanCheckSelectionConfiguration Missing()
        { return new PlanCheckSelectionConfiguration { Status = ReviewStatusCodes.NotConfigured,
            Message = "No PlanCheck inclusion configuration; all legacy findings remain included. No file was created." }; }
        private static PlanCheckSelectionConfiguration Invalid()
        { return new PlanCheckSelectionConfiguration { Status = ReviewStatusCodes.Unavailable,
            Message = "PlanCheck inclusion configuration is invalid, unreadable or unsupported. No finding is hidden; configuration requires review. Legacy calculation still ran; no aggregate pass may be inferred." }; }
    }
}
