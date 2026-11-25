using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ClearPlan.Core.Configuration;

namespace ClearPlan.Core.ReviewLog
{
    /// <summary>Describes a configuration file by content hash and, if it is a managed revision, its revision metadata.</summary>
    public static class ReviewLogRulesetResolver
    {
        public const long MaximumHashedBytes = 64L * 1024 * 1024;

        public static ReviewLogConfigurationInput Describe(string configKey, string category, string resolvedPath, string managedRoot)
        {
            var input = new ReviewLogConfigurationInput { ConfigKey = configKey, Category = category };
            if (string.IsNullOrWhiteSpace(resolvedPath)) { input.Availability = "not_configured"; return input; }
            try
            {
                if (Directory.Exists(resolvedPath)) { input.Availability = "not_hashed"; return input; }
                if (!File.Exists(resolvedPath)) { input.Availability = "unavailable"; return input; }
                var info = new FileInfo(resolvedPath);
                if (info.Length > MaximumHashedBytes) { input.Availability = "not_hashed"; return input; }
                input.Sha256 = Sha256Hex(File.ReadAllBytes(resolvedPath));
                input.Availability = "available";
            }
            catch (Exception) { input.Availability = "unavailable"; return input; }
            TryAttachManagedRevision(input, resolvedPath, managedRoot);
            return input;
        }

        /// <summary>Managed snapshots live at &lt;root&gt;/&lt;key&gt;/revisions/&lt;NNNNNNNN&gt;/content.&lt;ext&gt;.</summary>
        private static void TryAttachManagedRevision(ReviewLogConfigurationInput input, string path, string managedRoot)
        {
            if (string.IsNullOrWhiteSpace(managedRoot)) return;
            try
            {
                string full = Path.GetFullPath(path);
                string root = Path.GetFullPath(managedRoot).TrimEnd('\\', '/') + Path.DirectorySeparatorChar;
                if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase)) return;
                string[] parts = full.Substring(root.Length).Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length != 4 || !string.Equals(parts[1], "revisions", StringComparison.OrdinalIgnoreCase)) return;
                int number;
                if (!int.TryParse(parts[2], NumberStyles.None, CultureInfo.InvariantCulture, out number)) return;
                var store = new ConfigurationHistoryStore(managedRoot);
                ConfigurationRevision revision = store.ListVersions(parts[0]).FirstOrDefault(r => r.RevisionNumber == number);
                if (revision == null || !string.Equals(revision.Sha256, input.Sha256, StringComparison.OrdinalIgnoreCase)) return;
                input.RevisionNumber = revision.RevisionNumber;
                input.RevisionCreatedUtc = new DateTimeOffset(DateTime.SpecifyKind(revision.CreatedUtc, DateTimeKind.Utc));
                if (string.IsNullOrWhiteSpace(input.Category)) input.Category = revision.Category;
            }
            catch (Exception)
            {
                // Revision metadata is optional provenance; the content hash remains authoritative.
            }
        }

        public static string Sha256Hex(byte[] content)
        {
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(content ?? new byte[0]);
                var builder = new StringBuilder(64);
                foreach (byte b in hash) builder.Append(b.ToString("x2"));
                return builder.ToString();
            }
        }

        public static string Sha256Hex(string canonicalText)
        {
            return Sha256Hex(Encoding.UTF8.GetBytes(canonicalText ?? string.Empty));
        }

        public static string FileSha256(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(stream);
                var builder = new StringBuilder(64);
                foreach (byte b in hash) builder.Append(b.ToString("x2"));
                return builder.ToString();
            }
        }
    }
}
