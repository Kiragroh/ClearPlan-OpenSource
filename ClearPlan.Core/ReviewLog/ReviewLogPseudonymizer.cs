using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ClearPlan.Core.ReviewLog
{
    /// <summary>
    /// Keyed pseudonyms (HMAC-SHA256 with a local secret). Deterministic for one key, so the same
    /// patient/plan/user links across reviews; not reversible and not recomputable without the key.
    /// The key is never written to the log. Clear identifiers are only returned in explicit clear mode.
    /// </summary>
    public sealed class ReviewLogPseudonymizer
    {
        public const int MinimumKeyBytes = 32;
        public const int PseudonymHexLength = 32; // 128-bit truncated HMAC
        private readonly byte[] _key;
        private readonly bool _clear;

        public ReviewLogPseudonymizer(byte[] key, bool logClearIdentifiers)
        {
            if (!logClearIdentifiers && (key == null || key.Length < MinimumKeyBytes))
                throw new ArgumentException("A review-log key of at least 32 bytes is required for pseudonymisation.", "key");
            _key = key == null ? null : (byte[])key.Clone();
            _clear = logClearIdentifiers;
        }

        public bool ClearIdentifiers { get { return _clear; } }
        public string IdentifierMode { get { return _clear ? "clear" : "hmac_sha256"; } }

        public string Patient(string patientId)
        {
            return Key("pat", Normalize(patientId, true), patientId);
        }

        public string Course(string patientId, string courseId)
        {
            if (string.IsNullOrWhiteSpace(courseId)) return string.Empty;
            return Key("crs", Normalize(patientId, true) + "\u001f" + Normalize(courseId, false), courseId);
        }

        /// <summary>Plan key from the DICOM plan UID; plan sums (no UID) fall back to patient/course/id.</summary>
        public string Plan(string planUid, string patientId, string courseId, string planId)
        {
            if (!string.IsNullOrWhiteSpace(planUid)) return Key("pln", "uid\u001f" + planUid.Trim(), planUid);
            if (string.IsNullOrWhiteSpace(planId)) return string.Empty;
            string composite = "id\u001f" + Normalize(patientId, true) + "\u001f" + Normalize(courseId, false) + "\u001f" + Normalize(planId, false);
            return Key("pln", composite, (courseId ?? string.Empty) + "/" + planId);
        }

        public string User(string userId)
        {
            return Key("usr", Normalize(StripDomain(userId), true), userId);
        }

        private string Key(string prefix, string normalized, string clearValue)
        {
            if (string.IsNullOrWhiteSpace(clearValue) || normalized.Length == 0) return string.Empty;
            if (_clear) return clearValue.Trim();
            using (var hmac = new HMACSHA256(_key))
            {
                byte[] hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(prefix + "\u001e" + normalized));
                var builder = new StringBuilder(prefix.Length + 1 + PseudonymHexLength);
                builder.Append(prefix).Append('_');
                for (int i = 0; i < PseudonymHexLength / 2; i++) builder.Append(hash[i].ToString("x2"));
                return builder.ToString();
            }
        }

        private static string Normalize(string value, bool caseInsensitive)
        {
            string trimmed = (value ?? string.Empty).Trim();
            return caseInsensitive ? trimmed.ToUpperInvariant() : trimmed;
        }

        private static string StripDomain(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            string[] parts = value.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length == 0 ? value : parts[parts.Length - 1];
        }

        /// <summary>Reads the base64 key file or creates a new random 32-byte key exactly once (race safe).</summary>
        public static byte[] LoadOrCreateKey(string keyFilePath)
        {
            if (string.IsNullOrWhiteSpace(keyFilePath)) throw new ArgumentException("A review-log key file path is required.", "keyFilePath");
            for (int attempt = 0; attempt < 20; attempt++)
            {
                if (File.Exists(keyFilePath)) return ReadKey(keyFilePath);
                string directory = Path.GetDirectoryName(Path.GetFullPath(keyFilePath));
                if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
                byte[] key = new byte[MinimumKeyBytes];
                using (var random = RandomNumberGenerator.Create()) random.GetBytes(key);
                try
                {
                    using (var stream = new FileStream(keyFilePath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    {
                        byte[] text = Encoding.ASCII.GetBytes(Convert.ToBase64String(key));
                        stream.Write(text, 0, text.Length);
                        stream.Flush(true);
                    }
                    return key;
                }
                catch (IOException)
                {
                    // Another process created the key concurrently; read theirs.
                    System.Threading.Thread.Sleep(25);
                }
            }
            return ReadKey(keyFilePath);
        }

        private static byte[] ReadKey(string path)
        {
            for (int attempt = 0; ; attempt++)
            {
                try
                {
                    string text = File.ReadAllText(path, Encoding.ASCII).Trim();
                    byte[] key = Convert.FromBase64String(text);
                    if (key.Length < MinimumKeyBytes) throw new InvalidDataException("Review-log key file is shorter than 32 bytes.");
                    return key;
                }
                catch (IOException) when (attempt < 20) { System.Threading.Thread.Sleep(25); }
                catch (FormatException) when (attempt < 20) { System.Threading.Thread.Sleep(25); }
            }
        }
    }

    /// <summary>Replaces known identifying terms (patient/user names and IDs) in free text with [redacted].</summary>
    public sealed class ReviewLogTextScrubber
    {
        public const int MaximumLength = 2000;
        public const string Redacted = "[redacted]";
        private readonly List<string> _terms = new List<string>();

        public ReviewLogTextScrubber(IEnumerable<string> sensitiveTerms)
        {
            foreach (string term in sensitiveTerms ?? new string[0])
            {
                if (string.IsNullOrWhiteSpace(term)) continue;
                string trimmed = term.Trim();
                if (trimmed.Length >= 3 && !_terms.Contains(trimmed)) _terms.Add(trimmed);
            }
            _terms.Sort((a, b) => b.Length.CompareTo(a.Length));
        }

        public string Scrub(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            // One pass over the original text for all terms (longest first); replacements are never re-scanned.
            var builder = new StringBuilder(text.Length);
            int position = 0;
            while (position < text.Length)
            {
                string match = null;
                foreach (string term in _terms)
                {
                    if (term.Length <= text.Length - position &&
                        string.Compare(text, position, term, 0, term.Length, StringComparison.OrdinalIgnoreCase) == 0)
                    { match = term; break; }
                }
                if (match == null) { builder.Append(text[position]); position++; }
                else { builder.Append(Redacted); position += match.Length; }
            }
            string result = builder.ToString();
            return result.Length > MaximumLength ? result.Substring(0, MaximumLength) : result;
        }
    }
}
