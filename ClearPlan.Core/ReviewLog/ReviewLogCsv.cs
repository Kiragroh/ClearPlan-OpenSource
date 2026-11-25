using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ClearPlan.Core.ReviewLog
{
    /// <summary>RFC 4180 CSV (comma, CRLF, double-quote escaping) with invariant-culture value formatting.</summary>
    public static class ReviewLogCsv
    {
        public const string LineTerminator = "\r\n";

        public static string Escape(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            bool quote = value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0 ||
                char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1]);
            return quote ? "\"" + value.Replace("\"", "\"\"") + "\"" : value;
        }

        public static string FormatLine(IList<string> fields)
        {
            var builder = new StringBuilder();
            for (int i = 0; i < fields.Count; i++)
            {
                if (i > 0) builder.Append(',');
                builder.Append(Escape(fields[i]));
            }
            builder.Append(LineTerminator);
            return builder.ToString();
        }

        /// <summary>Finite numbers only, round-trip precision; null/NaN/Infinity become empty (unavailable), never 0.</summary>
        public static string Number(double? value)
        {
            if (!value.HasValue || double.IsNaN(value.Value) || double.IsInfinity(value.Value)) return string.Empty;
            return value.Value.ToString("R", CultureInfo.InvariantCulture);
        }

        public static string Integer(int? value)
        {
            return value.HasValue ? value.Value.ToString(CultureInfo.InvariantCulture) : string.Empty;
        }

        public static string Flag(bool? value)
        {
            return value.HasValue ? (value.Value ? "1" : "0") : string.Empty;
        }

        /// <summary>ISO 8601 UTC with millisecond precision and a literal Z suffix.</summary>
        public static string Utc(DateTimeOffset? value)
        {
            if (!value.HasValue) return string.Empty;
            return value.Value.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InvariantCulture);
        }

        public static bool TryParseInvariantNumber(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;
            if (!double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return false;
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        /// <summary>Minimal RFC 4180 reader used for explicit plan lists and header checks.</summary>
        public static List<List<string>> Parse(string text)
        {
            var rows = new List<List<string>>();
            if (string.IsNullOrEmpty(text)) return rows;
            if (text[0] == '\uFEFF') text = text.Substring(1);
            var row = new List<string>();
            var field = new StringBuilder();
            bool inQuotes = false, fieldStarted = false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                        else inQuotes = false;
                    }
                    else field.Append(c);
                    continue;
                }
                if (c == '"' && field.Length == 0) { inQuotes = true; fieldStarted = true; }
                else if (c == ',') { row.Add(field.ToString()); field.Clear(); fieldStarted = false; }
                else if (c == '\r' || c == '\n')
                {
                    if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    row.Add(field.ToString()); field.Clear(); fieldStarted = false;
                    if (!(row.Count == 1 && row[0].Length == 0)) rows.Add(row);
                    row = new List<string>();
                }
                else { field.Append(c); fieldStarted = true; }
            }
            if (inQuotes) throw new FormatException("Unterminated quoted CSV field.");
            if (fieldStarted || field.Length > 0 || row.Count > 0)
            {
                row.Add(field.ToString());
                if (!(row.Count == 1 && row[0].Length == 0)) rows.Add(row);
            }
            return rows;
        }
    }
}
