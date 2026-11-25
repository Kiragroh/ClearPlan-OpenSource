using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.ReviewLog
{
    public sealed class ReviewLogSchemaException : IOException
    {
        public ReviewLogSchemaException(string message) : base(message) { }
    }

    public sealed class ReviewLogWriteResult
    {
        public ReviewLogWriteResult() { RowCounts = new Dictionary<string, int>(StringComparer.Ordinal); }
        public Guid ReviewId { get; set; }
        public string Directory { get; set; }
        public Dictionary<string, int> RowCounts { get; private set; }
    }

    /// <summary>
    /// Append-only writer for the five UTF-8 (no BOM) RFC 4180 tables. A directory-wide lock file serialises
    /// concurrent writers across processes/workstations (bounded retry). Existing files must carry the exact
    /// expected header; a different header is never appended to (schema change => new directory/files).
    /// </summary>
    public sealed class ReviewLogWriter
    {
        public const string LockFileName = ".review-log.lock";
        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);
        private readonly string _directory;
        private readonly int _lockTimeoutMilliseconds;

        public ReviewLogWriter(string directory, int lockTimeoutMilliseconds = 5000)
        {
            if (string.IsNullOrWhiteSpace(directory)) throw new ArgumentException("A review-log directory is required.", "directory");
            _directory = directory;
            _lockTimeoutMilliseconds = Math.Max(0, lockTimeoutMilliseconds);
        }

        public ReviewLogWriteResult Write(ReviewLogRecord record)
        {
            if (record == null) throw new ArgumentNullException("record");
            Directory.CreateDirectory(_directory);
            var result = new ReviewLogWriteResult { ReviewId = record.ReviewId, Directory = _directory };
            using (AcquireLock())
            {
                // Validate every header before the first append so a schema mismatch never leaves partial rows.
                foreach (string table in ReviewLogSchema.WriteOrder) VerifyHeader(table);
                foreach (string table in ReviewLogSchema.WriteOrder)
                {
                    List<string[]> rows = record.Rows[table];
                    Append(table, rows);
                    result.RowCounts[table] = rows.Count;
                }
            }
            return result;
        }

        private IDisposable AcquireLock()
        {
            string path = Path.Combine(_directory, LockFileName);
            var watch = Stopwatch.StartNew();
            var random = new Random();
            while (true)
            {
                try
                {
                    return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.None);
                }
                catch (IOException)
                {
                    if (watch.ElapsedMilliseconds >= _lockTimeoutMilliseconds)
                        throw new TimeoutException("Review-log lock not acquired within " + _lockTimeoutMilliseconds + " ms.");
                    Thread.Sleep(20 + random.Next(60));
                }
                catch (UnauthorizedAccessException)
                {
                    if (watch.ElapsedMilliseconds >= _lockTimeoutMilliseconds) throw;
                    Thread.Sleep(20 + random.Next(60));
                }
            }
        }

        private string PathFor(string table)
        {
            return Path.Combine(_directory, ReviewLogSchema.FileName(table));
        }

        private static string HeaderLine(string table)
        {
            return ReviewLogCsv.FormatLine(ReviewLogSchema.Columns(table));
        }

        private void VerifyHeader(string table)
        {
            string path = PathFor(table);
            if (!File.Exists(path)) return;
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (stream.Length == 0) return;
                string expected = HeaderLine(table);
                byte[] buffer = new byte[Math.Min(stream.Length, Utf8.GetByteCount(expected) + 3)];
                int read = 0;
                while (read < buffer.Length)
                {
                    int n = stream.Read(buffer, read, buffer.Length - read);
                    if (n <= 0) break;
                    read += n;
                }
                string actual = Utf8.GetString(buffer, 0, read);
                if (actual.Length > 0 && actual[0] == '\uFEFF') actual = actual.Substring(1);
                if (!actual.StartsWith(expected, StringComparison.Ordinal))
                    throw new ReviewLogSchemaException("Existing " + ReviewLogSchema.FileName(table) +
                        " has a different header; refusing to append schema version " + ReviewLogSchema.SchemaVersion + " rows.");
            }
        }

        private void Append(string table, List<string[]> rows)
        {
            IList<string> columns = ReviewLogSchema.Columns(table);
            var builder = new StringBuilder();
            string path = PathFor(table);
            using (var stream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read))
            {
                if (stream.Length == 0) builder.Append(HeaderLine(table));
                else if (!EndsWithNewline(stream)) builder.Append(ReviewLogCsv.LineTerminator); // repair an interrupted last line
                foreach (string[] row in rows)
                {
                    if (row.Length != columns.Count)
                        throw new InvalidOperationException(table + " row has " + row.Length + " fields; expected " + columns.Count + ".");
                    builder.Append(ReviewLogCsv.FormatLine(row));
                }
                if (builder.Length == 0) return;
                byte[] bytes = Utf8.GetBytes(builder.ToString());
                stream.Seek(0, SeekOrigin.End);
                stream.Write(bytes, 0, bytes.Length);
                stream.Flush(true);
            }
        }

        private static bool EndsWithNewline(FileStream stream)
        {
            stream.Seek(-1, SeekOrigin.End);
            return stream.ReadByte() == '\n';
        }
    }

    /// <summary>Convenience facade: key handling + build + append. Throws on failure; callers in clinical paths must guard.</summary>
    public static class ReviewLogService
    {
        public static ReviewLogWriteResult Write(ReviewSnapshot snapshot, ReviewLogContext context, ReviewLogOptions options)
        {
            if (options == null) throw new ArgumentNullException("options");
            if (string.IsNullOrWhiteSpace(options.Directory)) throw new InvalidOperationException("Review-log directory is not configured.");
            byte[] key = options.LogClearIdentifiers ? null : ReviewLogPseudonymizer.LoadOrCreateKey(options.KeyFilePath);
            var pseudonymizer = new ReviewLogPseudonymizer(key, options.LogClearIdentifiers);
            if (key != null) Array.Clear(key, 0, key.Length);
            ReviewLogRecord record = ReviewLogRecordBuilder.Build(snapshot, context, pseudonymizer, options);
            return new ReviewLogWriter(options.Directory, options.LockTimeoutMilliseconds).Write(record);
        }
    }
}
