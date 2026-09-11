using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace ClearPlan.Core.Configuration
{
    /// <summary>
    /// Explicitly rooted configuration-only working copies and append-only snapshots.
    /// Callers validate document semantics and supply the Windows actor; no external source path is accepted.
    /// </summary>
    public sealed class ConfigurationHistoryStore
    {
        public const int MaximumContentBytes = 32 * 1024 * 1024;
        private const int MaximumManifestBytes = 16 * 1024 * 1024;
        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false, true);
        private static readonly Regex SafeKey = new Regex("\\A[a-z0-9][a-z0-9_-]{0,63}\\z", RegexOptions.CultureInvariant);
        private static readonly Regex ReservedKey = new Regex("\\A(con|prn|aux|nul|com[1-9]|lpt[1-9])\\z", RegexOptions.CultureInvariant);
        private static readonly Regex Sha256Pattern = new Regex("\\A[a-f0-9]{64}\\z", RegexOptions.CultureInvariant);
        private readonly string rootDirectory;

        public ConfigurationHistoryStore(string rootDirectory)
        {
            if (string.IsNullOrWhiteSpace(rootDirectory)) throw new ArgumentException("A managed configuration directory is required.", "rootDirectory");
            this.rootDirectory = Path.GetFullPath(rootDirectory);
            EnsureNoReparsePoints(this.rootDirectory);
        }

        public ConfigurationRevision Import(string key, string category, string extension, byte[] content,
            string actor, string reason, int expectedRevision = 0, Action<byte[]> validate = null)
        {
            ValidateKey(key);
            ValidateText(category, 128, "category");
            ValidateExtension(extension);
            byte[] ownedContent = PrepareContent(content, actor, reason, validate);
            using (AcquireLock(key))
            {
                Manifest head = ReadHead(key);
                CheckExpectedAndCurrent(key, head, expectedRevision);
                ConfigurationRevision current = head.Revisions.LastOrDefault();
                if (current != null && (current.Extension != extension || current.Category != category))
                    throw new ArgumentException("An existing configuration key must keep its category and extension.");
                return Commit(key, head, category, extension, ownedContent, actor, reason,
                    current == null ? "original" : "import", null);
            }
        }

        public ConfigurationRevision Save(string key, byte[] content, string actor, string reason,
            int expectedRevision, Action<byte[]> validate = null)
        {
            ValidateKey(key);
            byte[] ownedContent = PrepareContent(content, actor, reason, validate);
            using (AcquireLock(key))
            {
                Manifest head = ReadHead(key);
                CheckExpectedAndCurrent(key, head, expectedRevision);
                ConfigurationRevision current = RequireCurrent(head);
                return Commit(key, head, current.Category, current.Extension, ownedContent, actor, reason, "edit", null);
            }
        }

        public ConfigurationRevision Restore(string key, int sourceRevision, string actor, string reason,
            int expectedRevision, Action<byte[]> validate = null)
        {
            ValidateKey(key);
            ValidateText(actor, 256, "actor");
            ValidateText(reason, 2048, "reason");
            using (AcquireLock(key))
            {
                Manifest head = ReadHead(key);
                CheckExpectedAndCurrent(key, head, expectedRevision);
                ConfigurationRevision current = RequireCurrent(head);
                byte[] content = ReadVerifiedVersion(key, head, sourceRevision);
                if (validate != null) validate((byte[])content.Clone());
                return Commit(key, head, current.Category, current.Extension, content, actor, reason, "restore", sourceRevision);
            }
        }

        public IReadOnlyList<ConfigurationRevision> ListVersions(string key)
        {
            ValidateKey(key);
            using (FileStream readLock = AcquireReadLock(key))
                return (readLock == null ? new Manifest() : ReadHead(key)).Revisions.AsEnumerable().Reverse().ToList().AsReadOnly();
        }

        public ConfigurationRevision GetCurrent(string key)
        {
            ValidateKey(key);
            using (FileStream readLock = AcquireReadLock(key))
            {
                if (readLock == null) return null;
                Manifest head = ReadHead(key);
                VerifyCurrent(key, head);
                return head.Revisions.LastOrDefault();
            }
        }

        public string GetManagedPath(string key)
        {
            ConfigurationRevision current = GetCurrent(key);
            // Active settings pin a committed snapshot, never the mutable working file.
            return current == null ? null : Path.Combine(RevisionDirectory(key, current.RevisionNumber), "content" + current.Extension);
        }

        public byte[] ReadVersion(string key, int revision)
        {
            ValidateKey(key);
            using (FileStream readLock = AcquireReadLock(key))
                return ReadVerifiedVersion(key, readLock == null ? new Manifest() : ReadHead(key), revision);
        }

        public string CreateTemporaryFile(string subdirectory, string extension, byte[] content)
        {
            if (subdirectory != ".editing" && subdirectory != ".validation")
                throw new ArgumentException("Temporary files must use .editing or .validation.", "subdirectory");
            ValidateExtension(extension);
            byte[] bytes = CopyBoundedContent(content);
            string directory = Path.Combine(rootDirectory, subdirectory);
            EnsureNoReparsePoints(directory);
            Directory.CreateDirectory(directory);
            EnsureNoReparsePoints(directory);
            string path = Path.Combine(directory, Guid.NewGuid().ToString("N") + extension);
            WriteNewFile(path, bytes);
            return path;
        }

        public void DeleteTemporaryFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A managed temporary file path is required.", "path");
            string fullPath = Path.GetFullPath(path);
            string directory = Path.GetDirectoryName(fullPath);
            Guid identifier;
            if ((!string.Equals(directory, Path.Combine(rootDirectory, ".editing"), StringComparison.OrdinalIgnoreCase) &&
                 !string.Equals(directory, Path.Combine(rootDirectory, ".validation"), StringComparison.OrdinalIgnoreCase)) ||
                !Guid.TryParseExact(Path.GetFileNameWithoutExtension(fullPath), "N", out identifier))
                throw new ArgumentException("Only generated managed temporary files may be removed.", "path");
            ValidateExtension(Path.GetExtension(fullPath));
            EnsureNoReparsePoints(fullPath);
            File.Delete(fullPath);
        }

        private ConfigurationRevision Commit(string key, Manifest head, string category, string extension,
            byte[] content, string actor, string reason, string action, int? sourceRevision)
        {
            string keyDirectory = KeyDirectory(key);
            string historyDirectory = Path.Combine(keyDirectory, "revisions");
            Directory.CreateDirectory(historyDirectory);
            EnsureNoReparsePoints(historyDirectory);
            int number = head.Revisions.Count == 0 ? 1 : checked(head.Revisions.Last().RevisionNumber + 1);
            // An interrupted write can leave an unpublished directory. Never overwrite such evidence.
            while (Directory.Exists(RevisionDirectory(key, number))) number = checked(number + 1);
            var revision = new ConfigurationRevision
            {
                Key = key, Category = category, Extension = extension, RevisionNumber = number,
                Sha256 = Hash(content), CreatedUtc = DateTime.UtcNow, Actor = actor.Trim(),
                ChangeReason = reason.Trim(), Action = action, SourceRevision = sourceRevision
            };
            var nextHead = new Manifest { Revisions = head.Revisions.Concat(new[] { revision }).ToList() };
            byte[] nextHeadBytes = JsonBytes(nextHead);
            if (nextHeadBytes.Length > MaximumManifestBytes) throw new IOException("Configuration history has reached its manifest size limit; archive it administratively before continuing.");

            string revisionDirectory = RevisionDirectory(key, number);
            string staging = Path.Combine(historyDirectory, ".pending-" + Guid.NewGuid().ToString("N"));
            string currentPath = CurrentPath(key, extension);
            string headPath = Path.Combine(keyDirectory, "head.json");
            string currentTemp = Path.Combine(keyDirectory, ".current-" + Guid.NewGuid().ToString("N") + ".tmp");
            string headTemp = Path.Combine(keyDirectory, ".head-" + Guid.NewGuid().ToString("N") + ".tmp");
            string backup = Path.Combine(keyDirectory, ".rollback-" + Guid.NewGuid().ToString("N") + ".tmp");
            bool hadCurrent = head.Revisions.Count > 0;
            bool snapshotPublished = false;
            bool currentReplaced = false;
            bool committed = false;
            bool rollbackSucceeded = true;
            try
            {
                Directory.CreateDirectory(staging);
                WriteNewFile(Path.Combine(staging, "content" + extension), content);
                WriteNewFile(Path.Combine(staging, "metadata.json"), JsonBytes(revision));
                WriteNewFile(currentTemp, content);
                WriteNewFile(headTemp, nextHeadBytes);
                Directory.Move(staging, revisionDirectory);
                snapshotPublished = true;
                if (hadCurrent) File.Replace(currentTemp, currentPath, backup);
                else File.Move(currentTemp, currentPath);
                currentReplaced = true;
                if (head.Revisions.Count > 0) File.Replace(headTemp, headPath, null);
                else File.Move(headTemp, headPath);
                committed = true;
                return revision;
            }
            catch (Exception failure)
            {
                if (currentReplaced && !committed)
                {
                    try
                    {
                        if (hadCurrent) File.Replace(backup, currentPath, null);
                        else File.Delete(currentPath);
                    }
                    catch (Exception rollbackFailure)
                    {
                        rollbackSucceeded = false;
                        throw new IOException("Configuration was not committed, and restoring its previous working copy failed. Stop editing and inspect the managed history and rollback file.",
                            new AggregateException(failure, rollbackFailure));
                    }
                }
                throw;
            }
            finally
            {
                TryDeleteFile(currentTemp);
                TryDeleteFile(headTemp);
                if (committed || rollbackSucceeded) TryDeleteFile(backup);
                if (!committed && rollbackSucceeded)
                    TryDeleteSnapshot(snapshotPublished ? revisionDirectory : staging, extension);
            }
        }

        private FileStream AcquireLock(string key)
        {
            string directory = KeyDirectory(key);
            EnsureNoReparsePoints(directory);
            Directory.CreateDirectory(directory);
            EnsureNoReparsePoints(directory);
            string path = Path.Combine(directory, ".lock");
            EnsureNoReparsePoints(path);
            // FileShare.None coordinates threads, other processes, and supported SMB shares; do not spin indefinitely.
            return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        }

        private FileStream AcquireReadLock(string key)
        {
            string directory = KeyDirectory(key);
            EnsureNoReparsePoints(directory);
            string path = Path.Combine(directory, ".lock");
            EnsureNoReparsePoints(path);
            if (!File.Exists(path))
            {
                if (File.Exists(Path.Combine(directory, "head.json")) ||
                    new[] { ".json", ".ini", ".xlsx" }.Any(extension => File.Exists(CurrentPath(key, extension))))
                    throw new InvalidDataException("Configuration history has no lock file. Inspect the managed directory before continuing.");
                return null;
            }
            return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
        }

        private Manifest ReadHead(string key)
        {
            string path = Path.Combine(KeyDirectory(key), "head.json");
            if (!File.Exists(path)) return new Manifest();
            try
            {
                Manifest head = JsonConvert.DeserializeObject<Manifest>(Utf8.GetString(ReadBoundedFile(path, MaximumManifestBytes)),
                    new JsonSerializerSettings { DateTimeZoneHandling = DateTimeZoneHandling.Utc, MaxDepth = 16, TypeNameHandling = TypeNameHandling.None });
                if (head == null || head.Schema != 1 || head.Revisions == null || head.Revisions.Count == 0)
                    throw new InvalidDataException("Invalid configuration history manifest.");
                int previous = 0;
                foreach (ConfigurationRevision revision in head.Revisions)
                {
                    if (revision == null || revision.Key != key || revision.RevisionNumber <= previous ||
                        revision.Sha256 == null || !Sha256Pattern.IsMatch(revision.Sha256) || revision.CreatedUtc == default(DateTime))
                        throw new InvalidDataException("Invalid configuration revision metadata.");
                    ValidateExtension(revision.Extension);
                    ValidateText(revision.Category, 128, "category");
                    ValidateText(revision.Actor, 256, "actor");
                    ValidateText(revision.ChangeReason, 2048, "reason");
                    if (revision.Extension != head.Revisions[0].Extension || revision.Category != head.Revisions[0].Category ||
                        !new[] { "original", "import", "edit", "restore" }.Contains(revision.Action) ||
                        (revision.Action == "restore" && (!revision.SourceRevision.HasValue || !head.Revisions.Any(r => r != null && r.RevisionNumber == revision.SourceRevision && r.RevisionNumber < revision.RevisionNumber))))
                        throw new InvalidDataException("Inconsistent configuration revision metadata.");
                    previous = revision.RevisionNumber;
                }
                return head;
            }
            catch (JsonException exception) { throw new InvalidDataException("Configuration history metadata is unreadable.", exception); }
            catch (ArgumentException exception) { throw new InvalidDataException("Configuration history metadata is invalid.", exception); }
        }

        private void CheckExpectedAndCurrent(string key, Manifest head, int expectedRevision)
        {
            int currentRevision = head.Revisions.Count == 0 ? 0 : head.Revisions.Last().RevisionNumber;
            if (expectedRevision != currentRevision)
                throw new InvalidOperationException("Configuration changed after it was loaded. Reload its current revision before saving.");
            VerifyCurrent(key, head);
        }

        private void VerifyCurrent(string key, Manifest head)
        {
            ConfigurationRevision current = head.Revisions.LastOrDefault();
            if (current == null)
            {
                if (new[] { ".json", ".ini", ".xlsx" }.Any(extension => File.Exists(CurrentPath(key, extension))))
                    throw new InvalidDataException("A working configuration exists without committed history. Inspect it before continuing.");
                return;
            }
            string path = CurrentPath(key, current.Extension);
            if (!File.Exists(path) || Hash(ReadBoundedFile(path, MaximumContentBytes)) != current.Sha256)
                throw new InvalidDataException("The managed working file does not match its committed revision. Reload is unsafe; inspect external changes or an interrupted write.");
            ReadVerifiedVersion(key, head, current.RevisionNumber);
        }

        private byte[] ReadVerifiedVersion(string key, Manifest head, int number)
        {
            ConfigurationRevision revision = head.Revisions.SingleOrDefault(item => item.RevisionNumber == number);
            if (revision == null) throw new ArgumentException("The selected revision is not in committed history.", "revision");
            string directory = RevisionDirectory(key, number);
            byte[] bytes = ReadBoundedFile(Path.Combine(directory, "content" + revision.Extension), MaximumContentBytes);
            if (Hash(bytes) != revision.Sha256) throw new InvalidDataException("The historical configuration checksum does not match. Restore is blocked.");
            byte[] metadata = ReadBoundedFile(Path.Combine(directory, "metadata.json"), 64 * 1024);
            if (!metadata.SequenceEqual(JsonBytes(revision))) throw new InvalidDataException("The historical revision metadata does not match committed history. Restore is blocked.");
            return bytes;
        }

        private static ConfigurationRevision RequireCurrent(Manifest head)
        {
            if (head.Revisions.Count == 0) throw new InvalidOperationException("Import an initial configuration before editing or restoring it.");
            return head.Revisions.Last();
        }

        private static byte[] PrepareContent(byte[] content, string actor, string reason, Action<byte[]> validate)
        {
            ValidateText(actor, 256, "actor");
            ValidateText(reason, 2048, "reason");
            byte[] ownedContent = CopyBoundedContent(content);
            if (validate != null) validate((byte[])ownedContent.Clone());
            return ownedContent;
        }

        private static byte[] CopyBoundedContent(byte[] content)
        {
            if (content == null || content.Length == 0 || content.Length > MaximumContentBytes)
                throw new ArgumentException("Configuration content must contain 1 to 33554432 bytes.", "content");
            return (byte[])content.Clone();
        }

        private static void ValidateKey(string key)
        {
            if (key == null || !SafeKey.IsMatch(key) || ReservedKey.IsMatch(key))
                throw new ArgumentException("Configuration keys must be safe lower-case names, not paths or Windows device names.", "key");
        }

        private static void ValidateExtension(string extension)
        {
            if (extension != ".json" && extension != ".ini" && extension != ".xlsx")
                throw new ArgumentException("Only .json, .ini and .xlsx configuration documents are supported.", "extension");
        }

        private static void ValidateText(string value, int maximum, string argument)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length > maximum || value.Any(c => char.IsControl(c) && c != '\r' && c != '\n' && c != '\t'))
                throw new ArgumentException("A bounded, non-empty " + argument + " is required.", argument);
        }

        private string KeyDirectory(string key) { return Path.Combine(rootDirectory, key); }
        private string CurrentPath(string key, string extension) { return Path.Combine(KeyDirectory(key), "current" + extension); }
        private string RevisionDirectory(string key, int number)
        {
            return Path.Combine(KeyDirectory(key), "revisions", number.ToString("D8", CultureInfo.InvariantCulture));
        }

        private static void EnsureNoReparsePoints(string path)
        {
            string current = Path.GetFullPath(path);
            while (!string.IsNullOrEmpty(current))
            {
                if ((File.Exists(current) || Directory.Exists(current)) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("Managed configuration paths must not contain junctions or symbolic links.");
                current = Path.GetDirectoryName(current);
            }
        }

        private static byte[] ReadBoundedFile(string path, int maximum)
        {
            EnsureNoReparsePoints(path);
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (stream.Length <= 0 || stream.Length > maximum) throw new InvalidDataException("Configuration file exceeds its permitted size or is empty.");
                var bytes = new byte[(int)stream.Length];
                int offset = 0;
                while (offset < bytes.Length)
                {
                    int read = stream.Read(bytes, offset, bytes.Length - offset);
                    if (read == 0) throw new EndOfStreamException("Configuration changed while being read.");
                    offset += read;
                }
                return bytes;
            }
        }

        private static void WriteNewFile(string path, byte[] content)
        {
            EnsureNoReparsePoints(path);
            using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                stream.Write(content, 0, content.Length);
                stream.Flush(true);
            }
        }

        private static string Hash(byte[] bytes)
        {
            using (SHA256 algorithm = SHA256.Create())
                return BitConverter.ToString(algorithm.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
        }

        private static byte[] JsonBytes(object value) { return Utf8.GetBytes(JsonConvert.SerializeObject(value, Formatting.Indented)); }
        private static void TryDeleteFile(string path)
        {
            try { File.Delete(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private static void TryDeleteSnapshot(string directory, string extension)
        {
            TryDeleteFile(Path.Combine(directory, "content" + extension));
            TryDeleteFile(Path.Combine(directory, "metadata.json"));
            try { if (Directory.Exists(directory)) Directory.Delete(directory, false); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private sealed class Manifest
        {
            public int Schema { get; set; } = 1;
            public List<ConfigurationRevision> Revisions { get; set; } = new List<ConfigurationRevision>();
        }
    }
}
