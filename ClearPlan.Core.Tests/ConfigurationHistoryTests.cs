using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Text;
using ClearPlan.Core.Settings;

namespace ClearPlan.Core.Tests
{
    internal static class ConfigurationHistoryTests
    {
        public static void ImportKeepsExternalSourceAndRecordsMetadata()
        {
            WithDirectory(directory =>
            {
                string source = Path.Combine(directory, "external.json");
                File.WriteAllText(source, "{\"name\":\"Rückenmark\"}", new UTF8Encoding(false));
                byte[] original = File.ReadAllBytes(source);
                dynamic store = CreateStore(Path.Combine(directory, "managed"));
                dynamic revision = store.Import("aliases", "Aliases", ".json", original, "TEST\\author", "Initial import", 0, null);
                TestAssert.Equal(1, (int)revision.RevisionNumber);
                TestAssert.Equal("original", (string)revision.Action);
                TestAssert.Equal("TEST\\author", (string)revision.Actor);
                TestAssert.Equal("Initial import", (string)revision.ChangeReason);
                TestAssert.Equal(DateTimeKind.Utc, ((DateTime)revision.CreatedUtc).Kind);
                TestAssert.Equal(64, ((string)revision.Sha256).Length);
                TestAssert.True(original.SequenceEqual(File.ReadAllBytes((string)store.GetManagedPath("aliases"))));
                TestAssert.True(original.SequenceEqual(File.ReadAllBytes(source)), "Imported external files stay untouched.");
                TestAssert.Equal("Aliases", (string)revision.Category);
                TestAssert.Equal("aliases", (string)revision.Key);
                TestAssert.Equal(".json", (string)revision.Extension);
            });
        }

        public static void EditsAndRestoreAppendHistory()
        {
            WithDirectory(directory =>
            {
                dynamic store = CreateStore(directory);
                store.Import("defaults", "Defaults", ".json", Bytes("original"), "author", "Initial", 0, null);
                store.Save("defaults", Bytes("changed"), "editor", "Adjust settings", 1, null);
                dynamic restored = store.Restore("defaults", 1, "reviewer", "Restore known version", 2, null);
                TestAssert.Equal(3, (int)restored.RevisionNumber);
                TestAssert.Equal(1, (int)restored.SourceRevision);
                TestAssert.Equal("restore", (string)restored.Action);
                TestAssert.Equal("original", File.ReadAllText((string)store.GetManagedPath("defaults")));
                TestAssert.Equal("changed", Encoding.UTF8.GetString((byte[])store.ReadVersion("defaults", 2)));
                object[] versions = ((IEnumerable)store.ListVersions("defaults")).Cast<object>().ToArray();
                TestAssert.Equal(3, versions.Length);
                TestAssert.Equal(3, (int)((dynamic)versions[0]).RevisionNumber);
                store.Import("defaults", "Defaults", ".json", Bytes("replacement"), "importer", "Reviewed replacement", 3, null);
                TestAssert.Equal("import", (string)store.GetCurrent("defaults").Action);
            });
        }

        public static void RejectsStaleWritesAndUntrackedChanges()
        {
            WithDirectory(directory =>
            {
                dynamic first = CreateStore(directory);
                dynamic second = CreateStore(directory);
                first.Import("aliases", "Aliases", ".json", Bytes("one"), "author", "Initial", 0, null);
                second.Save("aliases", Bytes("two"), "author", "Edit", 1, null);
                TestAssert.Throws<InvalidOperationException>(() => first.Save("aliases", Bytes("stale"), "other", "Stale edit", 1, null));
                TestAssert.Equal("two", File.ReadAllText((string)first.GetManagedPath("aliases")));
                File.WriteAllText(Path.Combine(directory, "aliases", "current.json"), "outside edit");
                TestAssert.Throws<InvalidDataException>(() => first.Save("aliases", Bytes("three"), "other", "Edit", 2, null));
                TestAssert.Equal("outside edit", File.ReadAllText(Path.Combine(directory, "aliases", "current.json")));
            });
        }

        public static void RejectsUnsafeNamesBoundsAndInvalidContent()
        {
            WithDirectory(directory =>
            {
                dynamic store = CreateStore(directory);
                foreach (string key in new[] { "../escape", "a/b", "a\\b", "CON", "con", "a:b", "", "aliases.", "Upper" })
                    TestAssert.Throws<ArgumentException>(() => store.Import(key, "Aliases", ".json", Bytes("{}"), "author", "Initial", 0, null));
                ArgumentException newlineKey = TestAssert.Throws<ArgumentException>(() => store.Import("aliases\n", "Aliases", ".json", Bytes("{}"), "author", "Initial", 0, null));
                TestAssert.Equal("key", newlineKey.ParamName, "Reject path input before reaching filesystem APIs.");
                TestAssert.Throws<ArgumentException>(() => store.Import("aliases", "Aliases", ".exe", Bytes("{}"), "author", "Initial", 0, null));
                TestAssert.Throws<ArgumentException>(() => store.Import("aliases", "Aliases", ".json", new byte[32 * 1024 * 1024 + 1], "author", "Initial", 0, null));
                Action<byte[]> reject = ignored => { throw new FormatException("Invalid synthetic configuration"); };
                TestAssert.Throws<FormatException>(() => store.Import("aliases", "Aliases", ".json", Bytes("bad"), "author", "Initial", 0, reject));
                TestAssert.Equal<object>(null, store.GetCurrent("aliases"));
                store.Import("aliases", "Aliases", ".json", Bytes("good"), "author", "Initial", 0, null);
                TestAssert.Throws<FormatException>(() => store.Save("aliases", Bytes("bad"), "author", "Edit", 1, reject));
                TestAssert.Throws<FormatException>(() => store.Restore("aliases", 1, "author", "Restore", 1, reject));
                TestAssert.Equal(1, (int)store.GetCurrent("aliases").RevisionNumber);
                TestAssert.Throws<ArgumentException>(() => store.Save("aliases", Bytes("{}"), "", "Edit", 1, null));
                TestAssert.Throws<ArgumentException>(() => store.Save("aliases", Bytes("{}"), "author", " ", 1, null));
            });
        }

        public static void RefusesCorruptHistoryAndPreservesCurrent()
        {
            WithDirectory(directory =>
            {
                dynamic store = CreateStore(directory);
                store.Import("defaults", "Defaults", ".json", Bytes("one"), "author", "Initial", 0, null);
                store.Save("defaults", Bytes("two"), "author", "Edit", 1, null);
                string snapshot = Path.Combine(directory, "defaults", "revisions", "00000001", "content.json");
                File.WriteAllText(snapshot, "corrupted");
                TestAssert.Throws<InvalidDataException>(() => store.ReadVersion("defaults", 1));
                TestAssert.Throws<InvalidDataException>(() => store.Restore("defaults", 1, "author", "Restore", 2, null));
                TestAssert.Equal(2, (int)store.GetCurrent("defaults").RevisionNumber);
                TestAssert.Equal("two", File.ReadAllText((string)store.GetManagedPath("defaults")));
            });
        }

        public static void FailedCommitRollsBackAndLockBlocksOtherWriters()
        {
            WithDirectory(directory =>
            {
                dynamic store = CreateStore(directory);
                store.Import("defaults", "Defaults", ".json", Bytes("one"), "author", "Initial", 0, null);
                string head = Path.Combine(directory, "defaults", "head.json");
                using (var lockedHead = new FileStream(head, FileMode.Open, FileAccess.Read, FileShare.Read))
                    TestAssert.Throws<IOException>(() => store.Save("defaults", Bytes("two"), "author", "Edit", 1, null));
                TestAssert.Equal(1, (int)store.GetCurrent("defaults").RevisionNumber);
                TestAssert.Equal("one", File.ReadAllText((string)store.GetManagedPath("defaults")));
                TestAssert.Equal(1, ((IEnumerable)store.ListVersions("defaults")).Cast<object>().Count());
                using (var lockedStore = new FileStream(Path.Combine(directory, "defaults", ".lock"), FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    TestAssert.Throws<IOException>(() => store.Save("defaults", Bytes("two"), "author", "Edit", 1, null));
                dynamic saved = store.Save("defaults", Bytes("two"), "author", "Edit after failed commit", 1, null);
                TestAssert.Equal(2, (int)saved.RevisionNumber);
            });
        }

        public static void InterruptedAndMalformedHistoryFailsClosed()
        {
            WithDirectory(directory =>
            {
                dynamic store = CreateStore(directory);
                store.Import("defaults", "Defaults", ".json", Bytes("one"), "author", "Initial", 0, null);
                TestAssert.Throws<ArgumentException>(() => store.Import("defaults", "Other", ".json", Bytes("two"), "author", "Import", 1, null));
                TestAssert.Throws<ArgumentException>(() => store.Import("defaults", "Defaults", ".ini", Bytes("two"), "author", "Import", 1, null));
                string head = Path.Combine(directory, "defaults", "head.json");
                string originalHead = File.ReadAllText(head);
                File.WriteAllText(head, "{not-json");
                TestAssert.Throws<InvalidDataException>(() => store.GetCurrent("defaults"));
                File.WriteAllText(head, originalHead);
                string metadata = Path.Combine(directory, "defaults", "revisions", "00000001", "metadata.json");
                File.WriteAllText(metadata, "{\"actor\":\"replaced\"}");
                TestAssert.Throws<InvalidDataException>(() => store.Restore("defaults", 1, "author", "Restore", 1, null));
                File.Delete(head);
                TestAssert.Throws<InvalidDataException>(() => store.GetCurrent("defaults"));
                TestAssert.Throws<InvalidDataException>(() => store.Import("defaults", "Defaults", ".json", Bytes("two"), "author", "Import", 0, null));
                TestAssert.Equal("one", File.ReadAllText(Path.Combine(directory, "defaults", "current.json")));
            });
        }

        public static void MetadataValidationRejectsNullRowsWithoutUnexpectedErrors()
        {
            WithDirectory(directory =>
            {
                dynamic store = CreateStore(directory);
                store.Import("defaults", "Defaults", ".json", Bytes("one"), "author", "Initial", 0, null);
                store.Restore("defaults", 1, "author", "Restore", 1, null);
                string head = Path.Combine(directory, "defaults", "head.json");
                var json = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(head));
                var revisions = (Newtonsoft.Json.Linq.JArray)json["Revisions"];
                revisions[1]["source_revision"] = 999;
                revisions.Add(Newtonsoft.Json.Linq.JValue.CreateNull());
                // The malformed row must be a controlled integrity error, not a null-reference failure.
                TestAssert.Throws<InvalidDataException>(() =>
                {
                    File.WriteAllText(head, json.ToString());
                    store.GetCurrent("defaults");
                });
            });
        }

        public static void ReadsDoNotCreateAStoreAndActivationPinsARevision()
        {
            WithDirectory(directory =>
            {
                string managed = Path.Combine(directory, "not-created");
                dynamic store = CreateStore(managed);
                TestAssert.Equal<object>(null, store.GetCurrent("aliases"));
                TestAssert.Equal<string>(null, (string)store.GetManagedPath("aliases"));
                TestAssert.Equal(0, ((IEnumerable)store.ListVersions("aliases")).Cast<object>().Count());
                TestAssert.Throws<ArgumentException>(() => store.ReadVersion("aliases", 1));
                TestAssert.False(Directory.Exists(managed), "Reading absent configuration must not create its repository.");
                store.Import("aliases", "Aliases", ".json", Bytes("one"), "author", "Initial", 0, null);
                string activePath = store.GetManagedPath("aliases");
                store.Save("aliases", Bytes("two"), "author", "Edit", 1, null);
                TestAssert.Equal("one", File.ReadAllText(activePath), "An activated settings path must pin its committed snapshot.");
                TestAssert.Equal("two", File.ReadAllText((string)store.GetManagedPath("aliases")));
            });
        }

        public static void TemporaryFilesAreBoundedAndConfined()
        {
            WithDirectory(directory =>
            {
                dynamic store = CreateStore(directory);
                TestAssert.NotNull(((object)store).GetType().GetMethod("CreateTemporaryFile"), "Editors need a guarded managed temporary-file API.");
                string temporary = store.CreateTemporaryFile(".editing", ".xlsx", Bytes("synthetic"));
                TestAssert.Equal(Path.Combine(directory, ".editing"), Path.GetDirectoryName(temporary));
                TestAssert.Equal("synthetic", File.ReadAllText(temporary));
                TestAssert.Throws<ArgumentException>(() => store.CreateTemporaryFile("../external", ".xlsx", Bytes("synthetic")));
                TestAssert.Throws<ArgumentException>(() => store.CreateTemporaryFile(".editing", ".exe", Bytes("synthetic")));
                TestAssert.Throws<ArgumentException>(() => store.DeleteTemporaryFile(Path.Combine(directory, "source.xlsx")));
                store.DeleteTemporaryFile(temporary);
                TestAssert.False(File.Exists(temporary));
            });
        }

        private static dynamic CreateStore(string directory)
        {
            Type type = typeof(ClearPlanSettingsModel).Assembly.GetType("ClearPlan.Core.Configuration.ConfigurationHistoryStore");
            TestAssert.NotNull(type, "Configuration history must have an ESAPI-free store.");
            return Activator.CreateInstance(type, directory);
        }

        private static byte[] Bytes(string value) { return Encoding.UTF8.GetBytes(value); }

        private static void WithDirectory(Action<string> action)
        {
            string directory = Path.Combine(Path.GetTempPath(), "ClearPlanConfigurationTests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try { action(directory); }
            finally { Directory.Delete(directory, true); }
        }
    }
}
