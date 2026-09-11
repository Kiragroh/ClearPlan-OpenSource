using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Tests
{
    internal static class NativeMlcFileSourceTests
    {
        private const string ValidJson = "{\"SchemaVersion\":1,\"Profiles\":[{\"Model\":\"Synthetic MLC ä\",\"NativeLeafCount\":1,\"JawMode\":\"None\",\"Evidence\":\"Synthetic test\",\"Layers\":[{\"Label\":\"Synthetic\",\"LeafTravelAxis\":\"X\",\"SourceLeafIndices\":[0],\"LeafBoundariesMm\":[-5,5]}]}]}";

        public static void ValidUtf8Json()
        {
            WithDirectory(directory =>
            {
                foreach (bool bom in new[] { false, true })
                {
                    string path = Path.Combine(directory, "synthetic.json");
                    File.WriteAllText(path, ValidJson, new UTF8Encoding(bom));
                    var catalog = NativeMlcProfileFileSource.LoadAsync(path, CancellationToken.None).GetAwaiter().GetResult();
                    TestAssert.Equal(1, catalog.Profiles.Count); TestAssert.Equal("Synthetic MLC ä", catalog.Profiles[0].Model);
                }
            });
        }
        public static void MissingAndInaccessiblePaths()
        {
            WithDirectory(directory =>
            {
                foreach (string path in new[] { null, "", "  ", Path.Combine(directory, "private-missing.json"), directory })
                {
                    var error = TestAssert.Throws<IOException>(() => NativeMlcProfileFileSource.LoadAsync(path, CancellationToken.None).GetAwaiter().GetResult());
                    TestAssert.True(error.InnerException == null); TestAssert.False(error.Message.Contains(directory) || error.Message.Contains("private-missing"));
                }
                string locked = Path.Combine(directory, "private-locked.json");
                File.WriteAllText(locked, ValidJson);
                using (var handle = new FileStream(locked, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    var error = TestAssert.Throws<IOException>(() => NativeMlcProfileFileSource.LoadAsync(locked, CancellationToken.None).GetAwaiter().GetResult());
                    TestAssert.True(error.InnerException == null); TestAssert.False(error.Message.Contains("private-locked"));
                }
            });
        }
        public static void BoundedStrictUtf8()
        {
            WithDirectory(directory =>
            {
                string path = Path.Combine(directory, "synthetic-bounded.json");
                byte[] valid = new UTF8Encoding(false).GetBytes(ValidJson);
                byte[] exact = new byte[256 * 1024];
                for (int i = 0; i < exact.Length; i++) exact[i] = 32;
                Buffer.BlockCopy(valid, 0, exact, 0, valid.Length); File.WriteAllBytes(path, exact);
                TestAssert.Equal(1, NativeMlcProfileFileSource.LoadAsync(path, CancellationToken.None).GetAwaiter().GetResult().Profiles.Count);
                File.WriteAllBytes(path, new byte[256 * 1024 + 1]);
                TestAssert.Throws<IOException>(() => NativeMlcProfileFileSource.LoadAsync(path, CancellationToken.None).GetAwaiter().GetResult());
                File.WriteAllBytes(path, new byte[] { 0xc3, 0x28 });
                TestAssert.Throws<IOException>(() => NativeMlcProfileFileSource.LoadAsync(path, CancellationToken.None).GetAwaiter().GetResult());
                File.WriteAllText(path, "private-invalid-json");
                var error = TestAssert.Throws<NativeMlcProfileException>(() => NativeMlcProfileFileSource.LoadAsync(path, CancellationToken.None).GetAwaiter().GetResult());
                TestAssert.True(error.InnerException == null); TestAssert.False(error.Message.Contains("private-invalid-json"));
            });
        }
        public static void CancellationAndTimeout()
        {
            var preCancelled = new CancellationToken(true);
            TestAssert.Throws<OperationCanceledException>(() => NativeMlcProfileFileSource.LoadAsync("unused-before-cancellation", preCancelled).GetAwaiter().GetResult());
            UnfinishedRead(true);
            UnfinishedRead(false);
        }
        private static void UnfinishedRead(bool cancel)
        {
            // Exercise the production deadline boundary without relying on a real UNC outage.
            // .NET Framework FileStream rejects named pipes, so they cannot model slow file I/O.
            using (var cancellation = new CancellationTokenSource())
            {
                var completion = new TaskCompletionSource<NativeMlcProfileCatalog>();
                var workerCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellation.Token);
                MethodInfo awaitMethod = typeof(NativeMlcProfileFileSource).GetMethod("AwaitReadAsync", BindingFlags.NonPublic | BindingFlags.Static);
                TestAssert.NotNull(awaitMethod);
                var watch = Stopwatch.StartNew();
                var pending = (Task<NativeMlcProfileCatalog>)awaitMethod.Invoke(null, new object[] { completion.Task, workerCancellation, cancellation.Token });
                TestAssert.True(watch.Elapsed < TimeSpan.FromSeconds(2), "Starting profile loading must not block the caller.");
                TestAssert.False(pending.IsCompleted, "A blocked read must remain asynchronous.");
                if (cancel)
                {
                    cancellation.Cancel();
                    TestAssert.Throws<OperationCanceledException>(() => pending.GetAwaiter().GetResult());
                    TestAssert.True(watch.Elapsed < TimeSpan.FromSeconds(5), "Cancellation must not wait for the blocked worker.");
                }
                else
                {
                    var error = TestAssert.Throws<TimeoutException>(() => pending.GetAwaiter().GetResult());
                    TestAssert.True(watch.Elapsed >= TimeSpan.FromSeconds(9) && watch.Elapsed < TimeSpan.FromSeconds(15), "Profile timeout must bound caller latency to approximately ten seconds.");
                    TestAssert.True(error.InnerException == null);
                }
                if (cancel) completion.SetException(new IOException("Synthetic late read failure"));
                else completion.SetResult(NativeMlcProfileCatalog.Parse(ValidJson));
                TestAssert.True(cancel ? pending.IsCanceled : pending.IsFaulted, "A late worker must never replace the canceled/timed-out outcome.");
            }
        }
        private static void WithDirectory(Action<string> action)
        {
            string directory = Path.Combine(Path.GetTempPath(), "ClearPlanMlcFileTests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try { action(directory); }
            finally { Directory.Delete(directory, true); }
        }
    }
}
