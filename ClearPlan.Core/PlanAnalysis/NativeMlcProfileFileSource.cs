using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ClearPlan.Core.PlanAnalysis
{
    public static class NativeMlcProfileFileSource
    {
        private const int MaximumBytes = 256 * 1024;
        private const int TimeoutMilliseconds = 10000;

        /// <summary>
        /// All path access and JSON parsing occur off the caller thread. Timeout/cancellation
        /// prevents publication; Windows may finish an already blocked UNC read later.
        /// </summary>
        public static async Task<NativeMlcProfileCatalog> LoadAsync(string path, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var readCancellation = CancellationTokenSource.CreateLinkedTokenSource(token);
            CancellationToken readToken = readCancellation.Token;
            Task<NativeMlcProfileCatalog> read = Task.Run(() => ReadBounded(path, readToken), readToken);
            return await AwaitReadAsync(read, readCancellation, token).ConfigureAwait(false);
        }

        private static async Task<NativeMlcProfileCatalog> AwaitReadAsync(Task<NativeMlcProfileCatalog> read,
            CancellationTokenSource readCancellation, CancellationToken token)
        {
            using (readCancellation)
            {
                CancellationToken readToken = readCancellation.Token;
                ObserveFault(read);
                try
                {
                    Task completed = await Task.WhenAny(read, Task.Delay(TimeoutMilliseconds, readToken)).ConfigureAwait(false);
                    token.ThrowIfCancellationRequested();
                    if (completed != read)
                        throw new TimeoutException("Native MLC profile loading exceeded the ten-second time limit.");
                    return await read.ConfigureAwait(false);
                }
                finally
                {
                    // Cancel both the timer and cooperative worker checks without waiting for I/O.
                    readCancellation.Cancel();
                }
            }
        }

        private static NativeMlcProfileCatalog ReadBounded(string path, CancellationToken token)
        {
            try
            {
                token.ThrowIfCancellationRequested();
                if (string.IsNullOrWhiteSpace(path)) throw new IOException();
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var data = new MemoryStream())
                {
                    var buffer = new byte[8192];
                    int total = 0;
                    while (true)
                    {
                        token.ThrowIfCancellationRequested();
                        // Read at most one byte beyond the limit; do not trust a changing file length.
                        int count = stream.Read(buffer, 0, Math.Min(buffer.Length, MaximumBytes + 1 - total));
                        token.ThrowIfCancellationRequested();
                        if (count == 0) break;
                        total += count;
                        if (total > MaximumBytes) throw new IOException();
                        data.Write(buffer, 0, count);
                    }
                    byte[] bytes = data.ToArray();
                    int offset = bytes.Length >= 3 && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf ? 3 : 0;
                    string json = new UTF8Encoding(false, true).GetString(bytes, offset, bytes.Length - offset);
                    token.ThrowIfCancellationRequested();
                    NativeMlcProfileCatalog catalog = NativeMlcProfileCatalog.Parse(json);
                    token.ThrowIfCancellationRequested();
                    return catalog;
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (NativeMlcProfileException) { throw; }
            catch (Exception)
            {
                // Never attach raw filesystem/decoder exceptions, which can disclose the path.
                throw new IOException("Native MLC profile file is missing, inaccessible, not UTF-8, or exceeds the 256 KiB limit.");
            }
        }

        private static void ObserveFault(Task read)
        {
            read.ContinueWith(faulted => { var observed = faulted.Exception; }, CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
        }
    }
}
