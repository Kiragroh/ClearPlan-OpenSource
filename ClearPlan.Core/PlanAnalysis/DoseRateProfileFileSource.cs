using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ClearPlan.Core.PlanAnalysis
{
    public static class DoseRateProfileFileSource
    {
        // Optional network configuration must never block the ESAPI owner/UI thread.
        public static async Task<DoseRateEstimationProfileCatalog> LoadAsync(string path, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(path)) return null;
            using (var cancellation = CancellationTokenSource.CreateLinkedTokenSource(token))
            {
                var readToken = cancellation.Token;
                var read = Task.Run(() => Read(path, readToken), readToken);
                ObserveFault(read);
                try
                {
                    var completed = await Task.WhenAny(read, Task.Delay(10000, readToken)).ConfigureAwait(false);
                    token.ThrowIfCancellationRequested();
                    if (completed != read) throw new TimeoutException("Dose-rate profile read exceeded ten seconds.");
                    return await read.ConfigureAwait(false);
                }
                finally { cancellation.Cancel(); }
            }
        }

        private static DoseRateEstimationProfileCatalog Read(string path, CancellationToken token)
        {
            const int maximum = 65536;
            try
            {
                token.ThrowIfCancellationRequested();
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var data = new MemoryStream())
                {
                    var buffer = new byte[4096];
                    while (true)
                    {
                        token.ThrowIfCancellationRequested();
                        int count = stream.Read(buffer, 0, Math.Min(buffer.Length, maximum + 1 - (int)data.Length));
                        token.ThrowIfCancellationRequested();
                        if (count == 0) break;
                        data.Write(buffer, 0, count);
                        if (data.Length > maximum) throw new IOException();
                    }
                    byte[] bytes = data.ToArray();
                    int offset = bytes.Length >= 3 && bytes[0] == 239 && bytes[1] == 187 && bytes[2] == 191 ? 3 : 0;
                    var catalog = DoseRateEstimationProfileCatalog.Parse(new UTF8Encoding(false, true).GetString(bytes, offset, bytes.Length - offset));
                    token.ThrowIfCancellationRequested();
                    return catalog;
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception)
            {
                // Do not disclose filesystem paths or untrusted configuration content.
                throw new IOException("Dose-rate profiles are missing, invalid, inaccessible or exceed 64 KiB.");
            }
        }

        private static void ObserveFault(Task read)
        {
            read.ContinueWith(faulted => { var observed = faulted.Exception; }, CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
        }
    }
}
