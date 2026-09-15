using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ClearPlan.Core.Collision
{
    public static class CollisionProfileFileSource
    {
        // One outstanding UNC read at a time, even after timeout. Never block the UI.
        private static readonly SemaphoreSlim ReadSlot = new SemaphoreSlim(1, 1);
        public static Task<CollisionProfileCatalog> LoadAsync(string path, CancellationToken token)
        {
            return LoadCatalogAsync(path, token, CollisionProfileCatalog.Parse,
                "{\"SchemaVersion\":1,\"Profiles\":[]}");
        }
        public static Task<SourceCollisionModelCatalog> LoadSourceModelsAsync(string path, CancellationToken token)
        {
            return LoadCatalogAsync(path, token, SourceCollisionModelCatalog.Parse,
                "{\"SchemaVersion\":1,\"Units\":\"mm\",\"Models\":[],\"MachineBindings\":[]}");
        }
        private static async Task<T> LoadCatalogAsync<T>(string path, CancellationToken token, Func<string, T> parse, string empty)
        {
            token.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(path)) return parse(empty);
            if (!ReadSlot.Wait(0)) throw new IOException("Previous collision-profile read is still completing.");
            var read = Task.Run(() =>
            {
                try
                {
                    using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                    using (var memory = new MemoryStream())
                    {
                        var buffer = new byte[8192]; const int limit = 256 * 1024;
                        while (true)
                        {
                            token.ThrowIfCancellationRequested();
                            int count = stream.Read(buffer, 0, Math.Min(buffer.Length, limit + 1 - (int)memory.Length));
                            if (count == 0) break;
                            memory.Write(buffer, 0, count);
                            if (memory.Length > limit) throw new IOException();
                        }
                        token.ThrowIfCancellationRequested();
                        return parse(new UTF8Encoding(false, true).GetString(memory.ToArray()).TrimStart('\uFEFF'));
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception) { throw new IOException("Collision profile could not be read or validated (UTF-8 JSON, maximum 256 KiB)."); }
                finally { ReadSlot.Release(); }
            });
            ObserveFault(read);
            using (var timer = CancellationTokenSource.CreateLinkedTokenSource(token))
            {
                try
                {
                    var completed = await Task.WhenAny(read, Task.Delay(10000, timer.Token)).ConfigureAwait(false);
                    token.ThrowIfCancellationRequested();
                    if (completed != read) throw new TimeoutException("Collision profile read timed out. No model assumed.");
                    return await read.ConfigureAwait(false);
                }
                finally { timer.Cancel(); }
            }
        }
        private static void ObserveFault(Task read)
        {
            read.ContinueWith(t => { var observed = t.Exception; }, CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
        }
    }
}
