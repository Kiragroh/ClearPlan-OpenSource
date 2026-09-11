using System;
using System.Threading;
using System.Threading.Tasks;

namespace ClearPlan.Core.Integration
{
    /// <summary>A blocked SMB operation cannot be interrupted, but must not retain the UI or a late file lease.</summary>
    public static class AriaFileWork
    {
        public static async Task<T> RunAsync<T>(Func<T> work,CancellationToken cancellation,Action<T> discard=null)
        {
            cancellation.ThrowIfCancellationRequested();
            var operation=Task.Run(work);
            using(var timer=CancellationTokenSource.CreateLinkedTokenSource(cancellation))
            {
                timer.CancelAfter(TimeSpan.FromSeconds(10));
                var deadline=Task.Delay(-1,timer.Token);
                if(await Task.WhenAny(operation,deadline).ConfigureAwait(false)!=operation || cancellation.IsCancellationRequested)
                {
                    _=operation.ContinueWith(completed=> {
                        if(completed.IsFaulted){var observed=completed.Exception;}
                        else if(completed.Status==TaskStatus.RanToCompletion && discard!=null)
                            try{discard(completed.Result);}catch(Exception){}
                    },TaskScheduler.Default);
                    cancellation.ThrowIfCancellationRequested();
                    throw new TimeoutException("Report storage did not respond within the configured operation limit.");
                }
                return await operation.ConfigureAwait(false);
            }
        }
        public static void DisposeLater(IDisposable lease)
        {
            if(lease==null) return;
            _=Task.Run(()=> {try{lease.Dispose();}catch(Exception){}});
        }
    }
}
