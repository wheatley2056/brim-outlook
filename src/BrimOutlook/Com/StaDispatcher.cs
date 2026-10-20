using System.Collections.Concurrent;
using System.Runtime.Versioning;

namespace BrimOutlook.Com;

/// <summary>
/// Eigener STA-Thread für alle Aufrufe an Outlook. Das Objektmodell von Outlook ist ein STA-Server;
/// so laufen alle Aufrufe nacheinander auf einem Thread, und der Nachrichtenfilter greift.
/// </summary>
[SupportedOSPlatform("windows")]
internal sealed class StaDispatcher : IDisposable
{
    private readonly BlockingCollection<Action> _queue = [];
    private readonly Thread _thread;

    public StaDispatcher(string name)
    {
        _thread = new Thread(Run) { IsBackground = true, Name = name };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public Task<T> InvokeAsync<T>(Func<T> work)
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        try
        {
            _queue.Add(() =>
            {
                try
                {
                    tcs.SetResult(work());
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });
        }
        catch (InvalidOperationException)
        {
            tcs.SetException(new ObjectDisposedException(nameof(StaDispatcher)));
        }
        return tcs.Task;
    }

    public Task InvokeAsync(Action work) => InvokeAsync(() =>
    {
        work();
        return true;
    });

    private void Run()
    {
        MessageFilter.Register();
        try
        {
            foreach (var work in _queue.GetConsumingEnumerable())
            {
                work();
            }
        }
        finally
        {
            MessageFilter.Revoke();
        }
    }

    public void Dispose()
    {
        _queue.CompleteAdding();
        _thread.Join(TimeSpan.FromSeconds(1));
    }
}
