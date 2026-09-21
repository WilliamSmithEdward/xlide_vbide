using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// Native Enter blocks on error notices. A short-lived watcher captures and acknowledges them
/// off the host thread, equally for UI and API callers, without ever issuing Reset.
/// </summary>
internal sealed class NativeImmediateErrors : IDisposable
{
    private readonly object _gate = new();
    private readonly HashSet<string> _seen;
    private readonly Timer _timer;
    private string? _message;
    private bool _closed;

    public NativeImmediateErrors()
    {
        _seen = DialogWatch.Dialogs().Select(dialog => dialog.Window).ToHashSet(StringComparer.Ordinal);
        _timer = new Timer(Poll, null, 60, 60);
    }

    public string? Message { get { lock (_gate) { return _message; } } }

    private void Poll(object? state)
    {
        lock (_gate)
        {
            if (_closed) { return; }
            try
            {
                foreach (var dialog in DialogWatch.Dialogs())
                {
                    if (!DialogWatch.IsNotice(dialog) || !_seen.Add(dialog.Window)) { continue; }
                    _message = dialog.Text.Length > 0 ? dialog.Text : dialog.Caption;
                    if (DialogWatch.SafeAnswerFor(dialog) is { } button)
                    {
                        DialogWatch.DismissWindow(dialog.Window, button);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Info("immediate: native error notice could not be read: " + ex.Message);
            }
        }
    }

    public void Dispose()
    {
        lock (_gate) { _closed = true; }
        _timer.Dispose();
    }
}
