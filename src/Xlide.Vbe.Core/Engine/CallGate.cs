namespace Xlide.Vbe.Core.Engine;

/// <summary>
/// What a caller is putting on the engine's single pipe, which is all this gate needs to know to
/// order it.
/// </summary>
public enum CallKind
{
    /// <summary>
    /// Changes what the engine HOLDS, or must be seen in the order it was sent: every
    /// notification, the seed, the handshake, a close. Nothing may go in front of one of these,
    /// or a query answers about text or a project the engine has not been given yet.
    /// </summary>
    Barrier,

    /// <summary>
    /// A read-only query nobody is waiting on: the squiggles and the colouring. Correct whenever
    /// it lands, which is what makes it the thing to move.
    /// </summary>
    Background,

    /// <summary>
    /// A read-only query a person is waiting for with their hands on the keyboard: the
    /// completion list, the tooltip, the call tip, the re-casing of what they just typed.
    /// </summary>
    Interactive,
}

/// <summary>
/// The turn-taking on the engine's pipe: one call at a time, in arrival order, EXCEPT that a
/// query somebody is waiting for may go in front of background work that is only queued.
///
/// WHY THIS EXISTS. The engine serves one request at a time, and a whole-module diagnostics or
/// colouring pass over a large file is hundreds of milliseconds of it. A completion asked for
/// straight after a keystroke arrives while that pass is queued or running, and used to wait for
/// all of it. Measured on a 64,802-line module (2026-08-21): a completion cost 177ms of analyzer
/// and 324ms of QUEUE, worst 1083ms; the re-casing of what had just been typed cost 39ms of
/// analyzer and 554ms of queue, worst 1192ms - so the case correction landed more than a second
/// after the word it was correcting.
///
/// WHAT IT WILL NOT DO. Reordering two read-only queries is safe; reordering anything across a
/// BARRIER is not, because the engine's answers are about the text it has been told about. So an
/// interactive query moves ahead only of background work that is queued in front of it with no
/// barrier in between. It never interrupts the call already running - stopping that needs
/// cancellation the pipe cannot correlate - and it never starves the work it passes: a
/// background waiter passed over <see cref="Patience"/> times goes next whatever arrives.
/// </summary>
public sealed class CallGate
{
    /// <summary>
    /// How many times one piece of background work may be passed over before it goes next
    /// regardless. Each pass-over costs it roughly one interactive call, so this is the bound on
    /// how far behind the squiggles may fall during a burst of typing - not a preference.
    /// </summary>
    public const int Patience = 4;

    private readonly object _turn = new();
    private readonly LinkedList<Waiter> _waiting = new();
    private bool _busy;

    /// <summary>Whoever holds the pipe now, plus everyone queued: for tests and diagnostics.</summary>
    public int Waiting
    {
        get
        {
            lock (_turn)
            {
                return _waiting.Count;
            }
        }
    }

    /// <summary>
    /// Takes the pipe, waiting for it if somebody has it. The task completes when the turn is
    /// this caller's; whoever awaits it owes a <see cref="Leave"/>, in a finally.
    /// </summary>
    public Task EnterAsync(CallKind kind, CancellationToken cancellation = default)
    {
        if (cancellation.IsCancellationRequested)
        {
            return Task.FromCanceled(cancellation);
        }

        LinkedListNode<Waiter> node;
        lock (_turn)
        {
            if (!_busy)
            {
                _busy = true;
                return Task.CompletedTask;
            }

            node = _waiting.AddLast(new Waiter(kind));
        }

        if (!cancellation.CanBeCanceled)
        {
            return node.Value.Ready.Task;
        }

        // A caller that gives up is taken OUT of the queue, and only then told: cancelling one
        // that has already been handed the turn would drop the turn on the floor and hang the
        // pipe for the rest of the session.
        var registration = cancellation.Register(() =>
        {
            var dropped = false;
            lock (_turn)
            {
                if (node.List is not null)
                {
                    _waiting.Remove(node);
                    dropped = true;
                }
            }

            if (dropped)
            {
                node.Value.Ready.TrySetCanceled(cancellation);
            }
        });

        return Settle(node.Value.Ready.Task, registration);

        static async Task Settle(Task ready, CancellationTokenRegistration registration)
        {
            try
            {
                await ready.ConfigureAwait(false);
            }
            finally
            {
                await registration.DisposeAsync().ConfigureAwait(false);
            }
        }
    }

    /// <summary>Gives the pipe to whoever is next, or leaves it free when nobody is.</summary>
    public void Leave()
    {
        Waiter? next = null;
        lock (_turn)
        {
            var chosen = Next();
            if (chosen is null)
            {
                _busy = false;
            }
            else
            {
                next = chosen.Value;
                _waiting.Remove(chosen);
            }
        }

        next?.Ready.TrySetResult();
    }

    /// <summary>
    /// Whose turn it is. The head, unless the head is background work that an interactive query
    /// behind it may pass - which it may only when nothing between them is a barrier, and only
    /// while every piece of background work it passes still has patience left.
    /// </summary>
    private LinkedListNode<Waiter>? Next()
    {
        var head = _waiting.First;
        if (head is null || head.Value.Kind != CallKind.Background)
        {
            return head;
        }

        for (var node = head; node is not null; node = node.Next)
        {
            if (node.Value.Kind == CallKind.Barrier)
            {
                break;
            }

            if (node.Value.Kind == CallKind.Interactive)
            {
                for (var passed = head; passed is not null && passed != node; passed = passed.Next)
                {
                    passed.Value.PassedOver++;
                }

                return node;
            }

            if (node.Value.PassedOver >= Patience)
            {
                break;
            }
        }

        return head;
    }

    private sealed class Waiter(CallKind kind)
    {
        public CallKind Kind { get; } = kind;

        /// <summary>How many interactive queries have gone in front of this one.</summary>
        public int PassedOver { get; set; }

        public TaskCompletionSource Ready { get; } =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
    }
}
