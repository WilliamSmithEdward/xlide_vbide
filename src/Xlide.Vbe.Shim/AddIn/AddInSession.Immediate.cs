using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// The Immediate window's one way in, for the panel and the door alike.
///
/// A LINE THAT STOPPED THE EDITOR CUT THE PAGE OFF FROM THE HOST (#30). The panel's line ran on
/// the host thread inside the browser's WebMessageReceived handler. `debug.prnt "A"` does not
/// compile, and once its box was answered VBA stayed stopped inside the scratch procedure, with
/// `Application.Run` suspended under that handler. WebView2 runs its callbacks one at a time and
/// never re-enters one; its threading-model page says a nested message loop inside an event
/// handler "would leave the event handler in the stack indefinitely". So nothing from the page
/// reached the host after that: Debug.Print printed nothing, Break and Reset did nothing, and a
/// script asked of the page never answered. The poll and the door went on working, because their
/// work reaches the host thread by our own window timers, which are not WebView2 callbacks.
/// Closing the editor asked to stop the debugger, and stopping it unwound the call and let the
/// backlog through, as the issue reports.
///
/// So the panel's line now runs the way the door's always has: posted to the host thread through
/// the timer queue, outside any WebView2 callback, and watched from a pool thread.
/// </summary>
internal sealed partial class AddInSession
{
    /// <summary>What an Immediate line came to: whether it finished, its text, and whether that text is a failure.</summary>
    internal readonly record struct ImmediateAnswer(bool Ran, string Text, bool Failed);

    /// <summary>Who typed a line, which decides who answers the boxes it raises.</summary>
    internal enum ImmediateCaller
    {
        /// <summary>
        /// A caller at the api door. Nobody is at the screen for it, so the boxes its line raises
        /// are answered here, and the wait is held to the request's budget.
        /// </summary>
        Door,

        /// <summary>
        /// The developer at the panel, who answers the boxes, as in the editor's own Immediate
        /// window: a box may be their own MsgBox asking them something. The wait lasts as long
        /// as the line does.
        /// </summary>
        Panel,
    }

    /// <summary>
    /// Handed to one evaluation so that its own result can be withheld: set when the outcome is
    /// put in the panel from outside - the box's own words - so the error the suspended call
    /// returns once a reset has unwound it is not said as well.
    /// </summary>
    internal sealed class ImmediateTicket
    {
        private volatile bool _answered;

        public bool Answered
        {
            get => _answered;
            set => _answered = value;
        }
    }

    /// <summary>Where the editor is stopped, read on the host thread for a pool thread.</summary>
    private enum EditorStop
    {
        NotStopped,
        InScratch,
        Elsewhere,
        Unanswered,
    }

    /// <summary>
    /// One Immediate line at a time from off the host thread. Two at once would have two loops
    /// answering the same dialogs and clearing the same stop.
    /// </summary>
    private readonly SemaphoreSlim _immediateAwayGate = new(1, 1);

    /// <summary>The panel's lines, each waiting for the one typed before it.</summary>
    private Task _panelLines = Task.CompletedTask;

    private readonly Lock _panelLinesGate = new();

    /// <summary>
    /// How long a stop has to hold still before the panel acts on it. The evaluator resets a line
    /// that returned out of design mode on its own, from the host thread, and a stop seen in that
    /// moment is on its way out; a line suspended in the break loop stays put.
    /// </summary>
    private const int PanelStopSettleMilliseconds = 400;

    /// <summary>
    /// A line entered in the panel, evaluated off the host thread in the order it was typed.
    ///
    /// IN THE ORDER TYPED. The handler this replaced ran each line to its end before the browser
    /// delivered the next, so `Range("A1") = 1` and then `? Range("A1")` could never swap. Pool
    /// threads promise no order and neither does the gate, so each line waits on the one before it.
    /// </summary>
    private void EvaluateFromPanel(string line)
    {
        lock (_panelLinesGate)
        {
            _panelLines = _panelLines.ContinueWith(
                _ => EvaluatePanelLine(line),
                CancellationToken.None,
                TaskContinuationOptions.None,
                TaskScheduler.Default);
        }
    }

    private void EvaluatePanelLine(string line)
    {
        try
        {
            EvaluateImmediateAway(line, ImmediateCaller.Panel, Timeout.Infinite);
        }
        catch (Exception ex)
        {
            Log.Error("immediate: the panel's line could not be evaluated", ex);
        }
    }

    /// <summary>
    /// Evaluates a line from OFF the host thread, and answers what it came to.
    ///
    /// Clears a stop this product left on the way in, runs the line on the host thread, and
    /// watches it from here. For the door, the boxes the line raises are answered and their words
    /// are the outcome. For the panel, the developer answers them, and a stop the line leaves in
    /// the scratch module is cleared once they have. Never touches COM from this thread: every
    /// question of the editor crosses to the host thread and comes back.
    /// </summary>
    internal ImmediateAnswer EvaluateImmediateAway(string text, ImmediateCaller caller, int budgetMilliseconds)
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return new ImmediateAnswer(false, "the surface is not up yet", true);
        }

        var entered = Environment.TickCount64;
        if (!_immediateAwayGate.Wait(budgetMilliseconds))
        {
            return new ImmediateAnswer(false, "Another Immediate line is still being evaluated.", true);
        }

        try
        {
            // One budget for the whole call: time spent waiting for the gate is time the caller
            // has already waited, so a door request is never answered after twice its `waitMs`.
            var remaining = budgetMilliseconds == Timeout.Infinite
                ? Timeout.Infinite
                : (int)Math.Max(0, budgetMilliseconds - (Environment.TickCount64 - entered));
            return EvaluateImmediateAwayHeld(surface, text, caller, remaining);
        }
        finally
        {
            _immediateAwayGate.Release();
        }
    }

    private ImmediateAnswer EvaluateImmediateAwayHeld(
        EditorSurface surface, string text, ImmediateCaller caller, int budgetMilliseconds)
    {
        /*
         * Started on the host thread and NOT waited on there, then answered from here, which is
         * the shape `compile` already uses and for the same reason.
         *
         * A line that will not compile raises the editor's own "Compile error" box. That box owns
         * the host thread, so anything waiting on that thread waits for the box, and the box is
         * waiting for somebody to press OK. This thread is the only one still moving, so it is the
         * one that notices it. The door's box is answered from here; the panel's is left for the
         * developer, and what this thread clears for them is the stop behind it.
         *
         * The first version of this waited on an event and reported a timeout. It made things
         * worse rather than better: the request returned after ten seconds, the dialog stopped
         * being one this request had raised, and nothing cleared it at all -- so a mistyped line
         * left a modal standing in front of the editor for the rest of the session instead of for
         * thirteen seconds.
         */
        /*
         * A SESSION THIS PRODUCT LEFT STOPPED IS CLEARED ON THE WAY IN, FROM THIS THREAD.
         *
         * THE ROOT CAUSE, which three attempts on the host thread could not reach. When a line
         * will not compile, the editor's "Compile error" box goes up and dismissing it leaves VBA
         * stopped INSIDE the scratch procedure, with `Application.Run` suspended mid-call. A
         * suspended frame unwinds only when the host thread returns to its message loop -- so a
         * recovery running ON that thread, inside a RunOnHostThread callback, is holding the one
         * thing that has to happen for the recovery to work. Reset was issued, repeatedly, for
         * eight seconds, and could not take: not because the budget was short but because no
         * budget can be long enough when waiting is itself what prevents the wait from ending
         * (2026-08-07).
         *
         * Issued from here it is an ordinary request, the host thread goes back to its pump
         * between calls, the frame unwinds, and the mode is design again. The polling below is not
         * a timing guess for the same reason `compile` polls: this thread is the only one still
         * moving, and what it waits for can actually happen while it waits.
         *
         * TWO WAYS THIS PRODUCT LEAVES THE EDITOR STOPPED, and both are cleared here. The first
         * is a line that stopped inside the scratch module. The second is a COMPILE ERROR
         * elsewhere: evaluating anything in a project that will not compile makes the editor raise
         * its box and then drop out of design mode a moment AFTER the evaluation has returned -
         * measured at 40ms - so nothing on the way out can see it. What was left was an editor
         * answering "Not available while execution is stopped" to every later evaluation and
         * refusing every write, for ever, over one syntax error somebody was still typing.
         */
        bool StoppedByUs() =>
            ScratchBreakStanding() || (_immediateLeftItStopped && _inBreak);

        if (StoppedByUs())
        {
            Log.Info("immediate: the editor is stopped by something this product ran, clearing it");

            if (!ClearStandingBreak(surface, StoppedByUs))
            {
                _immediateLeftItStopped = false;
                var stuck = "The last line left the editor stopped and it could not be "
                    + "cleared. Press Reset in the editor, or POST command?name=reset.";

                Log.Warn($"immediate: {stuck}");
                surface.RunOnHostThread(() => surface.ShowImmediateResult(stuck, true));
                return new ImmediateAnswer(false, stuck, true);
            }

            _immediateLeftItStopped = false;
        }

        var raisedBefore = DialogWatch.Dialogs().Select(row => row.Window).ToHashSet(StringComparer.Ordinal);

        // An evaluation can outlive this wait, for example at a nested breakpoint. A late
        // completion must not signal a disposed ManualResetEventSlim.
        var evaluated = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var ticket = new ImmediateTicket();
        var preserveDebugSession = false;
        var started = false;
        var outcome = string.Empty;
        var failed = false;

        var queued = surface.RunOnHostThread(() =>
        {
            try
            {
                Volatile.Write(ref preserveDebugSession, ProjectModeNow() != DesignMode);

                // After the mode, so a watcher that sees the line started also sees whether it
                // started paused at the developer's own stop.
                Volatile.Write(ref started, true);
                var result = EvaluateImmediate(text, ticket);
                outcome = result.Text;
                failed = result.Failed;
            }
            finally
            {
                evaluated.TrySetResult();
            }
        });

        if (!queued)
        {
            return new ImmediateAnswer(false, "the editor's window is gone", true);
        }

        var complained = new List<string>();
        var raisedByLine = new HashSet<string>(StringComparer.Ordinal);
        var deadline = budgetMilliseconds == Timeout.Infinite
            ? long.MaxValue
            : Environment.TickCount64 + budgetMilliseconds;

        // Set when a dialog has been answered that the evaluation cannot come back from, so the
        // wait below stops rather than running out its budget - see the note at the break.
        var nothingLeftToWaitFor = false;

        // The panel's reading of where the editor is stopped, and since when.
        var stop = EditorStop.NotStopped;
        var stopSince = 0L;

        while (Environment.TickCount64 < deadline && !evaluated.Task.IsCompleted && !nothingLeftToWaitFor)
        {
            evaluated.Task.Wait(120);

            // Paused evaluation owns its notices and must never arm scratch recovery.
            if (Volatile.Read(ref preserveDebugSession)) { continue; }

            foreach (var raised in DialogWatch.Dialogs())
            {
                if (!raisedBefore.Add(raised.Window))
                {
                    continue;
                }

                // The box's own words ARE the answer. A compile error says what is wrong with the
                // line, which is exactly what the developer typed it to find out.
                complained.Add(raised.Text.Length > 0 ? raised.Text : raised.Caption);
                raisedByLine.Add(raised.Window);

                if (caller == ImmediateCaller.Panel)
                {
                    Log.Info($"immediate: \"{raised.Text}\" is up for the developer to answer");
                    continue;
                }

                var pressed = DialogWatch.SafeAnswerFor(raised) ?? "OK";
                DialogWatch.Dismiss(raised.Caption, pressed);
                Log.Info($"immediate: \"{raised.Text}\" answered with {pressed}");

                // MARKED HERE, not after the answer is composed, because the editor drops out of
                // design mode about forty milliseconds from now and the poll that notices runs
                // every hundred and fifty. Setting this at the end lost that race every time: the
                // poll saw the break, found the flag still false, and left it standing (measured
                // 2026-08-24).
                //
                // AND ONLY IF THE EVALUATION HAS NOT COMPLETED, which is the whole safety of it and
                // is knowable right here. A compile error is raised INSTEAD of running, so the line
                // never started and there is no session to lose. A dialog raised after the
                // evaluation completed belongs to code that DID run - `Debug.Print TheirFunction()`
                // failing inside their own procedure - and the stop that follows is theirs, with
                // their call stack on it. That is never cleared for them.
                if (!evaluated.Task.IsCompleted)
                {
                    _immediateLeftItStopped = true;

                    // The words go in the panel from here, so the error the suspended call
                    // returns once a reset unwinds it is withheld rather than said in their
                    // place.
                    ticket.Answered = true;

                    // AND STOP WAITING FOR A COMPLETION THAT CANNOT COME. A compile error is
                    // raised instead of running the line, so the evaluation never completes and
                    // this used to wait out the whole budget - 17.1 seconds, measured three times,
                    // for a developer who typed while their project had a syntax error somewhere.
                    // A short grace first, because dismissing the box is occasionally what lets
                    // the evaluation finish, and a result is better than a complaint.
                    evaluated.Task.Wait(500);
                    nothingLeftToWaitFor = true;
                    break;
                }
            }

            /*
             * THE PANEL WATCHES FOR THE STOP ITSELF, having answered nothing.
             *
             * The door's evidence that a stop is ours is the box it answered. The panel has none,
             * because the developer answers the box, so it reads where the editor stopped. In the
             * scratch module it is this line's own: a line that would not compile, or a `Stop` in
             * the line itself, and nothing of the developer's is below it on the stack. Anywhere
             * else it is theirs, at a breakpoint in code the line called, and it is left to them.
             *
             * Read only once the poll has published a break, and acted on only when the reading
             * has held for a moment with none of the line's boxes still up: the evaluator resets
             * a line that returned out of design mode on its own, and a second reset racing that
             * one is issue #6. The published mode and not `_inBreak`, which the poll never sets
             * for a stop inside the scratch module - the one stop this is looking for.
             */
            if (caller == ImmediateCaller.Panel
                && Volatile.Read(ref started)
                && Volatile.Read(ref _lastPublishedMode) == "break")
            {
                var now = EditorStopNow(surface);
                var boxUp = DialogWatch.Dialogs().Any(row => raisedByLine.Contains(row.Window));

                if (now != stop || boxUp || evaluated.Task.IsCompleted)
                {
                    stop = now;
                    stopSince = Environment.TickCount64;
                }
                else if (Environment.TickCount64 - stopSince >= PanelStopSettleMilliseconds)
                {
                    if (stop == EditorStop.InScratch)
                    {
                        return ClearPanelLineStop(surface, evaluated.Task, ticket, complained);
                    }

                    if (stop == EditorStop.Elsewhere)
                    {
                        Log.Info("immediate: the panel's line is stopped in the developer's own code; left to them");
                        return new ImmediateAnswer(false, string.Empty, false);
                    }
                }
            }
        }

        var ran = evaluated.Task.Wait(2000);

        if (caller == ImmediateCaller.Panel)
        {
            // What the line came to is in the panel already, put there by the evaluation itself.
            return new ImmediateAnswer(ran, outcome, failed);
        }

        // What the editor complained about outranks what the evaluator managed to return. A
        // cleared compile box leaves the run answering an empty string, which reads as a
        // successful evaluation of nothing.
        if (complained.Count > 0)
        {
            outcome = Joined(complained);
            failed = true;
        }

        // TAKEN BACK IF THE EVALUATION RAN. Running is what distinguishes the two stops: a compile
        // error means nothing of the developer's ever started, so the break is ours to clear, while
        // an evaluation that RAN and then stopped may be sitting at their own breakpoint, in their
        // own code, and that stop is theirs. Left standing when it did not run, for the poll to
        // clear at the break or the next line to clear on its way in.
        if (ran)
        {
            _immediateLeftItStopped = false;
        }

        if (ticket.Answered)
        {
            var said = outcome;
            surface.RunOnHostThread(() => surface.ShowImmediateResult(said, failed));
        }

        return new ImmediateAnswer(ran, outcome, failed);
    }

    /// <summary>
    /// Clears a stop the panel's line left in the scratch module: the box's words go in the panel,
    /// the ticket withholds the error the unwound call returns, and the stop is reset.
    /// </summary>
    private ImmediateAnswer ClearPanelLineStop(
        EditorSurface surface, Task evaluated, ImmediateTicket ticket, List<string> complained)
    {
        ticket.Answered = true;
        var said = complained.Count > 0
            ? Joined(complained)
            : "The line stopped the editor, so it was reset.";

        surface.RunOnHostThread(() => surface.ShowImmediateResult(said, true));
        Log.Info("immediate: the panel's line left the editor stopped in the scratch module, clearing it");

        if (!ClearStandingBreak(surface, () => !evaluated.IsCompleted || ScratchBreakStanding()))
        {
            Log.Warn("immediate: the panel's line left a stop that would not clear; press Reset in the editor");
        }

        return new ImmediateAnswer(false, said, true);
    }

    /// <summary>
    /// A line's box words as one line for the panel. The box sets its words on lines of their
    /// own with a blank one between, which came out as "Compile error:  Syntax error".
    /// </summary>
    private static string Joined(List<string> complained) =>
        string.Join(' ', string.Join(' ', complained).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    /// <summary>
    /// Where the editor is stopped, asked from off the host thread. Unanswered when the host
    /// thread is busy for two seconds, which the caller treats as nothing learned.
    /// </summary>
    private EditorStop EditorStopNow(EditorSurface surface)
    {
        var answer = new TaskCompletionSource<EditorStop>(TaskCreationOptions.RunContinuationsAsynchronously);
        var queued = surface.RunOnHostThread(() =>
        {
            try
            {
                answer.TrySetResult(ProjectModeNow() != BreakMode
                    ? EditorStop.NotStopped
                    : StoppedInScratchModule() ? EditorStop.InScratch : EditorStop.Elsewhere);
            }
            catch (Exception ex)
            {
                Log.Info($"immediate: could not read where the editor is stopped ({ex.GetType().Name})");
                answer.TrySetResult(EditorStop.Unanswered);
            }
        });

        return queued && answer.Task.Wait(2000) ? answer.Task.Result : EditorStop.Unanswered;
    }

    /// <summary>
    /// Resets a stop this product caused, from off the host thread, and answers whether it cleared.
    /// Reset asks "proceed anyway?", and the rescue that answers a dialog blocking the host thread
    /// declines every real question - right for a question nobody here asked, and fatal for this
    /// one (issue #6). So the confirmation is expected for as long as the wait lasts, and the reset
    /// is asked once.
    /// </summary>
    private bool ClearStandingBreak(EditorSurface surface, Func<bool> stillStopped)
    {
        using (DialogWatch.ExpectingConfirmation(6000))
        {
            surface.RunOnHostThread(() => ExecuteEditorCommand(VbeCommands.Command.Reset));

            var clearBy = Environment.TickCount64 + 5000;
            while (Environment.TickCount64 < clearBy && stillStopped())
            {
                Thread.Sleep(100);
            }
        }

        if (stillStopped())
        {
            return false;
        }

        // A reset does not take the scratch module away, and one standing would be compiled with
        // the project, appear in the explorer, and be saved into the workbook.
        surface.RunOnHostThread(RemoveScratchModule);
        return true;
    }
}
