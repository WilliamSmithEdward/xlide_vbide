using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// Evaluates what the developer types in the Immediate panel.
///
/// VBA has no way to evaluate a string as code. Its own Immediate window is a compiler front end
/// wired directly into the interpreter, and none of that is exposed: the window reports no handle
/// through the object model, and asking the window itself for its text returns its caption, so it
/// can be neither driven nor read from outside.
///
/// What is exposed is the ability to add a procedure to the project and run it by name. So a line
/// is compiled by writing it into a module of its own, running it, and taking the module away
/// again. The module's name is not one anybody would choose, so nothing of the developer's is at
/// risk of being replaced by it, and it is removed whether the line succeeded or not.
///
/// This is not available while execution is stopped. Adding a module resets the project, which
/// would end the debugging session the developer is in the middle of, and losing that is a far
/// worse answer than declining.
/// </summary>
internal sealed class ImmediateEvaluator
{
    /// <summary>
    /// Name of the module a line is compiled into. Removed after every evaluation.
    ///
    /// A VBA identifier has to start with a letter, so this cannot be given the leading underscores
    /// that would otherwise mark it as not the developer's. The name is chosen to be one nobody
    /// would pick instead. Internal because the session must treat this module as invisible: it
    /// briefly exists, and nothing of it may reach a tab, the explorer, or the editor.
    /// </summary>
    internal const string ScratchModule = "XlideImmediateScratch";

    /// <summary>Procedure the module exposes, which is what gets run by name.</summary>
    private const string ScratchProcedure = "XlideImmediateRun";

    /// <summary>
    /// Marks a value that is an error description rather than a result. Control characters cannot
    /// be typed into the panel, so no expression of the developer's can produce this by accident.
    /// </summary>
    private const char ErrorMarker = (char)1;

    /// <summary>Standard module. The editor's own numbering.</summary>
    private const int StandardModule = 1;

    /// <summary>The editor's own numbering again: design, run, break.</summary>
    private const int DesignMode = 2;

    private readonly DispatchObject _editor;

    /// <summary>
    /// Set when an evaluation of ours left the project out of design mode, so the NEXT one can
    /// put it back before deciding it is looking at a debugging session.
    ///
    /// The recovery has to be armed rather than inferred from the mode alone: a developer stopped
    /// at their own breakpoint is also out of design mode, and resetting that would throw away
    /// the session they are in the middle of.
    /// </summary>
    private bool _leftItRunning;

    public ImmediateEvaluator(DispatchObject editor) => _editor = editor;

    /// <summary>
    /// Whether running the line took the project out of design mode.
    ///
    /// True is the failure state, not a normal one. An expression with its own error handler
    /// returns cleanly and the project never moves; a line that will not compile puts the editor's
    /// "Compile error" box up and leaves the project stopped behind it. Unreadable answers false,
    /// deliberately: a reset that cannot be justified is worse than one that is skipped.
    /// </summary>
    private static bool LeftDesignMode(DispatchObject? project)
    {
        try
        {
            var mode = project?.GetInt32("Mode") ?? DesignMode;
            Log.Info($"immediate: project mode is {mode} ({(mode == DesignMode ? "design" : "NOT design")})");
            return mode != DesignMode;
        }
        catch (Exception ex)
        {
            Log.Info($"immediate: the project mode could not be read ({ex.GetType().Name}: {ex.Message})");
            return false;
        }
    }

    /// <summary>
    /// Called when an evaluation has left the project out of design mode, so the caller can put
    /// it back.
    ///
    /// The reset itself is NOT done here. Stopping a project is the editor's own Reset command,
    /// which the session already knows how to execute and which this type has no business
    /// reproducing with COM calls of its own invention. This type's job is to notice.
    /// </summary>
    public Action? StoppedUnexpectedly { get; set; }

    /// <summary>The outcome of one line: what to show, and whether it went wrong.</summary>
    public readonly record struct Result(string Text, bool Failed);

    /// <summary>
    /// Runs one line.
    ///
    /// A line beginning with a question mark is an expression whose value is wanted, which is the
    /// convention the editor's own Immediate window uses and the one every VBA developer already
    /// has in their fingers. Anything else is a statement, run for its effect.
    /// </summary>
    public Result Evaluate(string line, bool inBreakMode)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(line);

        /*
         * RECOVER FROM OUR OWN MESS FIRST, before deciding this is a debugging session.
         *
         * A line that will not compile raises the editor's "Compile error" box, which owns the
         * host thread until the dialog guard clears it, and leaves the project in RUN mode behind
         * it. From then on every evaluation answered "Not available while execution is stopped",
         * so one mistyped line made the Immediate window useless until somebody thought to press
         * Reset, and the message blamed a debugging session the developer had never started
         * (measured 2026-08-07 with `?((`).
         *
         * The check cannot simply be "not design mode": that is also what a developer stopped at
         * their own breakpoint looks like, and resetting THAT would throw away the session they
         * are in the middle of, which is the one thing worse than declining. So the recovery is
         * armed only when this type is the one that left the project running, and it is
         * disarmed the moment it has been used.
         */
        if (_leftItRunning)
        {
            _leftItRunning = false;
            StoppedUnexpectedly?.Invoke();

            // The mode is read AGAIN, because the caller read it before this recovery ran. Using
            // its answer would decline the very line the recovery was performed for, and the
            // developer would have to type it a second time to find out it works now.
            using var project = _editor.GetObject("ActiveVBProject");
            inBreakMode = LeftDesignMode(project);
        }

        if (inBreakMode)
        {
            return new Result(
                "Not available while execution is stopped: evaluating adds a procedure to the "
                + "project, which would reset it and end the debugging session.",
                Failed: true);
        }

        var text = line.Trim();
        var wantsValue = text.StartsWith('?');
        var body = wantsValue ? text[1..].Trim() : text;

        if (body.Length == 0)
        {
            return new Result("Nothing to evaluate.", Failed: true);
        }

        try
        {
            return Run(body, wantsValue);
        }
        catch (Exception ex)
        {
            // The message is the developer's answer, not a diagnostic: a mistyped expression is an
            // ordinary outcome here and belongs in the panel rather than in a log nobody reads.
            Log.Info($"immediate: {body} failed, {ex.Message}");
            return new Result(Describe(ex), Failed: true);
        }
    }

    private Result Run(string body, bool wantsValue)
    {
        using var project = _editor.GetObject("ActiveVBProject");
        using var components = project?.GetObject("VBComponents");
        if (components is null)
        {
            return new Result("No project is open.", Failed: true);
        }

        Remove(components);

        using var component = components.CallObject("Add", StandardModule);
        if (component is null)
        {
            return new Result("The project would not accept a scratch module.", Failed: true);
        }

        try
        {
            component.SetString("Name", ScratchModule);

            using var module = component.GetObject("CodeModule");
            module?.Invoke("AddFromString", Compose(body, wantsValue));

            using var application = HostApplication.Find();
            if (application is null)
            {
                return new Result("The host application could not be reached.", Failed: true);
            }

            /*
             * QUALIFIED BY WORKBOOK, because "current" means two different things here.
             *
             * The scratch module is added to the editor's ACTIVE VB PROJECT; `Application.Run`
             * resolves an unqualified name against the host's ACTIVE WORKBOOK. With one workbook
             * open those are always the same and the difference cannot be seen. With two they can
             * differ, and the evaluation fails with the host's own words: "Cannot run the macro
             * 'XlideImmediateScratch.XlideImmediateRun'. The macro may not be available in this
             * workbook or all macros may be disabled." Which is true, and unhelpful: it exists, in
             * the other workbook (2026-08-07).
             *
             * A project with no file yet has no name to qualify with, and there the unqualified
             * form is all there is.
             */
            var file = SafeFileName(project);
            var target = file is null
                ? $"{ScratchModule}.{ScratchProcedure}"
                : $"'{file}'!{ScratchModule}.{ScratchProcedure}";

            var value = application.CallToString("Run", target);

            // The scratch code marks an error it caught, and the marker cannot be typed.
            if (value.Length > 0 && value[0] == ErrorMarker)
            {
                return new Result(value[1..], Failed: true);
            }

            /*
             * A LINE THAT WILL NOT COMPILE never reaches the error handler above.
             *
             * `On Error GoTo` catches run-time errors, and a syntax error is not one: the project
             * never compiles, so the handler is never installed and the editor puts up its own
             * "Compile error" box instead. Measured 2026-08-07 with `?((`: the box owned the host
             * thread for THIRTEEN SECONDS until the dialog guard cleared it, `Run` then returned
             * an empty string, and this reported a successful evaluation of nothing.
             *
             * The lasting half was worse. The project was left OUT of design mode, and every
             * evaluation after it answered "Not available while execution is stopped" -- so one
             * mistyped line made the Immediate window useless until somebody thought to press
             * Reset, and the message blamed a debugging session the developer had never started.
             *
             * So the mode is checked rather than assumed. Leaving design mode without the handler
             * having answered means the line did not compile, which is reported as the failure it
             * is, and the project is put back so the next line works.
             */
            if (LeftDesignMode(project))
            {
                _leftItRunning = true;
                StoppedUnexpectedly?.Invoke();
                return new Result("The line could not be compiled.", Failed: true);
            }

            // A statement that succeeded says nothing, which is what the editor's own Immediate
            // window does: output belongs to the code (Debug.Print), not to the ceremony.
            return wantsValue ? new Result(value, Failed: false) : new Result(string.Empty, Failed: false);
        }
        finally
        {
            /*
             * THE MODE IS READ FIRST, BEFORE THE CLEAN-UP, because the clean-up is the thing most
             * likely to throw when the mode is wrong.
             *
             * A project stopped inside the scratch procedure will not let its module be removed,
             * so `Remove` throws, and anything after it in this block never runs. The arming used
             * to sit after it, so the one case it existed for -- the project left stopped -- was
             * the one case it never fired in (2026-08-07).
             */
            if (LeftDesignMode(project))
            {
                // Armed for the NEXT evaluation, which resets before deciding it is looking at a
                // debugging session, and reported now as well so the reset can start rather than
                // waiting for somebody to type again.
                _leftItRunning = true;
                StoppedUnexpectedly?.Invoke();
            }

            // Always. A scratch module left behind would be compiled with the project, would appear
            // in the explorer, and would be saved into the workbook. Attempted even when the
            // project was stopped, since the reset above may have already freed it, and a failure
            // here must not replace the answer the developer was waiting for.
            try
            {
                Remove(components);
            }
            catch (Exception ex)
            {
                Log.Info($"immediate: the scratch module could not be removed ({ex.GetType().Name})");
            }
        }
    }

    /// <summary>
    /// Wraps a line in something the project can compile.
    ///
    /// Always a function, and always with its own error handler. A run-time error in an unhandled
    /// frame is the editor's cue to put up its error dialog and drop into break mode INSIDE the
    /// scratch module, which put the scratch code on the developer's screen with a stopped-line
    /// marker on it. Handled here, the error comes back as a marked value instead, and the panel
    /// shows the language's own message the way it shows any other answer.
    /// </summary>
    /// <summary>
    /// The workbook file name a project belongs to, or null when it has never been saved.
    ///
    /// Only the name, not the path: that is what `Application.Run` wants between its quotes, and
    /// a path there fails to resolve.
    /// </summary>
    private static string? SafeFileName(DispatchObject? project)
    {
        try
        {
            var full = project?.GetString("FileName");
            return string.IsNullOrEmpty(full) ? null : Path.GetFileName(full);
        }
        catch (Exception)
        {
            // An unsaved project throws rather than answering empty. The unqualified name stands.
            return null;
        }
    }

    private static string Compose(string body, bool wantsValue)
    {
        var work = wantsValue ? $"    {ScratchProcedure} = ({body})\r\n" : $"    {body}\r\n";

        return $"Public Function {ScratchProcedure}() As Variant\r\n"
            + "    On Error GoTo Failed\r\n"
            + work
            + "    Exit Function\r\n"
            + "Failed:\r\n"
            + $"    {ScratchProcedure} = Chr$(1) & Err.Description & \" (error \" & Err.Number & \")\"\r\n"
            + "End Function\r\n";
    }

    private static void Remove(DispatchObject components)
    {
        try
        {
            var count = components.GetInt32("Count");
            for (var i = count; i >= 1; i--)
            {
                using var candidate = components.GetItem(i);
                if (candidate?.GetString("Name") == ScratchModule)
                {
                    components.InvokeWithObject("Remove", candidate);
                }
            }
        }
        catch (Exception ex)
        {
            Log.Warn($"immediate: the scratch module could not be removed, {ex.Message}");
        }
    }

    /// <summary>
    /// Turns an automation failure into something worth reading.
    ///
    /// A compile error arrives as a failed call with the editor's own message inside it, which is
    /// exactly what the developer needs; the wrapper around it is not.
    /// </summary>
    private static string Describe(Exception ex)
    {
        var message = ex.Message.Trim();
        return message.Length == 0 ? ex.GetType().Name : message;
    }
}
