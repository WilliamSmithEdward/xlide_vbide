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
    /// would pick instead.
    /// </summary>
    private const string ScratchModule = "XlideImmediateScratch";

    /// <summary>Procedure the module exposes, which is what gets run by name.</summary>
    private const string ScratchProcedure = "XlideImmediateRun";

    /// <summary>Standard module. The editor's own numbering.</summary>
    private const int StandardModule = 1;

    private readonly DispatchObject _editor;

    public ImmediateEvaluator(DispatchObject editor) => _editor = editor;

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

            // Run by name rather than through the module, because the project is what owns the
            // compiled procedure and the host is what knows how to call into it.
            var value = application.CallToString("Run", $"{ScratchModule}.{ScratchProcedure}");

            return wantsValue ? new Result(value, Failed: false) : new Result("done", Failed: false);
        }
        finally
        {
            // Always. A scratch module left behind would be compiled with the project, would appear
            // in the explorer, and would be saved into the workbook.
            Remove(components);
        }
    }

    /// <summary>
    /// Wraps a line in something the project can compile.
    ///
    /// A value is returned as a variant, so anything the expression produces survives the trip; a
    /// statement is run for its effect and reports nothing.
    /// </summary>
    private static string Compose(string body, bool wantsValue) => wantsValue
        ? $"Public Function {ScratchProcedure}() As Variant\r\n"
          + $"    {ScratchProcedure} = ({body})\r\n"
          + "End Function\r\n"
        : $"Public Sub {ScratchProcedure}()\r\n"
          + $"    {body}\r\n"
          + "End Sub\r\n";

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
