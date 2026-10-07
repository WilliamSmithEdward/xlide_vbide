using System.Text.RegularExpressions;
using Xlide.Vbe.Shim.Editor;
using Xlide.Vbe.Shim.Engine;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// Expands the native Locals tree and fills the one gap its accessibility provider leaves:
/// VBE paints user-class fields and properties but omits those rows from UI Automation.
/// Only a user click asks for class values; the native reader remains the source for everything
/// else. Public fields and no-argument getters are the members that can be evaluated safely by
/// name in the paused frame.
/// </summary>
internal sealed partial class AddInSession
{
    private readonly record struct LocalKey(string? Context, int Index, string Expression, int Depth);
    private readonly record struct ClassMember(string Name, string Kind);

    private readonly object _classLocalLock = new();
    private readonly Dictionary<LocalKey, SurfaceLocalRow[]> _classLocalRows = [];
    private readonly Dictionary<LocalKey, int> _classLocalTokens = [];
    private int _classLocalSerial;
    private int _classLocalEpoch;

    private void ToggleLocalFromPanel(int index, string expression, int depth, string? context)
    {
        var snapshot = _ghostReaders?.Locals;
        if (snapshot is null || index < 0 || index >= snapshot.Rows.Count
            || snapshot.Context != context || snapshot.Rows[index].Expression != expression
            || snapshot.Rows[index].Depth != depth)
        {
            return;
        }

        var key = new LocalKey(context, index, expression, depth);
        bool collapsed;
        lock (_classLocalLock)
        {
            collapsed = _classLocalRows.Remove(key);
            if (collapsed) _classLocalTokens.Remove(key);
        }
        if (collapsed)
        {
            PublishLocals(stopped: true);
            return;
        }

        var members = depth == 0 && Regex.IsMatch(expression, @"^[A-Za-z_][A-Za-z0-9_]*$")
            ? PublicClassMembers(snapshot.Rows[index].Type) : [];
        if (members.Length == 0)
        {
            _ghostReaders?.ToggleLocal(index, expression, depth, context);
            return;
        }

        var epoch = _classLocalEpoch;
        int token;
        lock (_classLocalLock)
        {
            token = ++_classLocalSerial;
            _classLocalTokens[key] = token;
            _classLocalRows[key] = [new SurfaceLocalRow("Loading members…", "", "", depth + 1)];
        }
        PublishLocals(stopped: true);

        _ = Task.Run(() =>
        {
            var rows = new List<SurfaceLocalRow>(members.Length);
            var started = Environment.TickCount64;
            Interlocked.Increment(ref _silentImmediateReads);
            try
            {
                foreach (var member in members)
                {
                    lock (_classLocalLock)
                    {
                        if (_classLocalEpoch != epoch || !_classLocalTokens.TryGetValue(key, out var current)
                            || current != token) return;
                    }
                    if (Environment.TickCount64 - started > 15_000)
                    {
                        rows.Add(new SurfaceLocalRow("More members", "Not read (time limit)", "", depth + 1));
                        break;
                    }
                    var answer = EvaluateImmediateAway($"? {expression}.{member.Name}", ImmediateCaller.Locals, 3000);
                    var value = answer.Ran && !answer.Failed ? answer.Text : "<unavailable>";
                    rows.Add(new SurfaceLocalRow(member.Name, value, member.Kind, depth + 1));
                }
            }
            catch
            {
                rows.Add(new SurfaceLocalRow("Members", "<unavailable>", "", depth + 1));
            }
            finally
            {
                // The native Immediate reader is asynchronous. Give it one poll to consume the
                // just-evaluated output before ordinary panel output is shown again.
                Thread.Sleep(500);
                Interlocked.Decrement(ref _silentImmediateReads);
            }

            _editorSurface?.RunOnHostThread(() =>
            {
                if (!_inBreak || _classLocalEpoch != epoch) return;
                lock (_classLocalLock)
                {
                    if (!_classLocalTokens.TryGetValue(key, out var current) || current != token) return;
                    _classLocalRows[key] = rows.ToArray();
                }
                PublishLocals(stopped: true);
            });
        });
    }

    private ClassMember[] PublicClassMembers(string kind)
    {
        var names = kind.Split('/');
        if (names.Length != 2) return [];

        var members = new List<ClassMember>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // The first half is the declared interface; querying members found only on the
        // runtime class would fail to compile for an interface-typed local.
        foreach (var className in names.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            using var component = FindComponent(className, _shownProject, out _);
            if (component is null || component.GetInt32("Type") != ClassModuleType) continue;
            var source = ProjectReader.ReadSource(component) ?? string.Empty;

            foreach (Match match in Regex.Matches(source,
                @"(?im)^\s*(?:Public\s+)?Property\s+Get\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(\s*\)\s*(?:As\s+([A-Za-z_][A-Za-z0-9_.]*))?"))
            {
                if (seen.Add(match.Groups[1].Value))
                    members.Add(new ClassMember(match.Groups[1].Value,
                        match.Groups[2].Success ? match.Groups[2].Value : "Variant"));
            }

            foreach (Match match in Regex.Matches(source,
                @"(?im)^\s*Public\s+([A-Za-z_][A-Za-z0-9_]*)\s+As\s+([A-Za-z_][A-Za-z0-9_.]*)"))
            {
                if (seen.Add(match.Groups[1].Value))
                    members.Add(new ClassMember(match.Groups[1].Value, match.Groups[2].Value));
            }
            break;
        }

        // An object model can expose hundreds of getters, some expensive or stateful. Keep an
        // explicit click bounded while still covering ordinary user-defined classes.
        return members.Take(24).ToArray();
    }

    private SurfaceLocalRow[] SurfaceLocals(LocalsReader.LocalsSnapshot snapshot)
    {
        var rows = new List<SurfaceLocalRow>(snapshot.Rows.Count);
        lock (_classLocalLock)
        {
            for (var i = 0; i < snapshot.Rows.Count; i++)
            {
                var native = snapshot.Rows[i];
                var key = new LocalKey(snapshot.Context, i, native.Expression, native.Depth);
                var expanded = _classLocalRows.TryGetValue(key, out var members);
                rows.Add(new SurfaceLocalRow(native.Expression, native.Value, native.Type,
                    native.Depth, native.Expandable, expanded || native.Expanded, i));
                if (expanded) rows.AddRange(members!);
            }
        }
        return rows.ToArray();
    }

    private void ClearClassLocalExpansions()
    {
        lock (_classLocalLock)
        {
            _classLocalEpoch++;
            _classLocalRows.Clear();
            _classLocalTokens.Clear();
        }
    }
}
