using System.Globalization;
using System.Text;

namespace Xlide.Vbe.Core.Forms;

/// <summary>
/// The form markup: a UserForm's design as text, in a dialect of xlide's own.
///
/// This is a PROJECTION, never a source of truth. The MSForms designer model keeps writing
/// `.frm` and `.frx`; the markup is generated from a walk of that model and applied back to it
/// as a name-keyed diff, so a developer without xlide sees a perfectly ordinary form and a
/// developer with it gets a text surface - diffable, undoable, searchable - over the same
/// design (docs/userform-designer.md, the markup layer).
///
/// The dialect is VBA's on purpose: apostrophe comments, True/False, doubled quotes inside
/// strings, &amp;H hex where a colour is spelled. A header line is
/// `Type Name ["Caption"] [at left,top] [size width x height]`; anything else about a control
/// is an indented `Path = value` line. Indentation is containment, four spaces per level.
///
/// Everything here is text in, records out, and back - no COM, no host - which is what makes
/// the language testable without Excel and shared verbatim between the generate and apply
/// sides. The control list is flat with parent NAMES rather than nested records, because that
/// is the shape the designer walk answers and the shape the diff wants.
/// </summary>
public static class FormMarkup
{
    // FOUR, the owner's call (2026-08-13): two-space levels read too shallow in a document
    // whose whole structure is its indentation. The parser refuses other multiples, so the
    // printer and every hand-written document agree about what a level is.
    private const int IndentWidth = 4;

    /// <summary>
    /// The kinds a MultiPage's children must be, and the kinds that may hold children at all.
    /// Case-insensitive membership, because the dialect is VBA's.
    /// </summary>
    private static readonly HashSet<string> Containers =
        new(StringComparer.OrdinalIgnoreCase) { "Frame", "MultiPage", "Page", "TabStrip" };

    // ------------------------------------------------------------------ printing

    /// <summary>
    /// The canonical text of a form. Deterministic - model order, fixed formatting - so a
    /// regeneration diffs cleanly against the last one.
    /// </summary>
    public static string Print(FormSpec form)
    {
        ArgumentNullException.ThrowIfNull(form);

        var text = new StringBuilder();
        text.Append("<Form Name=\"").Append(form.Name).Append('"');
        if (form.Caption is not null)
        {
            text.Append(" Caption=").Append(Quoted(form.Caption));
        }

        if (form.Width is { } width && form.Height is { } height)
        {
            text.Append(" Width=").Append(Number(width)).Append(" Height=").Append(Number(height));
        }

        foreach (var property in form.Properties)
        {
            AppendAttribute(text, property);
        }

        text.AppendLine(">");

        // Children under their container, containers in list order. Parent naming the form (or
        // nothing) means top level; the designer walk spells top level with the form's name.
        var byParent = new Dictionary<string, List<ControlSpec>>(StringComparer.OrdinalIgnoreCase);
        var roots = new List<ControlSpec>();
        foreach (var control in form.Controls)
        {
            if (control.Parent is null || string.Equals(control.Parent, form.Name, StringComparison.OrdinalIgnoreCase))
            {
                roots.Add(control);
            }
            else
            {
                if (!byParent.TryGetValue(control.Parent, out var siblings))
                {
                    siblings = [];
                    byParent[control.Parent] = siblings;
                }

                siblings.Add(control);
            }
        }

        foreach (var control in roots)
        {
            AppendControl(text, 1, control, byParent);
        }

        text.AppendLine("</Form>");
        return text.ToString();
    }

    /// <summary>
    /// One control as an element: everything it HAS is an attribute, everything it CONTAINS is a
    /// child element. That separation is the whole reason for the tags - under the old indented
    /// dialect a container's own property and a control sitting inside it were both one level in,
    /// and only the `=` told them apart.
    ///
    /// A control with no children closes itself, which keeps the common line to one line.
    /// </summary>
    private static void AppendControl(
        StringBuilder text, int depth, ControlSpec control, Dictionary<string, List<ControlSpec>> byParent)
    {
        text.Append(' ', depth * IndentWidth)
            .Append('<').Append(control.Type).Append(" Name=\"").Append(control.Name).Append('"');
        if (control.Caption is not null)
        {
            text.Append(" Caption=").Append(Quoted(control.Caption));
        }

        if (control.Left is { } left && control.Top is { } top)
        {
            text.Append(" Left=").Append(Number(left)).Append(" Top=").Append(Number(top));
        }

        if (control.Width is { } width && control.Height is { } height)
        {
            text.Append(" Width=").Append(Number(width)).Append(" Height=").Append(Number(height));
        }

        foreach (var property in control.Properties)
        {
            AppendAttribute(text, property);
        }

        var children = byParent.TryGetValue(control.Name, out var found) ? found : null;
        if (children is null || children.Count == 0)
        {
            text.AppendLine(" />");
            return;
        }

        text.AppendLine(">");
        foreach (var child in children)
        {
            AppendControl(text, depth + 1, child, byParent);
        }

        text.Append(' ', depth * IndentWidth).Append("</").Append(control.Type).AppendLine(">");
    }

    /// <summary>
    /// A property as an attribute. TEXT IS QUOTED AND EVERYTHING ELSE IS BARE, which is the
    /// owner's rule (2026-08-17) and the only spelling that survives a round trip: quoting a
    /// number the way XML does would make `Caption="True"` and a real flag identical on the way
    /// back, and the parser recovers the kind from the spelling rather than from a schema it does
    /// not have.
    /// </summary>
    private static void AppendAttribute(StringBuilder text, PropertySpec property)
    {
        text.Append(' ').Append(property.Path).Append('=').Append(property.Kind switch
        {
            PropertyValueKind.Text => Quoted(property.Value),
            PropertyValueKind.Colour when int.TryParse(
                property.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var ole)
                => SpellColour(ole),
            _ => property.Value,
        });
    }

    /// <summary>
    /// A colour as the document spells one: `#rrggbb`, or the VBA hex where there is no honest
    /// `#` to give. A SYSTEM colour is not an RGB at all but a question - what does this machine
    /// call a button face - so it keeps `&amp;H8000000F&amp;`, and a document that spelled it as
    /// today's answer would freeze that answer into the form.
    ///
    /// The document and the Properties panel agree about this on purpose (the owner, 2026-08-16),
    /// which is why the conversion lives here rather than twice: the panel spells the same
    /// values through the same arithmetic, and only its SYSTEM rows differ, where a panel can
    /// afford a name and a document cannot.
    /// </summary>
    public static string SpellColour(int ole) => (ole & SystemColourBit) != 0
        ? $"&H{(uint)ole:X8}&"
        : $"#{ole & 0xFF:x2}{(ole >> 8) & 0xFF:x2}{(ole >> 16) & 0xFF:x2}";

    /// <summary>
    /// A `#rrggbb` or `#rgb` as the number the model stores. OLE_COLOR keeps blue-green-red where
    /// CSS reads red-green-blue, which is the whole of the arithmetic and the whole of the reason
    /// nobody should write it twice. Null when the text is not a `#` colour at all.
    /// </summary>
    public static int? ReadColour(string spelled)
    {
        var text = spelled.Trim();
        if (text.Length < 4 || text[0] != '#')
        {
            return null;
        }

        var digits = text[1..];
        if (digits.Length == 3)
        {
            // #abc is #aabbcc, the shorthand every stylesheet takes.
            digits = string.Concat(digits.Select(character => new string(character, 2)));
        }

        return digits.Length == 6
            && int.TryParse(digits, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb)
            ? ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
            : null;
    }

    /// <summary>The bit that turns an OLE_COLOR from a colour into a question for the system.</summary>
    private const int SystemColourBit = unchecked((int)0x80000000);

    private static string Quoted(string value) => $"\"{value.Replace("\"", "\"\"")}\"";

    private static string Number(double value) => value.ToString(CultureInfo.InvariantCulture);

    // ------------------------------------------------------------------ parsing

    /// <summary>
    /// Every finding in a document, tolerantly - the squiggles' source. Parse stops at the
    /// first refusal because the apply's all-or-nothing promise depends on it; this collects
    /// them ALL, and the strict parser stays the ONE grammar: each refusal's line is blanked
    /// and the parse re-run, so a lint can never disagree with Parse about what is wrong,
    /// only continue past it. A blanked container orphans its children, and their findings
    /// are real - those lines ARE under nothing once the container line is bad. Semantic
    /// warnings - what an apply would note and skip - ride on the spec that survives.
    /// </summary>
    public static IReadOnlyList<FormMarkupFinding> Lint(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        var findings = new List<FormMarkupFinding>();
        var lines = text.Replace("\r\n", "\n").Split('\n');
        FormSpec? spec = null;

        // Bounded by the line count: every round either parses or retires one line.
        for (var attempt = 0; attempt <= lines.Length; attempt++)
        {
            try
            {
                spec = Parse(string.Join("\n", lines));
                break;
            }
            catch (FormMarkupException refused)
            {
                findings.Add(new FormMarkupFinding(refused.Line, refused.Reason, FormMarkupSeverity.Error));
                if (refused.Line < 1 || refused.Line > lines.Length)
                {
                    break;
                }

                // NO PROGRESS MEANS STOP. Retiring a line that is already empty parses the same
                // text again and refuses it again, so an EMPTY document squiggled "the document
                // is empty" twice on line 1 (found in the 2026-08-16 hunt). The round only
                // continues when it has actually taken a line out of the way.
                if (lines[refused.Line - 1].Length == 0)
                {
                    break;
                }

                lines[refused.Line - 1] = string.Empty;
            }
        }

        if (spec is null)
        {
            return findings;
        }

        // The line a control's header sits on, for anchoring a semantic finding: the nth
        // line whose body reads "<word> <name>". Good enough for a squiggle; the name is
        // the identity, not the position.
        int LineOf(string name, int occurrence = 1)
        {
            var seen = 0;
            for (var index = 0; index < lines.Length; index++)
            {
                var body = lines[index].TrimStart();
                var space = body.IndexOf(' ');
                if (space > 0)
                {
                    var rest = body[(space + 1)..];
                    if (rest.StartsWith(name, StringComparison.OrdinalIgnoreCase)
                        && (rest.Length == name.Length || !char.IsLetterOrDigit(rest[name.Length])))
                    {
                        if (++seen == occurrence)
                        {
                            return index + 1;
                        }
                    }
                }
            }

            return 1;
        }

        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { spec.Name };
        var typeOf = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var control in spec.Controls)
        {
            if (!taken.Add(control.Name))
            {
                findings.Add(new FormMarkupFinding(LineOf(control.Name, 2),
                    $"the name '{control.Name}' is already taken on this form", FormMarkupSeverity.Error));
            }
            else
            {
                typeOf[control.Name] = control.Type;
            }

            // A NAME MSFORMS WILL NOT TAKE, said here rather than left to the apply. A control's
            // name is a VBA identifier, and the model's own answer for one that is not is
            // `error 800a9c6c` and nothing else (measured 2026-08-16) - a squiggle on the line,
            // before anything is written, is the difference between a typo and a mystery.
            if (!IsIdentifier(control.Name))
            {
                findings.Add(new FormMarkupFinding(LineOf(control.Name),
                    $"'{control.Name}' is not a name MSForms will take: a control's name starts with a "
                        + "letter and holds only letters, digits and underscores",
                    FormMarkupSeverity.Error));
            }
        }

        // A stray Page needs no row here: the parser refuses one structurally, so it arrives
        // above as an Error and can never reach the spec.
        foreach (var control in spec.Controls)
        {
            if (!string.Equals(control.Type, "Page", StringComparison.OrdinalIgnoreCase)
                && !ToolboxTypes.Contains(control.Type)
                && !control.Properties.Any(p => string.Equals(p.Path, "ProgId", StringComparison.OrdinalIgnoreCase)))
            {
                findings.Add(new FormMarkupFinding(LineOf(control.Name),
                    $"'{control.Type}' is not a toolbox kind and no ProgId line names one, so an apply would skip it",
                    FormMarkupSeverity.Warning));
            }
        }

        return [.. findings.OrderBy(finding => finding.Line)];
    }

    /// <summary>
    /// A name MSForms will take for a control, which is VBA's own identifier rule: a letter
    /// first, then letters, digits and underscores, up to forty characters.
    /// </summary>
    public static bool IsIdentifier(string name) =>
        name.Length is > 0 and <= 40
        && char.IsLetter(name[0])
        && name.All(letter => char.IsLetterOrDigit(letter) || letter == '_');

    /// <summary>The toolbox spellings the dialect commits to - the same set the printer
    /// writes and an apply resolves without a ProgId line.</summary>
    private static readonly HashSet<string> ToolboxTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton",
        "Frame", "CommandButton", "TabStrip", "MultiPage", "Page", "Tab", "ScrollBar", "SpinButton", "Image",
    };

    /// <summary>
    /// Reads a whole document, or refuses it with the line that is wrong. Nothing partial: an
    /// apply built on this can promise that a document with an error changes nothing.
    /// </summary>
    public static FormSpec Parse(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        FormHeader? form = null;
        var formProperties = new List<PropertySpec>();
        var controls = new List<ControlSpec>();
        var properties = new Dictionary<string, List<PropertySpec>>(StringComparer.OrdinalIgnoreCase);

        // The container each indentation level is inside of. [0] is the form itself.
        var stack = new List<(string Name, string Type)>();

        var lines = text.Split('\n');
        for (var index = 0; index < lines.Length; index++)
        {
            var lineNumber = index + 1;
            var raw = lines[index].TrimEnd('\r');
            var content = StripComment(raw, lineNumber);
            if (content.TrimEnd().Length == 0)
            {
                continue;
            }

            if (content.Contains('\t'))
            {
                throw new FormMarkupException(lineNumber, "indent with spaces; this host never stores a tab");
            }

            var indent = content.Length - content.TrimStart(' ').Length;
            if (indent % IndentWidth != 0)
            {
                throw new FormMarkupException(lineNumber, $"indent by {IndentWidth} spaces per level");
            }

            var depth = indent / IndentWidth;
            var body = content.Trim();

            if (form is null)
            {
                if (depth != 0)
                {
                    throw new FormMarkupException(lineNumber, "the document opens with an unindented Form line");
                }

                form = ParseFormHeader(body, lineNumber);
                stack.Add((form.Name, "Form"));
                continue;
            }

            if (depth == 0)
            {
                throw new FormMarkupException(lineNumber, "one Form per document; nothing else sits unindented");
            }

            if (depth > stack.Count)
            {
                throw new FormMarkupException(lineNumber, "this line is indented under nothing");
            }

            // Stepping back out: the levels deeper than this line are done.
            stack.RemoveRange(depth, stack.Count - depth);
            var owner = stack[depth - 1];

            if (IsProperty(body))
            {
                var property = ParseProperty(body, lineNumber);
                if (depth == 1)
                {
                    formProperties.Add(property);
                }
                else if (owner.Type == "Form")
                {
                    throw new FormMarkupException(lineNumber, "a property line sits under the control it belongs to");
                }
                else
                {
                    properties[owner.Name].Add(property);
                }

                continue;
            }

            var control = ParseControlHeader(body, lineNumber);

            if (string.Equals(owner.Type, "MultiPage", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(control.Type, "Page", StringComparison.OrdinalIgnoreCase))
            {
                throw new FormMarkupException(lineNumber, "a MultiPage holds Pages; controls go on a Page");
            }

            // A TabStrip holds TABS and nothing else: what sits over its face belongs to the form,
            // because MSForms draws a strip over one set of controls and swaps them in code. The
            // dialect learned tabs on 2026-08-16, when a strip's line stood with nothing indented
            // under it while the canvas drew two tabs.
            if (string.Equals(owner.Type, "TabStrip", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(control.Type, "Tab", StringComparison.OrdinalIgnoreCase))
            {
                throw new FormMarkupException(lineNumber,
                    "a TabStrip holds Tabs; a control over its face belongs to the form");
            }

            if (string.Equals(control.Type, "Tab", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(owner.Type, "TabStrip", StringComparison.OrdinalIgnoreCase))
            {
                throw new FormMarkupException(lineNumber, "a Tab sits under a TabStrip");
            }

            if (string.Equals(control.Type, "Page", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(owner.Type, "MultiPage", StringComparison.OrdinalIgnoreCase))
            {
                throw new FormMarkupException(lineNumber, "a Page sits under a MultiPage");
            }

            if (owner.Type != "Form" && !Containers.Contains(owner.Type))
            {
                throw new FormMarkupException(lineNumber,
                    $"a {owner.Type} holds no controls; only Frame, MultiPage, Page and TabStrip contain");
            }

            var parent = depth == 1 ? null : owner.Name;
            var ownProperties = new List<PropertySpec>();
            properties[control.Name] = ownProperties;
            controls.Add(new ControlSpec(
                control.Type, control.Name, control.Caption,
                control.Left, control.Top, control.Width, control.Height,
                parent, ownProperties));

            stack.Add((control.Name, control.Type));
        }

        if (form is null)
        {
            throw new FormMarkupException(1, "the document is empty; it opens with a Form line");
        }

        return new FormSpec(form.Name, form.Caption, form.Width, form.Height, formProperties, controls);
    }

    /// <summary>An apostrophe outside quotes starts a comment, exactly as it does in VBA.</summary>
    private static string StripComment(string line, int lineNumber)
    {
        var inQuotes = false;
        for (var at = 0; at < line.Length; at++)
        {
            if (line[at] == '"')
            {
                inQuotes = !inQuotes;
            }
            else if (line[at] == '\'' && !inQuotes)
            {
                return line[..at];
            }
        }

        if (inQuotes)
        {
            throw new FormMarkupException(lineNumber, "a quote opens and never closes");
        }

        return line;
    }

    private static bool IsProperty(string body)
    {
        // An = outside quotes makes a property line; a header never carries one.
        var inQuotes = false;
        foreach (var character in body)
        {
            if (character == '"')
            {
                inQuotes = !inQuotes;
            }
            else if (character == '=' && !inQuotes)
            {
                return true;
            }
        }

        return false;
    }

    private sealed record FormHeader(string Name, string? Caption, double? Width, double? Height);

    private static FormHeader ParseFormHeader(string body, int lineNumber)
    {
        var reader = new LineReader(body, lineNumber);
        var keyword = reader.Word("Form");
        if (!string.Equals(keyword, "Form", StringComparison.OrdinalIgnoreCase))
        {
            throw new FormMarkupException(lineNumber, "the document opens with a Form line");
        }

        var name = reader.Word("the form's name");
        var caption = reader.QuotedIfPresent();

        double? width = null;
        double? height = null;
        if (reader.TakeKeyword("size"))
        {
            (width, height) = reader.Pair('x', "size");
        }

        reader.End();
        return new FormHeader(name, caption, width, height);
    }

    private sealed record ControlHeader(
        string Type, string Name, string? Caption, double? Left, double? Top, double? Width, double? Height);

    private static ControlHeader ParseControlHeader(string body, int lineNumber)
    {
        var reader = new LineReader(body, lineNumber);
        var type = reader.Word("a control kind");
        var name = reader.Word($"a name for the {type}");
        var caption = reader.QuotedIfPresent();

        double? left = null;
        double? top = null;
        double? width = null;
        double? height = null;
        if (reader.TakeKeyword("at"))
        {
            (left, top) = reader.Pair(',', "at");
        }

        if (reader.TakeKeyword("size"))
        {
            (width, height) = reader.Pair('x', "size");
        }

        reader.End();
        return new ControlHeader(type, name, caption, left, top, width, height);
    }

    private static PropertySpec ParseProperty(string body, int lineNumber)
    {
        var equals = IndexOfUnquoted(body, '=');
        var path = body[..equals].TrimEnd();
        var spelled = body[(equals + 1)..].Trim();

        if (path.Length == 0 || path.Split('.').Any(part => part.Length == 0 || !IsWord(part)))
        {
            throw new FormMarkupException(lineNumber, "the left of = is a property name, dots allowed");
        }

        if (spelled.Length == 0)
        {
            throw new FormMarkupException(lineNumber, $"{path} = what? The value is missing");
        }

        if (spelled[0] == '"')
        {
            var reader = new LineReader(spelled, lineNumber);
            var value = reader.QuotedIfPresent()
                ?? throw new FormMarkupException(lineNumber, "a quote opens and never closes");
            reader.End();
            return new PropertySpec(path, value, PropertyValueKind.Text);
        }

        if (bool.TryParse(spelled, out var flag))
        {
            return new PropertySpec(path, flag ? "True" : "False", PropertyValueKind.Flag);
        }

        // #c0dcc0 - the spelling the Properties panel shows and every developer already knows.
        // Carried onward as the decimal the designer model takes, like every other colour.
        if (spelled[0] == '#')
        {
            return ReadColour(spelled) is { } rgb
                ? new PropertySpec(path, rgb.ToString(CultureInfo.InvariantCulture), PropertyValueKind.Colour)
                : throw new FormMarkupException(lineNumber, $"{spelled} is not a #rrggbb colour");
        }

        // &H8000000F& - VBA's own spelling, and the only one a SYSTEM colour has. Carried onward
        // as the decimal the designer model actually takes, signed through the OLE_COLOR bit.
        if (spelled.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
        {
            var digits = spelled.TrimEnd('&')[2..];
            if (digits.Length is 0 or > 8
                || !long.TryParse(digits, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var hex))
            {
                throw new FormMarkupException(lineNumber, $"{spelled} is not a &H hex number");
            }

            return new PropertySpec(path, ((int)(uint)hex).ToString(CultureInfo.InvariantCulture), PropertyValueKind.Colour);
        }

        if (double.TryParse(spelled, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
        {
            return new PropertySpec(path, spelled, PropertyValueKind.Number);
        }

        throw new FormMarkupException(
            lineNumber, $"{spelled} is not a value: quoted text, a number, True/False, or &H hex");
    }

    private static int IndexOfUnquoted(string body, char wanted)
    {
        var inQuotes = false;
        for (var at = 0; at < body.Length; at++)
        {
            if (body[at] == '"')
            {
                inQuotes = !inQuotes;
            }
            else if (body[at] == wanted && !inQuotes)
            {
                return at;
            }
        }

        return -1;
    }

    private static bool IsWord(string text) =>
        text.Length > 0 && (char.IsLetter(text[0]) || text[0] == '_')
        && text.All(character => char.IsLetterOrDigit(character) || character == '_');

    /// <summary>
    /// One header line, read left to right. Small and hand-rolled, because the grammar is one
    /// line deep and a real lexer would be more code than the language.
    /// </summary>
    private ref struct LineReader(string body, int lineNumber)
    {
        private readonly string _body = body;
        private readonly int _line = lineNumber;
        private int _at;

        private void SkipSpaces()
        {
            while (_at < _body.Length && _body[_at] == ' ')
            {
                _at++;
            }
        }

        public string Word(string wanted)
        {
            SkipSpaces();
            var start = _at;
            while (_at < _body.Length && (char.IsLetterOrDigit(_body[_at]) || _body[_at] == '_'))
            {
                _at++;
            }

            if (_at == start)
            {
                throw new FormMarkupException(_line, $"expected {wanted}");
            }

            return _body[start.._at];
        }

        public bool TakeKeyword(string keyword)
        {
            SkipSpaces();
            if (_at + keyword.Length > _body.Length
                || !string.Equals(_body.Substring(_at, keyword.Length), keyword, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var after = _at + keyword.Length;
            if (after < _body.Length && _body[after] != ' ')
            {
                return false;
            }

            _at = after;
            return true;
        }

        public string? QuotedIfPresent()
        {
            SkipSpaces();
            if (_at >= _body.Length || _body[_at] != '"')
            {
                return null;
            }

            var value = new StringBuilder();
            _at++;
            while (_at < _body.Length)
            {
                if (_body[_at] == '"')
                {
                    if (_at + 1 < _body.Length && _body[_at + 1] == '"')
                    {
                        value.Append('"');
                        _at += 2;
                        continue;
                    }

                    _at++;
                    return value.ToString();
                }

                value.Append(_body[_at]);
                _at++;
            }

            throw new FormMarkupException(_line, "a quote opens and never closes");
        }

        public (double First, double Second) Pair(char separator, string keyword)
        {
            var first = Value(keyword);
            SkipSpaces();
            if (_at >= _body.Length || char.ToLowerInvariant(_body[_at]) != char.ToLowerInvariant(separator))
            {
                throw new FormMarkupException(_line, $"{keyword} takes two numbers split by {separator}");
            }

            _at++;
            var second = Value(keyword);
            return (first, second);
        }

        private double Value(string keyword)
        {
            SkipSpaces();
            var start = _at;
            while (_at < _body.Length && (char.IsDigit(_body[_at]) || _body[_at] == '.' || _body[_at] == '-'))
            {
                _at++;
            }

            if (start == _at
                || !double.TryParse(_body[start.._at], NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            {
                throw new FormMarkupException(_line, $"{keyword} takes numbers");
            }

            return value;
        }

        public void End()
        {
            SkipSpaces();
            if (_at < _body.Length)
            {
                throw new FormMarkupException(_line, $"could not read the rest of the line: \"{_body[_at..]}\"");
            }
        }
    }
}

/// <summary>A form as the markup describes it. Controls are flat, with parent NAMES - the
/// designer walk's own shape, and the shape a name-keyed diff wants.</summary>
public sealed record FormSpec(
    string Name,
    string? Caption,
    double? Width,
    double? Height,
    IReadOnlyList<PropertySpec> Properties,
    IReadOnlyList<ControlSpec> Controls);

public sealed record ControlSpec(
    string Type,
    string Name,
    string? Caption,
    double? Left,
    double? Top,
    double? Width,
    double? Height,
    string? Parent,
    IReadOnlyList<PropertySpec> Properties);

/// <summary>One `Path = value` line. The kind survives so the apply side writes the type the
/// developer spelled rather than re-guessing it.</summary>
public sealed record PropertySpec(string Path, string Value, PropertyValueKind Kind);

public enum PropertyValueKind
{
    Text,
    Number,
    Flag,

    /// <summary>A colour: carried as the decimal the model stores, spelled `#rrggbb` - or the
    /// VBA hex where the value is a system colour, which has no honest `#`. It applies as a
    /// number like any other; the kind exists so the PRINTER knows to spell it.</summary>
    Colour,
}

/// <summary>A refusal with the line that earned it. Parsing is all-or-nothing, so an apply
/// built on it can promise that a document with an error changes nothing.</summary>
public sealed class FormMarkupException(int line, string message)
    : Exception($"line {line}: {message}")
{
    public int Line { get; } = line;

    /// <summary>The reason alone, for a finding that carries the line separately.</summary>
    public string Reason { get; } = message;
}

/// <summary>One lint finding: the line, the reason, and how hard it is wrong. An Error is
/// what Parse refuses - an apply of this document changes nothing; a Warning is what an
/// apply would note and skip.</summary>
public sealed record FormMarkupFinding(int Line, string Message, FormMarkupSeverity Severity);

public enum FormMarkupSeverity
{
    Error,
    Warning,
}
