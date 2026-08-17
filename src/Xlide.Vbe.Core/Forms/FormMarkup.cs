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
            text.Append(" Width=").Append(Quoted(Number(width)))
                .Append(" Height=").Append(Quoted(Number(height)));
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
            text.Append(" Left=").Append(Quoted(Number(left)))
                .Append(" Top=").Append(Quoted(Number(top)));
        }

        if (control.Width is { } width && control.Height is { } height)
        {
            text.Append(" Width=").Append(Quoted(Number(width)))
                .Append(" Height=").Append(Quoted(Number(height)));
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
        text.Append(' ').Append(property.Path).Append('=').Append(Quoted(property.Kind switch
        {
            PropertyValueKind.Colour when int.TryParse(
                property.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var ole)
                => SpellColour(ole),
            _ => property.Value,
        }));
    }

    /// <summary>The bit that turns an OLE_COLOR from a colour into a question for the system.</summary>
    private const int SystemColourBit = unchecked((int)0x80000000);

    /// <summary>
    /// Win32's COLOR_ constants in plain words - the one table here that cannot be measured,
    /// because no call hands back a display name for COLOR_BTNFACE. The VALUE is always measured
    /// through GetSysColor; only the wording is written down, and it is the wording the native
    /// picker's System tab uses.
    ///
    /// One list for both surfaces. The Properties panel shows these names with their spaces
    /// ("Button Face") because a panel can afford them; the document spells the same colour with
    /// the spaces closed up (`ButtonFace`) because an attribute value that is not quoted cannot
    /// hold a space - and quoting it would make it indistinguishable from a caption.
    /// </summary>
    public static readonly IReadOnlyList<(int Index, string Name)> SystemColourNames =
    [
        (0, "Scroll Bars"), (1, "Desktop"), (2, "Active Title Bar"), (3, "Inactive Title Bar"),
        (4, "Menu Bar"), (5, "Window Background"), (6, "Window Frame"), (7, "Menu Text"),
        (8, "Window Text"), (9, "Active Title Bar Text"), (10, "Active Border"),
        (11, "Inactive Border"), (12, "Application Workspace"), (13, "Highlight"),
        (14, "Highlight Text"), (15, "Button Face"), (16, "Button Shadow"), (17, "Disabled Text"),
        (18, "Button Text"), (19, "Inactive Title Bar Text"), (20, "Button Highlight"),
        (21, "Button Dark Shadow"), (22, "Button Light Shadow"), (23, "Tooltip Text"),
        (24, "Tooltip"), (26, "Hot-Tracked Item"), (27, "Gradient Active Title Bar"),
        (28, "Gradient Inactive Title Bar"), (29, "Menu Highlight"), (30, "Menu Bar Background"),
    ];

    /// <summary>A name as the document spells it: no spaces, no hyphens, so it is one bare token.</summary>
    private static string Compact(string name) => name.Replace(" ", "").Replace("-", "");

    /// <summary>
    /// A colour as the document spells one: `#rrggbb` for a literal, and a NAME for a system
    /// colour (the owner, 2026-08-17: "can we support both then? either hex, or by friendly
    /// name?").
    ///
    /// The name is the point rather than decoration. A system colour is not an RGB but a question
    /// - what does this machine call a button face - and `ButtonFace` keeps it a question, where
    /// `#f0f0f0` would freeze today's answer into the form and stop it following the theme. So the
    /// two spellings carry the two different things a colour can BE, and a round trip loses
    /// neither.
    ///
    /// A system index with no name of its own keeps VBA's own hex, which is honest: a document
    /// that cannot say what a value means says what it is.
    /// </summary>
    public static string SpellColour(int ole)
    {
        if ((ole & SystemColourBit) == 0)
        {
            return $"#{ole & 0xFF:x2}{(ole >> 8) & 0xFF:x2}{(ole >> 16) & 0xFF:x2}";
        }

        var index = ole & 0xFF;
        foreach (var (known, name) in SystemColourNames)
        {
            if (known == index)
            {
                return Compact(name);
            }
        }

        return $"&H{(uint)ole:X8}&";
    }

    /// <summary>
    /// A system colour by name, or null when the word names none. Spaces, hyphens and case are all
    /// forgiven, so `ButtonFace`, `Button Face` and `buttonface` are one colour - the document
    /// writes the compact form and a developer may type whichever they remember.
    /// </summary>
    public static int? ReadColourName(string spelled)
    {
        var wanted = Compact(spelled.Trim());
        foreach (var (index, name) in SystemColourNames)
        {
            if (string.Equals(Compact(name), wanted, StringComparison.OrdinalIgnoreCase))
            {
                return SystemColourBit | index;
            }
        }

        return null;
    }

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

        // The line an element's OPENING tag sits on, for anchoring a semantic finding: the nth
        // line carrying `Name="..."` for this name. Good enough for a squiggle - the name is the
        // identity, not the position - and it finds the tag wherever in the element the attribute
        // was typed, because an element may wrap across lines.
        int LineOf(string name, int occurrence = 1)
        {
            var wanted = $"Name=\"{name}\"";
            var seen = 0;
            for (var index = 0; index < lines.Length; index++)
            {
                if (lines[index].Contains(wanted, StringComparison.OrdinalIgnoreCase)
                    && ++seen == occurrence)
                {
                    return index + 1;
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

        var tags = ScanTags(text);
        if (tags.Count == 0)
        {
            throw new FormMarkupException(1, "the document is empty; it opens with a <Form> tag");
        }

        FormSpec? built = null;
        var controls = new List<ControlSpec>();

        // What each open element is, innermost last. [0] is the form.
        var stack = new List<(string Name, string Type)>();
        var closed = false;

        foreach (var tag in tags)
        {
            if (closed)
            {
                throw new FormMarkupException(tag.Line, "one Form per document; nothing follows </Form>");
            }

            if (tag.IsClose)
            {
                if (stack.Count == 0)
                {
                    throw new FormMarkupException(tag.Line, $"</{tag.Name}> closes nothing");
                }

                var open = stack[^1];
                if (!string.Equals(open.Type, tag.Name, StringComparison.OrdinalIgnoreCase))
                {
                    throw new FormMarkupException(
                        tag.Line, $"</{tag.Name}> closes a <{open.Type}>; tags nest, they do not overlap");
                }

                stack.RemoveAt(stack.Count - 1);
                closed = stack.Count == 0;
                continue;
            }

            if (stack.Count == 0)
            {
                if (!string.Equals(tag.Name, "Form", StringComparison.OrdinalIgnoreCase))
                {
                    throw new FormMarkupException(tag.Line, "the document opens with a <Form> tag");
                }

                var (formName, caption, _, _, width, height, properties) = ReadAttributes(tag, isForm: true);
                built = new FormSpec(formName, caption, width, height, properties, controls);
                if (tag.SelfCloses)
                {
                    closed = true;
                    continue;
                }

                stack.Add((formName, "Form"));
                continue;
            }

            if (string.Equals(tag.Name, "Form", StringComparison.OrdinalIgnoreCase))
            {
                throw new FormMarkupException(tag.Line, "one Form per document; a Form holds no Form");
            }

            var owner = stack[^1];
            RefuseBadContainment(owner.Type, tag.Name, tag.Line);

            var (name, ownCaption, left, top, ownWidth, ownHeight, ownProperties) =
                ReadAttributes(tag, isForm: false);

            controls.Add(new ControlSpec(
                tag.Name, name, ownCaption, left, top, ownWidth, ownHeight,
                stack.Count == 1 ? null : owner.Name, ownProperties));

            if (!tag.SelfCloses)
            {
                stack.Add((name, tag.Name));
            }
        }

        if (built is null)
        {
            throw new FormMarkupException(1, "the document is empty; it opens with a <Form> tag");
        }

        if (stack.Count > 0)
        {
            throw new FormMarkupException(
                tags[^1].Line, $"<{stack[^1].Type}> is never closed");
        }

        return built;
    }

    /// <summary>
    /// The containment rules, unchanged by the move to tags because they are MSForms' rules
    /// rather than the syntax's: a MultiPage holds Pages, a TabStrip holds Tabs, a Page sits on a
    /// MultiPage and a Tab on a TabStrip, and nothing else contains anything.
    /// </summary>
    private static void RefuseBadContainment(string ownerType, string childType, int line)
    {
        if (string.Equals(ownerType, "MultiPage", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(childType, "Page", StringComparison.OrdinalIgnoreCase))
        {
            throw new FormMarkupException(line, "a MultiPage holds Pages; controls go on a Page");
        }

        // A TabStrip holds TABS and nothing else: what sits over its face belongs to the form,
        // because MSForms draws a strip over one set of controls and swaps them in code.
        if (string.Equals(ownerType, "TabStrip", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(childType, "Tab", StringComparison.OrdinalIgnoreCase))
        {
            throw new FormMarkupException(
                line, "a TabStrip holds Tabs; a control over its face belongs to the form");
        }

        if (string.Equals(childType, "Tab", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(ownerType, "TabStrip", StringComparison.OrdinalIgnoreCase))
        {
            throw new FormMarkupException(line, "a Tab sits under a TabStrip");
        }

        if (string.Equals(childType, "Page", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(ownerType, "MultiPage", StringComparison.OrdinalIgnoreCase))
        {
            throw new FormMarkupException(line, "a Page sits under a MultiPage");
        }

        if (!string.Equals(ownerType, "Form", StringComparison.OrdinalIgnoreCase)
            && !Containers.Contains(ownerType))
        {
            throw new FormMarkupException(
                line, $"a {ownerType} holds no controls; only Frame, MultiPage, Page and TabStrip contain");
        }
    }

    /// <summary>
    /// One element's attributes, split into the six the header owns and the rest, which are
    /// properties. `Name` is required and is the identity everything else keys on.
    /// </summary>
    private static (string Name, string? Caption, double? Left, double? Top,
        double? Width, double? Height, List<PropertySpec> Properties) ReadAttributes(Tag tag, bool isForm)
    {
        string? name = null;
        string? caption = null;
        double? left = null;
        double? top = null;
        double? width = null;
        double? height = null;
        var properties = new List<PropertySpec>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var attribute in tag.Attributes)
        {
            if (!seen.Add(attribute.Name))
            {
                throw new FormMarkupException(
                    attribute.Line, $"{attribute.Name} is set twice on this <{tag.Name}>");
            }

            switch (attribute.Name.ToUpperInvariant())
            {
                // Name and Caption are TEXT, and text is quoted - the same rule the properties
                // keep, and the reason `Caption=maybe` is a refusal rather than a caption reading
                // "maybe". Without this the two structural strings were the one place a bare word
                // slipped through.
                case "NAME":
                    name = RequireQuoted(attribute);
                    continue;
                case "CAPTION":
                    caption = RequireQuoted(attribute);
                    continue;
                case "LEFT" when !isForm:
                    left = ReadNumber(attribute);
                    continue;
                case "TOP" when !isForm:
                    top = ReadNumber(attribute);
                    continue;
                case "WIDTH":
                    width = ReadNumber(attribute);
                    continue;
                case "HEIGHT":
                    height = ReadNumber(attribute);
                    continue;
                default:
                    properties.Add(ReadValue(attribute));
                    continue;
            }
        }

        if (name is null)
        {
            throw new FormMarkupException(tag.Line, $"a <{tag.Name}> needs a Name");
        }

        if (!IsIdentifier(name))
        {
            throw new FormMarkupException(
                tag.Line, $"{name} is not a name a control can take; letters, digits and _, not starting with a digit");
        }

        if ((left is null) != (top is null))
        {
            throw new FormMarkupException(tag.Line, "Left and Top come together or not at all");
        }

        if ((width is null) != (height is null))
        {
            throw new FormMarkupException(tag.Line, "Width and Height come together or not at all");
        }

        return (name, caption, left, top, width, height, properties);
    }

    private static string RequireQuoted(Attribute attribute) =>
        attribute.Quoted
            ? attribute.Value
            : throw new FormMarkupException(
                attribute.Line, $"{attribute.Name} is not a value: {attribute.Name} takes quoted text");

    private static double ReadNumber(Attribute attribute) =>
        double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value
            : throw new FormMarkupException(attribute.Line, $"{attribute.Name} takes a number");

    /// <summary>
    /// One attribute as a property, with its KIND recovered from the spelling. Quoted is text and
    /// nothing else is - which is the whole reason the printer leaves numbers bare, and the reason
    /// `Caption="True"` is a caption while `Enabled=True` is a flag.
    /// </summary>
    private static PropertySpec ReadValue(Attribute attribute)
    {
        var spelled = attribute.Value;
        if (spelled.Length == 0)
        {
            return new PropertySpec(attribute.Name, string.Empty, PropertyValueKind.Text);
        }

        if (bool.TryParse(spelled, out var flag))
        {
            return new PropertySpec(attribute.Name, flag ? "True" : "False", PropertyValueKind.Flag);
        }

        // #c0dcc0 - the spelling the Properties panel shows and every developer already knows.
        // Carried onward as the decimal the designer model takes, like every other colour.
        if (spelled[0] == '#')
        {
            return ReadColour(spelled) is { } rgb
                ? new PropertySpec(attribute.Name, rgb.ToString(CultureInfo.InvariantCulture), PropertyValueKind.Colour)
                : throw new FormMarkupException(attribute.Line, $"{spelled} is not a #rrggbb colour");
        }

        // &H8000000F& - VBA's own spelling, still taken from a hand-written document even though
        // the printer now writes a system colour's NAME.
        if (spelled.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
        {
            var digits = spelled.TrimEnd('&')[2..];
            if (digits.Length is 0 or > 8
                || !long.TryParse(digits, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var hex))
            {
                throw new FormMarkupException(attribute.Line, $"{spelled} is not a &H hex number");
            }

            return new PropertySpec(
                attribute.Name, ((int)(uint)hex).ToString(CultureInfo.InvariantCulture), PropertyValueKind.Colour);
        }

        if (double.TryParse(spelled, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
        {
            return new PropertySpec(attribute.Name, spelled, PropertyValueKind.Number);
        }

        // A bare word that names a system colour - `ButtonFace`, `Highlight`. Readable as a colour
        // only because every other bare shape is taken and text is always quoted.
        if (ReadColourName(spelled) is { } named)
        {
            return new PropertySpec(
                attribute.Name, named.ToString(CultureInfo.InvariantCulture), PropertyValueKind.Colour);
        }

        // Anything else is TEXT, which is the only honest default now that every attribute is
        // quoted: `Caption="12"` and `MaxLength="12"` are spelled identically, so the document
        // cannot carry the type and does not try to. The KIND read here is a good guess for the
        // canvas and the printer; the APPLY does not rely on it, because it asks the property
        // itself what it is (FormDesignService.InferKind) and a guess cannot mislead a write.
        return new PropertySpec(attribute.Name, spelled, PropertyValueKind.Text);
    }

    // ------------------------------------------------------------------ the scanner

    /// <summary>An element as the scanner found it. A close tag carries no attributes.</summary>
    private readonly record struct Tag(
        string Name, bool IsClose, bool SelfCloses, List<Attribute> Attributes, int Line);

    /// <summary>One attribute. `Quoted` is what tells TEXT from everything else, which is why the
    /// printer never quotes a number: the kind is recovered from the spelling.</summary>
    private readonly record struct Attribute(string Name, string Value, bool Quoted, int Line);

    /// <summary>
    /// Every tag in the document, in order. Comments are skipped, whitespace between tags is
    /// ignored, and anything ELSE between tags is refused - a value typed as element content
    /// rather than as an attribute would otherwise vanish without a word.
    /// </summary>
    private static List<Tag> ScanTags(string text)
    {
        var tags = new List<Tag>();
        var line = 1;
        var at = 0;

        while (at < text.Length)
        {
            var open = text.IndexOf('<', at);
            if (open < 0)
            {
                RefuseStrayText(text, at, text.Length, ref line);
                break;
            }

            RefuseStrayText(text, at, open, ref line);

            if (text.AsSpan(open).StartsWith("<!--"))
            {
                var end = text.IndexOf("-->", open, StringComparison.Ordinal);
                if (end < 0)
                {
                    throw new FormMarkupException(line, "a comment opens and never closes");
                }

                for (var i = open; i < end; i++)
                {
                    if (text[i] == '\n')
                    {
                        line++;
                    }
                }

                at = end + 3;
                continue;
            }

            var tagLine = line;
            var body = new StringBuilder();
            var inQuotes = false;
            var cursor = open + 1;
            var closedTag = false;
            while (cursor < text.Length)
            {
                var character = text[cursor];
                if (character == '\n')
                {
                    line++;
                }

                if (character == '"')
                {
                    // A doubled quote inside a quoted value is one quote, as it is in VBA.
                    if (inQuotes && cursor + 1 < text.Length && text[cursor + 1] == '"')
                    {
                        body.Append("\"\"");
                        cursor += 2;
                        continue;
                    }

                    inQuotes = !inQuotes;
                }
                else if (character == '>' && !inQuotes)
                {
                    closedTag = true;
                    cursor++;
                    break;
                }

                body.Append(character);
                cursor++;
            }

            if (!closedTag)
            {
                throw new FormMarkupException(tagLine, "a tag opens and never closes");
            }

            tags.Add(ReadTag(body.ToString(), tagLine));
            at = cursor;
        }

        return tags;
    }

    private static void RefuseStrayText(string text, int from, int to, ref int line)
    {
        for (var i = from; i < to; i++)
        {
            if (text[i] == '\n')
            {
                line++;
            }
            else if (!char.IsWhiteSpace(text[i]))
            {
                throw new FormMarkupException(
                    line, "text between tags says nothing here; a value belongs in an attribute");
            }
        }
    }

    /// <summary>The inside of one `&lt;...&gt;`: a name, then attributes.</summary>
    private static Tag ReadTag(string body, int line)
    {
        var trimmed = body.Trim();
        if (trimmed.Length == 0)
        {
            throw new FormMarkupException(line, "an empty tag names nothing");
        }

        var isClose = trimmed[0] == '/';
        if (isClose)
        {
            trimmed = trimmed[1..].Trim();
        }

        var selfCloses = !isClose && trimmed.EndsWith('/');
        if (selfCloses)
        {
            trimmed = trimmed[..^1].TrimEnd();
        }

        var nameEnd = 0;
        while (nameEnd < trimmed.Length && !char.IsWhiteSpace(trimmed[nameEnd]))
        {
            nameEnd++;
        }

        var name = trimmed[..nameEnd];
        if (name.Length == 0 || !IsIdentifier(name))
        {
            throw new FormMarkupException(line, $"<{trimmed}> does not open with a control kind");
        }

        var attributes = new List<Attribute>();
        var rest = trimmed[nameEnd..];
        if (isClose)
        {
            if (rest.Trim().Length > 0)
            {
                throw new FormMarkupException(line, $"</{name}> takes no attributes");
            }

            return new Tag(name, true, false, attributes, line);
        }

        var at = 0;
        while (at < rest.Length)
        {
            while (at < rest.Length && char.IsWhiteSpace(rest[at]))
            {
                at++;
            }

            if (at >= rest.Length)
            {
                break;
            }

            var start = at;
            while (at < rest.Length && (char.IsLetterOrDigit(rest[at]) || rest[at] == '_' || rest[at] == '.'))
            {
                at++;
            }

            var attributeName = rest[start..at];
            if (attributeName.Length == 0)
            {
                throw new FormMarkupException(line, $"{rest[at..].Trim()} is not an attribute");
            }

            while (at < rest.Length && char.IsWhiteSpace(rest[at]))
            {
                at++;
            }

            if (at >= rest.Length || rest[at] != '=')
            {
                throw new FormMarkupException(line, $"{attributeName} is missing its value");
            }

            at++;
            while (at < rest.Length && char.IsWhiteSpace(rest[at]))
            {
                at++;
            }

            if (at >= rest.Length)
            {
                throw new FormMarkupException(line, $"{attributeName} is missing its value");
            }

            if (rest[at] == '"')
            {
                at++;
                var value = new StringBuilder();
                var done = false;
                while (at < rest.Length)
                {
                    if (rest[at] == '"')
                    {
                        if (at + 1 < rest.Length && rest[at + 1] == '"')
                        {
                            value.Append('"');
                            at += 2;
                            continue;
                        }

                        at++;
                        done = true;
                        break;
                    }

                    value.Append(rest[at]);
                    at++;
                }

                if (!done)
                {
                    throw new FormMarkupException(line, "a quote opens and never closes");
                }

                attributes.Add(new Attribute(attributeName, value.ToString(), true, line));
                continue;
            }

            var bareStart = at;
            while (at < rest.Length && !char.IsWhiteSpace(rest[at]))
            {
                at++;
            }

            attributes.Add(new Attribute(attributeName, rest[bareStart..at], false, line));
        }

        return new Tag(name, false, selfCloses, attributes, line);
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

        // A bare word that names a system colour - `ButtonFace`, `Highlight`. It can only be read
        // as a colour because every other bare shape is already taken and TEXT is always quoted,
        // so nothing here has to guess between a caption and a colour.
        if (ReadColourName(spelled) is { } named)
        {
            return new PropertySpec(
                path, named.ToString(CultureInfo.InvariantCulture), PropertyValueKind.Colour);
        }

        throw new FormMarkupException(
            lineNumber,
            $"{spelled} is not a value: quoted text, a number, True/False, #rrggbb, "
            + "a system colour's name, or &H hex");
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
