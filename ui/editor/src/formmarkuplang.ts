/*
 * The markup language, as monaco sees it: its grammar, its indentation rules, and the language
 * service behind the designer tab's document half - completions, hover, and the header hint
 * (docs/userform-designer.md, the language service, step 3).
 *
 * WHAT IT ANSWERS FROM. Nothing here is a table of the language's own vocabulary. The host
 * measures it - a bare instance of each control's coclass, read through its own type library -
 * and sends it once per session; these providers offer exactly that, so a completion names a
 * property MSForms on this machine actually has, and a hover repeats what the Object Browser
 * would say. Until the answer arrives the providers offer the grammar alone, which is the honest
 * degradation: the shape of a line is page knowledge, its vocabulary is not.
 *
 * AND WHAT IT DOES NOT DECIDE. There is ONE grammar for this language and it is Core's parser,
 * host-side, the apply's own. The reading here - which container a line sits in, whether it is a
 * header or a property - is positional and suggestion-only: it can be wrong about what to OFFER,
 * never about what is valid, because it never refuses anything. The squiggles keep coming from
 * the host's tolerant parse.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";
import type { FormMarkupKind, FormMarkupProperty } from "./bridge.js";

/** The markup's own language id, for the designer tab's document. */
export const FORM_MARKUP_LANGUAGE = "xlide-form";


/**
 * The toolbox: the kinds a document can name, in the order the native palette lists them, each
 * with the size MSForms gives a control dropped rather than drawn. The palette draws from this
 * and a completion scaffolds from it, because "what a new Label looks like" is one fact.
 *
 * A Page is not here: a page is added to a MultiPage, not dropped on a form. It still completes -
 * the vocabulary carries it - but with no size, because a Page has no geometry of its own.
 */
export const TOOLBOX: readonly { kind: string; width: number; height: number }[] = [
  { kind: "Label", width: 66, height: 16 },
  { kind: "TextBox", width: 120, height: 20 },
  { kind: "ComboBox", width: 120, height: 20 },
  { kind: "ListBox", width: 120, height: 42 },
  { kind: "CheckBox", width: 66, height: 16 },
  { kind: "OptionButton", width: 76, height: 16 },
  { kind: "ToggleButton", width: 92, height: 22 },
  { kind: "Frame", width: 92, height: 66 },
  { kind: "CommandButton", width: 72, height: 24 },
  { kind: "TabStrip", width: 122, height: 86 },
  { kind: "MultiPage", width: 192, height: 86 },
  { kind: "ScrollBar", width: 14, height: 96 },
  { kind: "SpinButton", width: 14, height: 42 },
  { kind: "Image", width: 76, height: 42 },
];

/** The kinds that hold children, until the host's vocabulary says otherwise - it carries the
 * same fact, measured, and replaces this the moment it arrives. */
const CONTAINERS = new Set(["Frame", "MultiPage", "Page"]);

/* -------------------------------------------------------------- the vocabulary */

let vocabulary: FormMarkupKind[] = [];

/** The host's answer, held for the page's life. Kinds are few and the answer is measured from
 * coclasses that do not change while Excel is up, so one request per session is the whole
 * traffic - no keystroke ever waits on a round trip. */
export function setMarkupVocabulary(kinds: FormMarkupKind[]): void {
  vocabulary = kinds;
}

/** What the language service knows, for a probe that wants to check it against the route. */
export function markupVocabulary(): FormMarkupKind[] {
  return vocabulary;
}

function kindNamed(name: string): FormMarkupKind | null {
  return vocabulary.find((one) => one.kind.toLowerCase() === name.toLowerCase()) ?? null;
}

function propertyNamed(kind: string, path: string): FormMarkupProperty | null {
  return kindNamed(kind)?.properties.find((one) => one.name.toLowerCase() === path.toLowerCase()) ?? null;
}

/** Every kind a header line may name, vocabulary first and the toolbox as the floor: the
 * completions work before the host has answered, with the shape but not the meanings. */
function everyKind(): string[] {
  const names = vocabulary.filter((one) => one.kind !== "Form").map((one) => one.kind);
  return names.length > 0 ? names : [...TOOLBOX.map((one) => one.kind), "Page"];
}

function contains(kind: string): boolean {
  return kindNamed(kind)?.container ?? CONTAINERS.has(kind);
}

/* -------------------------------------------------------------- reading a place in the document */

/** Where the caret is, in the document's own terms. Positional, suggestion-only - see the
 * header of this file for why that is not a second grammar. */
interface Spot {
  /** 0 is the Form line and everything unindented; 1 its properties and the top-level controls. */
  depth: number;
  /** The kind whose vocabulary owns this line: the header above it, or Form at the top. */
  owner: string;
  /** The line as it stands, and the part of it before the caret. */
  line: string;
  before: string;
  /** The property path left of `=`, when this line carries one. */
  path: string | null;
  /** True once the caret is past the `=`: the value's side of a property line. */
  onValue: boolean;
  /** A header line's own words - the kind it names and the name it gives. */
  header: { kind: string; name: string } | null;
  /** How many words of the header stand before the caret, which is what the hint follows. */
  words: number;
}


/**
 * Where the caret stands, worked out by scanning from the top of the document rather than by
 * counting spaces. Indentation is presentation in the tagged dialect - the parser reads tags - so
 * the only way to know which element the caret is in is to keep the same stack the parser keeps,
 * and the only way to know it is on an attribute is to notice a `<` that has no `>` yet.
 *
 * Positional and suggestion-only, exactly as the line version was: it never has to be right the
 * way the parser has to be right, because nothing is written from it.
 */
function scanTo(text: string): { stack: string[]; inTag: string | null; tagText: string } {
  const stack: string[] = [];
  let at = 0;

  while (at < text.length) {
    const open = text.indexOf("<", at);
    if (open < 0) {
      break;
    }

    if (text.startsWith("<!--", open)) {
      const end = text.indexOf("-->", open);
      if (end < 0) {
        return { stack, inTag: null, tagText: "" };
      }

      at = end + 3;
      continue;
    }

    // The end of this tag, with quoted stretches skipped so a `>` inside a caption is text.
    let cursor = open + 1;
    let quoted = false;
    let end = -1;
    while (cursor < text.length) {
      const character = text[cursor];
      if (character === '"') {
        quoted = !quoted;
      } else if (character === ">" && !quoted) {
        end = cursor;
        break;
      }
      cursor++;
    }

    const body = text.slice(open + 1, end < 0 ? text.length : end);
    const name = body.replace(/^\//, "").trim().split(/[\s/>]/)[0] ?? "";

    if (end < 0) {
      // The caret is INSIDE this tag - it has no `>` yet - which is the attribute case.
      return { stack, inTag: name, tagText: body };
    }

    if (body.startsWith("/")) {
      stack.pop();
    } else if (!body.trimEnd().endsWith("/")) {
      stack.push(name);
    }

    at = end + 1;
  }

  return { stack, inTag: null, tagText: "" };
}

function spotAt(model: monaco.editor.ITextModel, position: monaco.IPosition): Spot {
  const line = model.getLineContent(position.lineNumber);
  const before = line.slice(0, position.column - 1);
  const { stack, inTag, tagText } = scanTo(model.getValueInRange({
    startLineNumber: 1,
    startColumn: 1,
    endLineNumber: position.lineNumber,
    endColumn: position.column,
  }));

  // Inside a tag the vocabulary is that element's own; between tags it is the element the caret
  // sits inside, which is the top of the stack, and the Form when the stack is empty.
  const owner = inTag ?? stack[stack.length - 1] ?? "Form";

  // The attribute being typed: the last `name=` in the tag so far. A value is still being typed
  // while nothing follows the `=` or while its quote stands open; once the quote closes, the
  // caret has moved on to where the next attribute goes.
  const lastEquals = inTag === null ? -1 : tagText.lastIndexOf("=");
  const path = lastEquals >= 0
    ? (tagText.slice(0, lastEquals).trim().split(/\s+/).pop() ?? null)
    : null;
  const afterEquals = lastEquals >= 0 ? tagText.slice(lastEquals + 1) : "";
  const onValue = lastEquals >= 0
    && (afterEquals.trim() === "" || (afterEquals.match(/"/g)?.length ?? 0) % 2 === 1);

  return {
    depth: stack.length,
    owner,
    line,
    before,
    path: onValue ? path : null,
    onValue,
    header: inTag !== null
      ? { kind: inTag, name: /\bName\s*=\s*"([^"]*)"/.exec(tagText)?.[1] ?? "" }
      : null,
    words: lastEquals >= 0 ? 2 : 1,
  };
}

/** The range a completion replaces when the token may carry dots - a `Font.Si` is one path, not
 * a word called `Si`, and monaco's own word rules split it. */
function tokenRange(model: monaco.editor.ITextModel, position: monaco.IPosition): monaco.Range {
  const line = model.getLineContent(position.lineNumber);
  let start = position.column - 1;
  while (start > 0 && /[\w.&]/.test(line[start - 1] ?? "")) {
    start--;
  }

  let end = position.column - 1;
  while (end < line.length && /[\w.&]/.test(line[end] ?? "")) {
    end++;
  }

  return new monaco.Range(position.lineNumber, start + 1, position.lineNumber, end + 1);
}

/**
 * The STRING already standing on the value's side of a property line, quotes included, so
 * accepting a value the developer has begun REPLACES it rather than nesting a second string
 * inside it: without this, completing at `FontName = "Tah` writes `FontName = ""Tahoma"`, because
 * a quote is not a word character and the token range stops short of it.
 *
 * The whole run, not the part before the caret, so picking a face with the caret sitting after a
 * finished value replaces the value rather than appending to it. Null when the value side holds
 * no string at all, where the plain token range is right.
 */
function valueStringRange(model: monaco.editor.ITextModel, position: monaco.IPosition): monaco.Range | null {
  const line = model.getLineContent(position.lineNumber);

  // The property path cannot hold a quote, so the first `=` is the one that splits the line.
  const equals = line.indexOf("=");
  if (equals < 0) {
    return null;
  }

  const open = line.indexOf('"', equals + 1);
  if (open < 0) {
    return null;
  }

  let close = open + 1;
  while (close < line.length && line[close] !== '"') {
    close++;
  }

  return new monaco.Range(
    position.lineNumber, open + 1,
    position.lineNumber, (close < line.length ? close + 1 : line.length) + 1);
}

/** The name the native toolbox would give a new control of this kind: the kind plus the first
 * free number, counted off the document itself. */
function freeName(model: monaco.editor.ITextModel, kind: string): string {
  const text = model.getValue();
  for (let number = 1; number < 1000; number++) {
    const candidate = `${kind}${number}`;
    if (!new RegExp(`\\b${candidate}\\b`, "i").test(text)) {
      return candidate;
    }
  }

  return `${kind}1`;
}

/* -------------------------------------------------------------- how a value is spelled */

/** A colour as the DOCUMENT spells one, which is what Core's printer writes and its parser
 * reads: `#rrggbb` for a plain colour, and the VBA hex for a system colour, which is a question
 * about the machine rather than an RGB. The same rule the Properties panel keeps. */
function spellColour(value: number): string {
  return (value & 0x80000000) !== 0
    ? `&H${(value >>> 0).toString(16).toUpperCase().padStart(8, "0")}&`
    : `#${(value & 0xFF).toString(16).padStart(2, "0")}`
      + `${((value >> 8) & 0xFF).toString(16).padStart(2, "0")}`
      + `${((value >> 16) & 0xFF).toString(16).padStart(2, "0")}`;
}

function asNumber(value: string | null): number | null {
  if (value === null) {
    return null;
  }

  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

/** What a property's untouched value looks like in a document, for a hover and for the one
 * completion that offers it back. */
function spelledDefault(property: FormMarkupProperty): string | null {
  if (property.default === null) {
    return null;
  }

  const number = asNumber(property.default);
  if (property.members && number !== null) {
    return property.members.find((member) => member.value === number)?.name ?? property.default;
  }

  if (property.colour && number !== null) {
    return spellColour(number);
  }

  return property.default;
}

/** The line a hover leads with: the path, and the type it takes. */
function signatureOf(owner: string, property: FormMarkupProperty): string {
  const type = property.type ?? (property.colour ? "OLE_COLOR" : "Variant");
  return `${owner}.${property.name} As ${type}`;
}

function documentationOf(property: FormMarkupProperty): string[] {
  const lines: string[] = [];
  if (property.doc) {
    lines.push(property.doc);
  }

  const untouched = spelledDefault(property);
  if (untouched !== null) {
    lines.push(`Default \`${untouched}\`.`);
  }

  if (property.members && property.members.length > 0) {
    lines.push(property.members.map((member) => `- \`${member.name}\` = ${member.value}`).join("\n"));
  } else if (property.colour) {
    lines.push("A colour: `#rrggbb`, or a system colour like `&H8000000F&`.");
  }

  return lines;
}

/* -------------------------------------------------------------- completions */

export function completionsAt(
  model: monaco.editor.ITextModel, position: monaco.IPosition,
): monaco.languages.CompletionList {
  const spot = spotAt(model, position);
  const range = tokenRange(model, position);
  const items: monaco.languages.CompletionItem[] = [];

  // The value's side of a property line: what this property can hold, by name.
  if (spot.onValue && spot.path !== null) {
    const property = propertyNamed(spot.owner, spot.path);
    if (property) {
      for (const member of property.members ?? []) {
        items.push({
          label: member.name,
          kind: monaco.languages.CompletionItemKind.EnumMember,
          detail: `${member.value}`,
          ...(property.doc ? { documentation: property.doc } : {}),
          insertText: member.name,
          range,
        });
      }

      // Values the machine was asked for rather than the type library - a font's faces. Offered
      // the same way and no more binding: the property takes anything typed instead, which is why
      // these are a plain Value rather than an EnumMember.
      //
      // A face is a STRING in the dialect, so the whole quoted run is what gets replaced when the
      // developer has already opened one, and the filter text carries the quote with it - monaco
      // filters on the text from the range's start to the caret, and `"Tah` matches nothing at all
      // against a bare `Tahoma`.
      const inString = valueStringRange(model, position);
      for (const value of property.values ?? []) {
        items.push({
          label: value,
          kind: monaco.languages.CompletionItemKind.Value,
          insertText: `"${value}"`,
          ...(inString ? { filterText: `"${value}"` } : {}),
          range: inString ?? range,
        });
      }

      if ((property.type ?? "") === "Boolean") {
        for (const flag of ["True", "False"]) {
          items.push({
            label: flag,
            kind: monaco.languages.CompletionItemKind.Value,
            insertText: flag,
            range,
          });
        }
      }

      const untouched = spelledDefault(property);
      if (untouched !== null && !items.some((item) => item.label === untouched)) {
        items.push({
          label: untouched,
          kind: monaco.languages.CompletionItemKind.Value,
          detail: "default",
          insertText: untouched,
          range,
        });
      }
    }

    return { suggestions: items };
  }

  // Inside a caption nothing is offered: the text between quotes is the developer's, and a
  // completion widget over it is noise. An odd number of quotes behind the caret means one
  // is still open, which is the same test the parser's comment stripper makes.
  if ((spot.before.match(/"/g)?.length ?? 0) % 2 === 1) {
    return { suggestions: items };
  }

  // INSIDE A TAG the offer is ATTRIBUTES, never control kinds: an element's own properties plus
  // the universals it does not carry yet. A tag already spelling `Left=` is not offered a second
  // one, which is the same "only what this can still take" rule the clauses had.
  if (spot.header !== null) {
    const already = (name: string) => new RegExp(`\\b${name}\\s*=`, "i").test(spot.before);
    for (const universal of ["Name", "Caption", "Left", "Top", "Width", "Height"]) {
      if (!already(universal)) {
        items.push({
          label: universal,
          kind: monaco.languages.CompletionItemKind.Property,
          detail: universal === "Name" || universal === "Caption" ? "text" : "points",
          insertText: `${universal}="$1"`,
          insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
          range,
        });
      }
    }

    for (const property of kindNamed(spot.header.kind)?.properties ?? []) {
      if (already(property.name)) {
        continue;
      }

      items.push({
        label: property.name,
        kind: monaco.languages.CompletionItemKind.Property,
        ...(property.type ? { detail: property.type } : {}),
        documentation: { value: documentationOf(property).join("\n\n") },
        insertText: `${property.name}="$1"`,
        insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
        // The value list follows without a second keystroke: the point of picking a property
        // is to say what it holds.
        command: { id: "editor.action.triggerSuggest", title: "values" },
        range,
      });
    }

    return { suggestions: items };
  }

  // BETWEEN TAGS the offer is the control KINDS this container takes, as whole elements. Outside
  // the Form's own tag there is nothing to add, because a document holds one Form.
  if (spot.depth === 0) {
    return { suggestions: items };
  }

  if (contains(spot.owner) || spot.owner === "Form") {
    for (const kind of kindsUnder(spot.owner)) {
      items.push(kindCompletion(model, kind, range));
    }
  }

  return { suggestions: items };
}

/** Which kinds may open a line inside this owner: a MultiPage holds Pages and nothing else, and
 * a Page belongs to nothing else. The parser refuses the rest; this says so early. */
function kindsUnder(owner: string): string[] {
  if (owner.toLowerCase() === "multipage") {
    return ["Page"];
  }

  return everyKind().filter((kind) => kind.toLowerCase() !== "page");
}

function kindCompletion(
  model: monaco.editor.ITextModel, kind: string, range: monaco.Range,
): monaco.languages.CompletionItem {
  const size = TOOLBOX.find((one) => one.kind === kind);
  const name = freeName(model, kind);
  const known = kindNamed(kind);

  // A whole ELEMENT, self-closing, because that is what a control is now. A Page has no geometry
  // of its own; everything else arrives placed and sized at MSForms' own drop size, which is the
  // scaffolding half of this feature - an element a developer can apply without finishing it.
  const insertText = size
    ? `<${kind} Name="\${1:${name}}" Caption="\${2:${name}}" Left="\${3:12}" Top="\${4:12}" `
      + `Width="\${5:${size.width}}" Height="\${6:${size.height}}" />`
    : `<${kind} Name="\${1:${name}}" Caption="\${2:${name}}" />`;

  // The coclass only where it is not simply `Forms.` plus this kind, for the hover's own reason:
  // a card that restates the word being completed is a card in the way.
  const documentation = [
    known?.doc ?? null,
    known?.progId && known.progId !== `Forms.${known.kind}.1` ? `\`${known.progId}\`` : null,
  ].filter((line): line is string => line !== null);

  return {
    label: kind,
    kind: monaco.languages.CompletionItemKind.Class,
    ...(size ? { detail: `${size.width}x${size.height} points` } : {}),
    ...(documentation.length > 0 ? { documentation: { value: documentation.join("  \n") } } : {}),
    insertText,
    insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
    range,
  };
}

/** The six every element carries, which no type library describes because they are the
 * document's own scaffolding rather than MSForms properties. */
const UNIVERSAL_ATTRIBUTES: Readonly<Record<string, string>> = {
  name: "The control's identity, and how the code-behind reaches it.",
  caption: "The text the control shows.",
  left: "Points from the container's left edge.",
  top: "Points from the container's top edge.",
  width: "Width in points.",
  height: "Height in points.",
};

/* -------------------------------------------------------------- hover */

export function hoverAt(
  model: monaco.editor.ITextModel, position: monaco.IPosition,
): monaco.languages.Hover | null {
  const spot = spotAt(model, position);
  const range = tokenRange(model, position);
  const token = model.getValueInRange(range);
  if (token === "") {
    return null;
  }

  const contents = (...values: (string | null)[]): monaco.languages.Hover => ({
    range,
    contents: values
      .filter((value): value is string => value !== null && value !== "")
      .map((value) => ({ value })),
  });

  // A property line: the path, or a value that names an enum member.
  if (spot.path !== null) {
    const property = propertyNamed(spot.owner, spot.path);

    // FALL THROUGH rather than answering nothing. `Name` and `Caption` are attributes the type
    // library never lists - the vocabulary drops them - so returning null here swallowed the
    // hover for a control's own name, which is the one a developer points at most.
    if (property !== null) {
      if (spot.onValue) {
        const member = property.members?.find((one) => one.name.toLowerCase() === token.toLowerCase());
        return member
          ? contents("```vba\n" + `${member.name} = ${member.value}` + "\n```", property.doc)
          : contents("```vba\n" + signatureOf(spot.owner, property) + "\n```", ...documentationOf(property));
      }

      return contents(
        "```vba\n" + signatureOf(spot.owner, property) + "\n```", ...documentationOf(property));
    }
  }

  if (spot.header === null) {
    return null;
  }

  // AN ATTRIBUTE NAME, which is where most of a tagged document's hovering happens: the token is
  // neither the kind nor the control's name, so it is looked up as a property of this element.
  // `spot.path` only fills once the caret is past the `=`, and the pointer over `MatchEntry` is
  // not - which is why this asks the vocabulary rather than the spot.
  if (token.toLowerCase() !== spot.header.kind.toLowerCase()
    && token.toLowerCase() !== spot.header.name.toLowerCase()) {
    const attribute = propertyNamed(spot.header.kind, token);
    if (attribute) {
      return contents(
        "```vba\n" + signatureOf(spot.header.kind, attribute) + "\n```",
        ...documentationOf(attribute));
    }

    // The universals carry no type library entry of their own, and saying nothing about `Left`
    // would leave the hole where the commonest attributes are.
    const said = UNIVERSAL_ATTRIBUTES[token.toLowerCase()];
    if (said !== undefined) {
      return contents("```vba\n" + `${spot.header.name || spot.header.kind}.${token}` + "\n```", said);
    }
  }

  /*
   * THE KIND: what that class of control IS. Not what the line already says - a card that reads
   * `at 12,110 size 92x66` back as "at 12,110, 92 by 66 points" is standing in the way of the
   * text it is quoting (the owner, of the first cut: "the current hover information is
   * superfluous... that's obvious from the markdown").
   *
   * The coclass rode along here for a while and went the same way (the owner, 2026-08-16: "the
   * part on class rollover that says forms.optionbutton.1 says that on all of them... seems not
   * helpful"). For the fifteen standard kinds the ProgID IS the kind - `Forms.` plus the word
   * under the pointer plus `.1` - so every card carried a line that restated its own heading. It
   * stays only where it cannot be worked out, which is the case it was there for.
   */
  // Matched against the TOKEN rather than against the scanned kind, because the scan stops at the
  // caret: with the pointer in the middle of `<Frame` the element's name so far is "Fra", and
  // comparing to that answered nothing over the very word being hovered.
  if (kindNamed(token) !== null && !/"[^"]*$/.test(spot.before)) {
    const known = kindNamed(token);
    if (known === null) {
      return contents("Not a toolbox kind. An apply needs a `ProgId = \"...\"` line to create one.");
    }

    // The TYPE first, the way the name hover leads with a declaration (the owner, 2026-08-16:
    // "can you add the class type to the class hover?"). It is what a developer writes in a Dim,
    // and it is the one line about a kind that is not already on screen - unlike the ProgID,
    // which for the standard fifteen is this same word with a prefix and a suffix.
    const derivable = known.progId === `Forms.${known.kind}.1`;
    return contents(
      "```vba\n" + (known.kind === "Form" ? "MSForms.UserForm" : `MSForms.${known.kind}`) + "\n```",
      known.doc ?? null,
      known.progId && !derivable ? `\`${known.progId}\`` : null);
  }

  // THE NAME: the control declared the way VBA declares it, which is also what hovering the same
  // name in the code-behind answers - one identifier, one sentence about it, whichever half of
  // the tab the pointer is in.
  // Inside the Name attribute's own value. Same reason as the kind above: the scan cannot have
  // read the closing quote yet, so the element's name is still "" while the pointer is on it.
  if (/\bName\s*=\s*"[^"]*$/i.test(spot.before)) {
    const spelled = /<\s*([A-Za-z_]\w*)/.exec(spot.line)?.[1] ?? spot.header.kind;
    const known = kindNamed(spelled);
    const type = known === null ? spelled : `MSForms.${known.kind}`;
    return contents("```vba\n" + `${token} As ${type}` + "\n```", known?.doc ?? null);
  }

  return null;
}

/* -------------------------------------------------------------- the header hint */

/** The element's shape, one attribute per parameter, which is what the hint walks along. The
 * brackets are the documented spelling of "optional" and stay in the labels, so the hint reads
 * as the grammar rather than as a demand. A Form has no position of its own. */
const HEADER_PARAMETERS = [
  'Name="..."', '[Caption="..."]', '[Left="0" Top="0"]', '[Width="60" Height="20"]',
];
const FORM_PARAMETERS = ['Name="..."', '[Caption="..."]', '[Width="240" Height="180"]'];

export function headerHintAt(
  model: monaco.editor.ITextModel, position: monaco.IPosition,
): monaco.languages.SignatureHelpResult | null {
  const spot = spotAt(model, position);
  if (spot.header === null) {
    return null;
  }

  const form = spot.header.kind.toLowerCase() === "form";
  const parameters = form ? FORM_PARAMETERS : HEADER_PARAMETERS;

  // Which attribute the hand is on, followed by the attribute NAME standing rather than by
  // counting words: attributes may be typed in any order, so a count would drift the moment a
  // developer writes Width before Left.
  const beforeCaret = spot.before.toLowerCase();
  const wrote = (name: string) => new RegExp(`\\b${name}\\s*=[^=]*$`).test(beforeCaret);
  let active = 0;
  if (wrote("width") || wrote("height")) {
    active = parameters.length - 1;
  } else if (!form && (wrote("left") || wrote("top"))) {
    active = 2;
  } else if (wrote("caption")) {
    active = 1;
  }

  return {
    value: {
      // No documentation line: the label IS the grammar, and a sentence under it explaining
      // that a Form line is a form line is the kind of filler a hint has no room for.
      signatures: [{
        label: parameters.join(" "),
        parameters: parameters.map((parameter) => ({ label: parameter })),
      }],
      activeSignature: 0,
      activeParameter: active,
    },
    dispose: () => { },
  };
}

/* -------------------------------------------------------------- registration */

let registered = false;

/**
 * The language, once per page: its grammar, its indentation, and its service. Called by every
 * designer view's constructor and guarded, because the second form to open must not register a
 * second set of providers - monaco would ask both and merge the answers.
 */
export function registerMarkupLanguage(): void {
  if (registered) {
    return;
  }
  registered = true;

  monaco.languages.register({ id: FORM_MARKUP_LANGUAGE });

  monaco.languages.setMonarchTokensProvider(FORM_MARKUP_LANGUAGE, {
    // The toolbox kinds plus the two structural words; a type outside this list still reads
    // as an identifier, which is honest - the apply treats it as foreign too.
    controlKinds: [
      "Form", "Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton",
      "ToggleButton", "Frame", "CommandButton", "TabStrip", "MultiPage", "Page",
      "ScrollBar", "SpinButton", "Image",
    ],
    tokenizer: {
      root: [
        // A comment swallows everything to its close, across lines.
        [/<!--/, "comment", "@comment"],
        // An opening or closing tag hands the rest of the element to @tag, so an attribute list
        // that wraps across lines keeps painting as attributes rather than as loose words.
        // The `next` rides the LAST GROUP rather than sitting as a third element: with an array
        // action Monarch ignores a rule-level next, so @tag was never entered and every
        // attribute and string fell through to root uncoloured while the kind alone painted
        // (the owner, looking at the tab: "no colors").
        [/(<\/?)([A-Za-z_][\w]*)/, [
          "delimiter.angle",
          {
            cases: {
              "@controlKinds": { token: "type", next: "@tag" },
              "@default": { token: "identifier", next: "@tag" },
            },
          },
        ]],
        [/[<>]/, "delimiter.angle"],
      ],
      comment: [
        [/-->/, "comment", "@pop"],
        [/[^-]+/, "comment"],
        [/./, "comment"],
      ],
      tag: [
        [/\s+/, ""],
        [/\/?>/, "delimiter.angle", "@pop"],
        [/"(?:[^"]|"")*"/, "string"],
        [/"/, "string", "@value"],
        [/[A-Za-z_][\w.]*(?=\s*=)/, "attribute.name"],
        [/=/, "delimiter"],
        [/[A-Za-z_][\w.]*/, "identifier"],
      ],
      // A value whose closing quote is still to come - which is every value mid-typing, and the
      // state that keeps the rest of the line from painting as attributes behind the widget.
      value: [
        [/""/, "string"],
        [/"/, "string", "@pop"],
        [/[^"]+/, "string"],
      ],
    },
  });

  monaco.languages.setLanguageConfiguration(FORM_MARKUP_LANGUAGE, {
    // An unclosed container tag opens a level, the way the printer indents its children; a
    // self-closing one does not, which is why the rule looks for a `>` that no `/` precedes.
    onEnterRules: [{
      beforeText: /^\s*<(?:Frame|MultiPage|Page|TabStrip|Form)\b(?:[^>]|"[^"]*")*[^/]>\s*$/,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    }],
    brackets: [],
    autoClosingPairs: [{ open: '"', close: '"' }],
    comments: { blockComment: ["<!--", "-->"] },
  });

  monaco.languages.registerCompletionItemProvider(FORM_MARKUP_LANGUAGE, {
    // A dot reaches a font's members, an equals opens the value's side, and a space is where
    // a clause or a value begins. Typing a word triggers the widget on its own.
    triggerCharacters: [".", "=", " "],
    provideCompletionItems: (model, position) => completionsAt(model, position),
  });

  monaco.languages.registerHoverProvider(FORM_MARKUP_LANGUAGE, {
    provideHover: (model, position) => hoverAt(model, position),
  });

  monaco.languages.registerSignatureHelpProvider(FORM_MARKUP_LANGUAGE, {
    signatureHelpTriggerCharacters: [" ", ","],
    signatureHelpRetriggerCharacters: [" ", ",", "x"],
    provideSignatureHelp: (model, position) => headerHintAt(model, position),
  });
}
