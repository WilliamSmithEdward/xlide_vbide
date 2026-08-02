import * as monaco from "monaco-editor/editor/editor.api.js";

export const VBA_LANGUAGE_ID = "vba";

// VBA is case insensitive. The Monarch tokenizer is told so with `ignoreCase`, which also makes
// the `@keywords` case-block match without lowering every entry here.
/**
 * Canonical spelling of every keyword.
 *
 * VBA is case insensitive and its own editor respells keywords as they are typed, so a module
 * that has been through it is consistent. This list is what "consistent" means here, and the
 * formatter uses it as well as the tokenizer.
 */
export const CANONICAL_KEYWORDS: readonly string[] = [
  "Sub", "Function", "End", "If", "Then", "Else", "ElseIf", "For", "Next", "Do", "Loop",
  "While", "Wend", "Select", "Case", "Dim", "Set", "Let", "Const", "Public", "Private",
  "Friend", "Static", "Option", "Explicit", "On", "Error", "Resume", "GoTo", "With",
  "Property", "Get", "Class", "Type", "Enum", "Declare", "PtrSafe", "ByVal", "ByRef",
  "Optional", "ParamArray", "As", "New", "Nothing", "True", "False", "Null", "Empty",
  "And", "Or", "Not", "Xor", "Mod", "Is", "Like", "To", "Step", "Each", "In", "Exit",
  "Call", "RaiseEvent", "Implements", "WithEvents", "Preserve", "ReDim", "Erase", "Stop",
  "Debug", "Me",
];

export const vbaMonarchLanguage: monaco.languages.IMonarchLanguage = {
  ignoreCase: true,
  defaultToken: "",
  tokenPostfix: ".vba",
  keywords: CANONICAL_KEYWORDS as string[],

  tokenizer: {
    root: [
      // Rem is a statement level comment. It is matched before identifiers so `Rem` never
      // tokenizes as a plain word; `\b` keeps `Remaining` out of it.
      [/rem\b.*$/, "comment"],
      [/'.*$/, "comment"],

      // Date literal. Bounded to a single line so conditional compilation directives
      // (`#If Win64 Then`) fall through to the keyword rules instead of being eaten.
      [/#[^#\n]*#/, "number.date"],

      // Radix prefixes must precede the `&` operator rule and the plain number rules.
      [/&[hH][0-9a-fA-F]+&?/, "number.hex"],
      [/&[oO][0-7]+&?/, "number.octal"],

      [/[ \t\r\n]+/, ""],

      // Line continuation: whitespace already consumed above, so the underscore sits at the
      // current offset and only counts when nothing but whitespace follows it.
      [/_[ \t]*$/, "keyword.continuation"],

      // Floats before integers, exponents before both. VBA allows D as an exponent marker
      // (Double literals) and the trailing type-declaration characters !#@&%.
      [/\d*\.\d+(?:[eEdD][-+]?\d+)?[!#@&%]?/, "number.float"],
      [/\d+\.(?![.\w])(?:[eEdD][-+]?\d+)?[!#@&%]?/, "number.float"],
      [/\d+[eEdD][-+]?\d+[!#@&%]?/, "number.float"],
      [/\d+[!#@&%]?/, "number"],

      [/[a-zA-Z_]\w*/, { cases: { "@keywords": "keyword", "@default": "identifier" } }],

      [/"/, { token: "string.quote", bracket: "@open", next: "@string" }],

      [/[()]/, "@brackets"],
      [/[<>]=?|[-+*/\\^&=]|<>/, "operator"],
      [/[,;:.]/, "delimiter"],
    ],

    string: [
      [/[^"]+/, "string"],
      // "" is the escaped quote; it must be tested before the closing quote.
      [/""/, "string.escape"],
      [/"/, { token: "string.quote", bracket: "@close", next: "@pop" }],
    ],
  },
};

// The configuration below is the extension's own language-configuration JSON, transliterated:
// its patterns spell case-insensitivity as [Xx] character classes because JSON has no flags,
// and here they are the same expressions with /i. The behaviours are a contract shared with
// Smart Enter, and they should only change together.
export const vbaLanguageConfiguration: monaco.languages.LanguageConfiguration = {
  comments: {
    lineComment: "'",
  },
  brackets: [
    ["(", ")"],
    ["[", "]"],
  ],
  autoClosingPairs: [
    { open: "(", close: ")" },
    { open: "[", close: "]" },
    { open: '"', close: '"', notIn: ["string"] },
  ],
  surroundingPairs: [
    { open: "(", close: ")" },
    { open: "[", close: "]" },
    { open: '"', close: '"' },
  ],
  // VBA identifiers are letters, digits and underscore; the trailing type-declaration
  // characters are deliberately excluded so double click selects the bare name.
  wordPattern: /[A-Za-z_][A-Za-z0-9_]*/g,

  indentationRules: {
    // Opens a block. `Declare [PtrSafe] Function` never matches: the line starts with Declare,
    // which is not among the modifiers. A one-line `If ... Then <stmt>` never matches: Then
    // must end the line, give or take a trailing comment.
    increaseIndentPattern: new RegExp(
      "^[ \\t]*(?:(?:Public|Private|Friend|Global|Static)[ \\t]+)*(?:Sub|Function)[ \\t]+\\w"
      + "|^[ \\t]*(?:(?:Public|Private|Friend)[ \\t]+)*Property[ \\t]+(?:Get|Let|Set)[ \\t]+\\w"
      + "|^[ \\t]*If[ \\t]+\\S.+\\bThen[ \\t]*(?:'.*)?$"
      + "|^[ \\t]*(?:For[ \\t]+Each[ \\t]+[A-Za-z_][A-Za-z0-9_]*[ \\t]+In[ \\t]+\\S.*"
      + "|For[ \\t]+[A-Za-z_][A-Za-z0-9_]*[ \\t]*=.*\\bTo\\b.+)[ \\t]*(?:'.*)?$"
      + "|^[ \\t]*Do(?:[ \\t]+(?:While|Until)[ \\t]+\\S.*)?[ \\t]*(?:'.*)?$"
      + "|^[ \\t]*While[ \\t]+\\S.*[ \\t]*(?:'.*)?$"
      + "|^[ \\t]*With[ \\t]+\\S.*[ \\t]*(?:'.*)?$"
      + "|^[ \\t]*Select[ \\t]+Case[ \\t]+\\S.*[ \\t]*(?:'.*)?$"
      + "|^[ \\t]*(?:Type|Enum)[ \\t]+\\w"
      + "|^[ \\t]*#[ \\t]*If[ \\t]+\\S.+\\bThen[ \\t]*(?:'.*)?$",
      "i",
    ),
    // Closes a block. Else/ElseIf/Case pull their own line back; the on-enter rules indent the
    // lines that follow them again.
    decreaseIndentPattern: new RegExp(
      "^[ \\t]*(?:End[ \\t]+(?:Sub|Function|Property|If|With|Select|Type|Enum)"
      + "|#[ \\t]*End[ \\t]*If\\b|Next\\b|Loop\\b|Wend\\b|Else\\b|ElseIf\\b|Case\\b)",
      "i",
    ),
  },

  folding: {
    markers: {
      start: new RegExp(
        "^[ \\t]*(?:(?:Public|Private|Friend|Global|Static)[ \\t]+)*(?:Sub|Function)[ \\t]+\\w"
        + "|^[ \\t]*(?:(?:Public|Private|Friend)[ \\t]+)*Property[ \\t]+(?:Get|Let|Set)[ \\t]+\\w"
        + "|^[ \\t]*(?:(?:Public|Private|Global)[ \\t]+)*(?:Type|Enum)[ \\t]+\\w",
        "i",
      ),
      end: new RegExp("^[ \\t]*End[ \\t]+(?:Sub|Function|Property|Type|Enum)\\b", "i"),
    },
  },

  onEnterRules: [
    {
      beforeText: /^[ \t]*If[ \t]+\S.+\bThen[ \t]*(?:'.*)?$/i,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    },
    {
      beforeText: /^[ \t]*Else(?:If\b.*\bThen)?[ \t]*(?:'.*)?$/i,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    },
    {
      beforeText: /^[ \t]*Case\b.+$/i,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    },
    {
      beforeText: /^[ \t]*(?:For[ \t]+Each[ \t]+[A-Za-z_][A-Za-z0-9_]*[ \t]+In[ \t]+\S.*|For[ \t]+[A-Za-z_][A-Za-z0-9_]*[ \t]*=.*\bTo\b.+|Do(?:[ \t]+(?:While|Until)[ \t]+\S.*)?|While[ \t]+\S.*|With[ \t]+\S.*)[ \t]*(?:'.*)?$/i,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    },
    {
      beforeText: /^[ \t]*Select[ \t]+Case[ \t]+\S.*$/i,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    },
    {
      beforeText: /^[ \t]*(?:(?:Public|Private|Friend|Global|Static)[ \t]+)*(?:Sub|Function)[ \t]+\w.*$/i,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    },
    {
      beforeText: /^[ \t]*(?:(?:Public|Private|Friend)[ \t]+)*Property[ \t]+(?:Get|Let|Set)[ \t]+\w.*$/i,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    },
    {
      beforeText: /^[ \t]*(?:Type|Enum)[ \t]+\w.*$/i,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    },
    {
      beforeText: /^[ \t]*#[ \t]*If[ \t]+\S.+\bThen[ \t]*(?:'.*)?$/i,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    },
    {
      // A line continuation: the underscore at the end of a line, after a space.
      beforeText: /.*[ \t]_$/,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    },
  ],
};

export function registerVba(): void {
  const known = monaco.languages.getLanguages().some((lang) => lang.id === VBA_LANGUAGE_ID);
  if (known) {
    return;
  }

  monaco.languages.register({
    id: VBA_LANGUAGE_ID,
    extensions: [".bas", ".cls", ".frm", ".vba"],
    aliases: ["VBA", "vba", "Visual Basic for Applications"],
  });
  monaco.languages.setMonarchTokensProvider(VBA_LANGUAGE_ID, vbaMonarchLanguage);
  monaco.languages.setLanguageConfiguration(VBA_LANGUAGE_ID, vbaLanguageConfiguration);
}
