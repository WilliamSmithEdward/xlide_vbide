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

export const vbaLanguageConfiguration: monaco.languages.LanguageConfiguration = {
  comments: {
    lineComment: "'",
  },
  brackets: [["(", ")"]],
  autoClosingPairs: [
    { open: "(", close: ")" },
    { open: '"', close: '"', notIn: ["string", "comment"] },
  ],
  surroundingPairs: [
    { open: "(", close: ")" },
    { open: '"', close: '"' },
  ],
  // VBA identifiers are letters, digits and underscore; the trailing type-declaration
  // characters are deliberately excluded so double click selects the bare name.
  wordPattern: /[A-Za-z_][A-Za-z0-9_]*/g,

  indentationRules: {
    // Opens a block. Guards:
    //  - `Declare [PtrSafe] Function|Sub` is a one line declaration, so it is excluded.
    //  - `If ... Then <stmt>` is a one line If, so `Then` must end the line (a trailing
    //    comment is allowed).
    increaseIndentPattern:
      /^(?!.*\bDeclare\b)\s*(?:(?:Public|Private|Friend|Global)\s+)?(?:(?:Static\s+)?(?:Sub|Function)\b|Property\s+(?:Get|Let|Set)\b|(?:Type|Enum)\s+\w+|If\b.*\bThen\s*(?:'.*)?$|ElseIf\b.*\bThen\s*(?:'.*)?$|Else\s*(?:'.*)?$|For\b|Do\b|While\b|With\b|Select\s+Case\b|Case\b)/i,
    // Closes a block. Else/ElseIf/Case are in both patterns: the line itself pulls back one
    // level, the following lines indent again.
    decreaseIndentPattern: /^\s*(?:End\b|Next\b|Loop\b|Wend\b|Else\b|ElseIf\b|Case\b)/i,
  },
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
