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

// The classification below is the companion editor's TextMate grammar, re-expressed as a
// Monarch tokenizer, and its colours in theme.ts are what that grammar renders as over there:
// every keyword one colour, declaration names and calls another, types a third, the language's
// literal values a fourth, ordinary identifiers a fifth. The two products should read the same
// module the same way at a glance.

/** Every word the grammar paints as a keyword, canonical list included. */
const TOKEN_KEYWORDS: string[] = [
  ...CANONICAL_KEYWORDS,
  "Eqv", "Imp", "TypeOf", "AddressOf", "Lib", "Alias", "Spc", "Any", "Shared",
  "Base", "Compare", "Attribute", "Write", "Seek", "Lock", "Unlock", "Put", "Open", "Close",
  "Input", "Print", "LSet", "RSet", "Wend", "GoSub", "Return",
  "DefBool", "DefByte", "DefCur", "DefDate", "DefDbl", "DefDec", "DefInt", "DefLng",
  "DefLngLng", "DefLngPtr", "DefObj", "DefSng", "DefStr", "DefVar",
];

/** The built-in type names, painted as types wherever they appear. */
const BUILTIN_TYPES: string[] = [
  "Boolean", "Byte", "Integer", "Long", "LongLong", "LongPtr", "Single", "Double",
  "Currency", "Date", "String", "Variant", "Object",
];

/** Literal values and intrinsic constants, painted as constants. */
const LANGUAGE_CONSTANTS: string[] = [
  "True", "False", "Nothing", "Null", "Empty",
  "vbCrLf", "vbCr", "vbLf", "vbNewLine", "vbTab", "vbNullString", "vbNullChar", "vbBack",
  "vbFormFeed", "vbVerticalTab", "vbBlack", "vbRed", "vbGreen", "vbYellow", "vbBlue",
  "vbMagenta", "vbCyan", "vbWhite", "vbOKOnly", "vbOKCancel", "vbAbortRetryIgnore",
  "vbYesNoCancel", "vbYesNo", "vbRetryCancel", "vbCritical", "vbQuestion", "vbExclamation",
  "vbInformation", "vbDefaultButton1", "vbDefaultButton2", "vbDefaultButton3",
  "vbApplicationModal", "vbSystemModal", "vbOK", "vbCancel", "vbAbort", "vbRetry", "vbIgnore",
  "vbYes", "vbNo", "vbString", "vbBinaryCompare", "vbTextCompare",
  "xlUp", "xlDown", "xlToLeft", "xlToRight", "xlValues", "xlFormulas", "xlComments",
  "xlWhole", "xlPart", "xlByRows", "xlByColumns", "xlNext", "xlPrevious",
  "xlCellTypeLastCell", "xlCellTypeBlanks", "xlShiftUp", "xlShiftToLeft", "xlShiftDown",
  "xlShiftToRight",
];

/** Names whose value is the language itself: `Me`. Painted with the constants. */
const LANGUAGE_VALUES: string[] = ["Me"];

/** The VBA standard library's functions, painted as calls wherever they appear. */
const BUILTIN_FUNCTIONS: string[] = [
  "Abs", "Array", "Asc", "AscB", "AscW", "Atn", "CBool", "CByte", "CCur", "CDate", "CDbl",
  "CDec", "Choose", "Chr", "ChrB", "ChrW", "CInt", "CLng", "CLngLng", "CLngPtr", "Cos",
  "CreateObject", "CSng", "CStr", "CurDir", "CVar", "CVErr", "DateAdd", "DateDiff", "DatePart",
  "DateSerial", "DateValue", "Day", "DDB", "Dir", "DoEvents", "Environ", "EOF", "Err", "Exp",
  "FileAttr", "FileCopy", "FileDateTime", "FileLen", "Filter", "Fix", "Format",
  "FormatCurrency", "FormatDateTime", "FormatNumber", "FormatPercent", "FreeFile", "FV",
  "GetAttr", "GetObject", "GetSetting", "Hex", "Hour", "IIf", "IMEStatus", "InputBox", "InStr",
  "InStrB", "InStrRev", "Int", "IPmt", "IRR", "IsArray", "IsDate", "IsEmpty", "IsError",
  "IsMissing", "IsNull", "IsNumeric", "IsObject", "Join", "LBound", "LCase", "Left", "LeftB",
  "Len", "LenB", "Loc", "LOF", "Log", "LTrim", "Mid", "MidB", "Minute", "MIRR", "Month",
  "MonthName", "MsgBox", "Now", "NPer", "NPV", "Oct", "Partition", "Pmt", "PPmt", "PSet", "PV",
  "QBColor", "Randomize", "Rate", "Replace", "RGB", "Right", "RightB", "Rnd", "Round", "RTrim",
  "Scale", "Second", "Sgn", "Shell", "Sin", "SLN", "Space", "Split", "Sqr", "Str", "StrComp",
  "StrConv", "StrReverse", "Switch", "SYD", "Tan", "Time", "Timer", "TimeSerial", "TimeValue",
  "Trim", "TypeName", "UBound", "UCase", "Val", "VarType", "Weekday", "WeekdayName", "Year",
];

/**
 * The tokenizer, built around what the project itself declares. The companion editor gets this
 * knowledge from its semantic tokens; here the engine sends the project's words — names that
 * denote types, names that denote procedures — and the tokenizer is rebuilt around them, which
 * is what lets `ROneCOne.Create(...)` read as a type and a call while `values(index, 1)` stays
 * a variable.
 */
function buildVbaMonarch(
  projectTypes: readonly string[],
  projectProcedures: readonly string[],
): monaco.languages.IMonarchLanguage {
  return {
  ignoreCase: true,
  defaultToken: "",
  tokenPostfix: ".vba",
  keywords: TOKEN_KEYWORDS,
  builtinTypes: BUILTIN_TYPES,
  languageConstants: LANGUAGE_CONSTANTS,
  languageValues: LANGUAGE_VALUES,
  builtinFunctions: BUILTIN_FUNCTIONS,
  projectTypes: projectTypes as string[],
  projectProcedures: projectProcedures as string[],

  tokenizer: {
    root: [
      // Rem is a statement level comment. It is matched before identifiers so `Rem` never
      // tokenizes as a plain word; `\b` keeps `Remaining` out of it.
      [/rem\b.*$/, "comment"],
      [/'.*$/, "comment"],

      // Conditional compilation directives, before the date literal can eat their #.
      [/^(\s*)(#\s*(?:If|ElseIf|Else|End\s+If|EndIf|Const))\b/, ["", "keyword"]],

      // Date literal. Bounded to a single line.
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

      // Option's argument has its own colour, the way the grammar scopes it apart.
      [/(Option)(\s+)(Explicit|Base|Compare|Private\s+Module)\b/, ["keyword", "", "keyword.option"]],

      // Declaration names: what follows the declaring keyword is the declared thing, and it
      // is painted as one. Order matters: these run before the general word rule below.
      [/\b(Sub|Function)(\s+)([A-Za-z_]\w*)/, ["keyword", "", "function"]],
      [/\b(Property)(\s+)(Get|Let|Set)(\s+)([A-Za-z_]\w*)/, ["keyword", "", "keyword", "", "function"]],
      [/\b(Event)(\s+)([A-Za-z_]\w*)/, ["keyword", "", "function"]],
      [/\b(Implements)(\s+)([A-Za-z_]\w*)/, ["keyword", "", "type"]],
      [/\b(Type|Enum)(\s+)([A-Za-z_]\w*)/, ["keyword", "", "type"]],

      // What follows As or New names a type, built-in or the project's own.
      [/\b(As|New)(\s+)([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*)/, ["keyword", "", "type"]],

      // A statement that starts with a bare name is a call: `InternalBeginAsyncOpen`,
      // `RequireDbConnection Me, "..."`. A name being assigned, labelled, or dotted is not.
      [/^(\s*)([A-Za-z_]\w*)(?![\w])(?!\s*[=:.(])/, ["", {
        cases: {
          "@keywords": "keyword",
          "@languageConstants": "constant",
          "@languageValues": "constant",
          "@builtinTypes": "type",
          "@builtinFunctions": "function",
          "@default": "function",
        },
      }]],

      // A member with an argument list is a call; a plain member is a member. The grammar's
      // quirk is kept: a member spelled like a keyword paints as one (`.Open`, `.Print`).
      [/(\.)([A-Za-z_]\w*)(?=\s*\()/, ["delimiter", {
        cases: { "@keywords": "keyword", "@default": "function" },
      }]],
      [/(\.)([A-Za-z_]\w*)/, ["delimiter", {
        cases: { "@keywords": "keyword", "@languageConstants": "constant", "@default": "identifier" },
      }]],

      // A name with an argument list is a call when the project (or the runtime) declares a
      // procedure by that name, and stays a variable otherwise: array indexing wears the same
      // parentheses, and `values(index, 1)` is data, not a call.
      [/([A-Za-z_]\w*)(?=\s*\()/, {
        cases: {
          "@keywords": "keyword",
          "@languageConstants": "constant",
          "@languageValues": "constant",
          "@builtinTypes": "type",
          "@projectTypes": "type",
          "@builtinFunctions": "function",
          "@projectProcedures": "function",
          "@default": "identifier",
        },
      }],

      // Every word that remains: constants, types, built-ins, keywords, then identifiers.
      [/[a-zA-Z_]\w*/, {
        cases: {
          "@languageConstants": "constant",
          "@languageValues": "constant",
          "@builtinTypes": "type",
          "@projectTypes": "type",
          "@builtinFunctions": "function",
          "@projectProcedures": "function",
          "@keywords": "keyword",
          "@default": "identifier",
        },
      }],

      // A foreign name is an identifier wearing brackets.
      [/\[[^\]\n]+\]/, "identifier"],

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
}

/** The words the tokenizer was last built around, so an unchanged push costs nothing. */
let appliedFactsKey = "\0";

/**
 * Rebuilds the tokenizer around the project's words. Open models re-tokenize on registration —
 * the whole module, on this thread — so the rebuild happens only when the words actually
 * changed: the lists arrive after every analysis pass, and they are almost always the same.
 */
export function updateVbaLanguageFacts(
  types: readonly string[],
  procedures: readonly string[],
): void {
  const key = `${types.join("\n")}\0${procedures.join("\n")}`;
  if (key === appliedFactsKey) {
    return;
  }

  appliedFactsKey = key;
  monaco.languages.setMonarchTokensProvider(VBA_LANGUAGE_ID, buildVbaMonarch(types, procedures));
}

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
  monaco.languages.setMonarchTokensProvider(VBA_LANGUAGE_ID, buildVbaMonarch([], []));
  monaco.languages.setLanguageConfiguration(VBA_LANGUAGE_ID, vbaLanguageConfiguration);
}
