import * as monaco from "monaco-editor/editor/editor.api.js";

export const THEME_DARK = "xlide-dark";
export const THEME_LIGHT = "xlide-light";

export type XlideTheme = typeof THEME_DARK | typeof THEME_LIGHT;

// Token rules are keyed on the scopes emitted by the VBA Monarch tokenizer. The `.vba`
// postfix is dropped so the same rule covers any future language that reuses the scope.
//
// The colours are the ones the companion editor's grammar renders as: keywords in one voice,
// declared and called names in another, types a third, the language's own values a fourth,
// ordinary identifiers a fifth. A module should read identically in the two products.
const darkRules: monaco.editor.ITokenThemeRule[] = [
  { token: "comment", foreground: "6a9955", fontStyle: "italic" },
  { token: "keyword", foreground: "c586c0" },
  { token: "keyword.option", foreground: "ce9178" },
  { token: "keyword.continuation", foreground: "808080" },
  { token: "constant", foreground: "569cd6" },
  { token: "type", foreground: "4ec9b8" },
  { token: "function", foreground: "dcdcaa" },
  { token: "identifier", foreground: "9cdcfe" },
  { token: "string", foreground: "ce9178" },
  { token: "string.escape", foreground: "d7ba7d" },
  { token: "number", foreground: "b5cea8" },
  { token: "number.date", foreground: "b5cea8" },
  { token: "operator", foreground: "d4d4d4" },
  { token: "delimiter", foreground: "d4d4d4" },
];

const lightRules: monaco.editor.ITokenThemeRule[] = [
  { token: "comment", foreground: "008000", fontStyle: "italic" },
  { token: "keyword", foreground: "af00db" },
  { token: "keyword.option", foreground: "a31515" },
  { token: "keyword.continuation", foreground: "808080" },
  { token: "constant", foreground: "0000ff" },
  { token: "type", foreground: "267f99" },
  { token: "function", foreground: "795e26" },
  { token: "identifier", foreground: "001080" },
  { token: "string", foreground: "a31515" },
  { token: "string.escape", foreground: "b06000" },
  { token: "number", foreground: "098658" },
  { token: "number.date", foreground: "098658" },
  { token: "operator", foreground: "1b1b1f" },
  { token: "delimiter", foreground: "1b1b1f" },
];

const darkTheme: monaco.editor.IStandaloneThemeData = {
  base: "vs-dark",
  inherit: true,
  rules: darkRules,
  colors: {
    "editor.background": "#1e1e1e",
    "editor.foreground": "#d4d4d4",
    "editorLineNumber.foreground": "#6e7681",
    "editorLineNumber.activeForeground": "#c6c6c6",
    "editorCursor.foreground": "#d4d4d4",
    "editor.selectionBackground": "#264f78",
    "editor.inactiveSelectionBackground": "#3a3d41",
    "editor.lineHighlightBackground": "#242424",
    "editor.lineHighlightBorder": "#00000000",
    "editorGutter.background": "#1e1e1e",
    "editorIndentGuide.background1": "#2f2f2f",
    "editorWhitespace.foreground": "#3b3b3b",
    "editorWidget.background": "#252526",
    "editorWidget.border": "#3c3c3c",
    "editorHoverWidget.background": "#252526",
    "scrollbarSlider.background": "#4e4e4e66",
    "scrollbarSlider.hoverBackground": "#5a5a5a99",
  },
};

const lightTheme: monaco.editor.IStandaloneThemeData = {
  base: "vs",
  inherit: true,
  rules: lightRules,
  colors: {
    "editor.background": "#ffffff",
    "editor.foreground": "#1b1b1f",
    "editorLineNumber.foreground": "#9a9aa6",
    "editorLineNumber.activeForeground": "#3b3b42",
    "editorCursor.foreground": "#1b1b1f",
    "editor.selectionBackground": "#add6ff",
    "editor.inactiveSelectionBackground": "#e5ebf1",
    "editor.lineHighlightBackground": "#f3f3f5",
    "editor.lineHighlightBorder": "#00000000",
    "editorGutter.background": "#ffffff",
    "editorIndentGuide.background1": "#e4e4e8",
    "editorWhitespace.foreground": "#d6d6dc",
    "editorWidget.background": "#f5f5f7",
    "editorWidget.border": "#d9d9de",
    "editorHoverWidget.background": "#f5f5f7",
    "scrollbarSlider.background": "#64646466",
    "scrollbarSlider.hoverBackground": "#64646499",
  },
};

export function defineThemes(): void {
  monaco.editor.defineTheme(THEME_DARK, darkTheme);
  monaco.editor.defineTheme(THEME_LIGHT, lightTheme);
}

// Dark is the default: only an explicit light preference flips it, so a browser or host that
// reports no preference lands on xlide-dark.
export function preferredTheme(): XlideTheme {
  const light = globalThis.matchMedia?.("(prefers-color-scheme: light)").matches === true;
  return light ? THEME_LIGHT : THEME_DARK;
}

export function watchPreferredTheme(onChange: (theme: XlideTheme) => void): void {
  const query = globalThis.matchMedia?.("(prefers-color-scheme: light)");
  query?.addEventListener("change", (event) => {
    onChange(event.matches ? THEME_LIGHT : THEME_DARK);
  });
}
