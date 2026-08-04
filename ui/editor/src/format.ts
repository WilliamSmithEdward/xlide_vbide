/*
 * Formatting for VBA: indentation, and the canonical spelling of keywords.
 *
 * Two rules, and only two, because everything else is a matter of taste and this runs over code
 * somebody else wrote. Blank lines, blank line runs, alignment inside a line and where a developer
 * chose to break an expression are all left exactly as found.
 *
 * Nothing here reflows a line. A line's leading whitespace is replaced and its keywords are
 * respelled; the rest of it is the developer's.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";
import { CANONICAL_KEYWORDS, VBA_LANGUAGE_ID } from "./vba.js";

export interface FormatOptions {
  /** Spaces per indent level. */
  indentSize: number;
  /** Indent with tabs rather than spaces. */
  useTabs: boolean;
  /** Respell keywords in their canonical case. */
  canonicalKeywords: boolean;
}

export const DEFAULT_FORMAT_OPTIONS: FormatOptions = {
  indentSize: 4,
  useTabs: true,
  canonicalKeywords: true,
};

/** Canonical spelling by lower-cased keyword, built once. */
const CANONICAL = new Map(CANONICAL_KEYWORDS.map((word) => [word.toLowerCase(), word]));

/*
 * The block openers and closers, as anchored patterns over a line with its indentation and any
 * trailing comment already removed.
 *
 * `If` is the awkward one: it opens a block only when `Then` ends the line. A single-line
 * `If x Then y = 1` opens nothing, and treating it as a block indents the whole rest of the
 * procedure.
 */
const OPENS = [
  /^(?:public\s+|private\s+|friend\s+|static\s+)*(?:sub|function)\b/i,
  /^(?:public\s+|private\s+|friend\s+|static\s+)*property\s+(?:get|let|set)\b/i,
  /^(?:public\s+|private\s+)?(?:type|enum)\b/i,
  /^if\b.*\bthen$/i,
  /^(?:for|do|while)\b/i,
  /^with\b/i,
  /^select\s+case\b/i,
];

const CLOSES = [
  /^end\s+(?:sub|function|property|if|with|type|enum|select)\b/i,
  /^(?:next|loop|wend)\b/i,
];

/** Lines that step out for their own line and back in for the next. */
const MIDDLES = [
  /^(?:else|elseif)\b/i,
  /^case\b/i,
];

/** A label, which VBA puts hard against the left margin. */
const LABEL = /^[A-Za-z_]\w*:(?!=)/;

/** Directives, which are never indented. */
const DIRECTIVE = /^#(?:if|elseif|else|end\s+if|const)\b/i;

/**
 * Splits a line into the part that decides indentation and the part that must not be examined.
 *
 * Strings and comments are removed for the decision only. Without this a comment saying "end if"
 * closes a block, and so does the string "For".
 */
function significant(line: string): string {
  let result = "";
  let inString = false;

  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];

    if (inString) {
      if (ch === '"') {
        // A doubled quote is an escaped quote and does not end the string.
        if (line[i + 1] === '"') {
          i += 1;
        } else {
          inString = false;
        }
      }
      continue;
    }

    if (ch === '"') {
      inString = true;
      continue;
    }

    if (ch === "'") {
      break;
    }

    result += ch;
  }

  const trimmed = result.trim();
  return /^rem\b/i.test(trimmed) ? "" : trimmed;
}

/** True when the line ends in a continuation, so the next line belongs to this statement. */
function continues(text: string): boolean {
  return /(^|\s)_$/.test(text.replace(/\s+$/, ""));
}

function respell(line: string): string {
  // Identifiers only, and never inside a string or a comment. The callback receives each run of
  // word characters; anything not a keyword is returned untouched, which leaves every name the
  // developer chose exactly as they wrote it.
  let inString = false;
  let result = "";
  let index = 0;

  while (index < line.length) {
    const ch = line[index];

    if (inString) {
      result += ch;
      if (ch === '"') {
        inString = line[index + 1] === '"';
        if (inString) {
          result += line[index + 1];
          index += 1;
        }
      }
      index += 1;
      continue;
    }

    if (ch === '"') {
      inString = true;
      result += ch;
      index += 1;
      continue;
    }

    if (ch === "'") {
      result += line.slice(index);
      break;
    }

    if (ch !== undefined && /[A-Za-z_]/.test(ch)) {
      let end = index;
      while (end < line.length && /\w/.test(line.charAt(end))) {
        end += 1;
      }

      const word = line.slice(index, end);

      // A word followed by a dot or preceded by one is a member name, not a keyword: `Range.Type`
      // and `Sheet.Cells` must keep whatever the object model calls them.
      const isMember = line.charAt(index - 1) === "." || line.charAt(end) === ".";
      const canonical = isMember ? undefined : CANONICAL.get(word.toLowerCase());

      result += canonical ?? word;
      index = end;
      continue;
    }

    result += ch;
    index += 1;
  }

  return result;
}

/** Formats a whole module. Returns the text unchanged when there is nothing to do. */
export function formatVba(text: string, options: FormatOptions = DEFAULT_FORMAT_OPTIONS): string {
  const unit = options.useTabs ? "\t" : " ".repeat(Math.max(1, options.indentSize));
  const lines = text.split(/\r?\n/);
  const formatted: string[] = [];

  let depth = 0;
  // While a statement is continued, its following lines are indented one further and none of
  // them are examined for block structure: they are the middle of one statement.
  let continuing = false;

  for (const original of lines) {
    const body = original.trim();

    if (body.length === 0) {
      formatted.push("");
      continue;
    }

    const code = significant(body);
    const respelled = options.canonicalKeywords ? respell(body) : body;

    if (continuing) {
      formatted.push(unit.repeat(depth + 1) + respelled);
      continuing = continues(code);
      continue;
    }

    if (LABEL.test(code) || DIRECTIVE.test(code)) {
      // Labels and directives sit at the margin, and neither changes the depth.
      formatted.push(respelled);
      continuing = continues(code);
      continue;
    }

    const closes = CLOSES.some((pattern) => pattern.test(code));
    const middle = MIDDLES.some((pattern) => pattern.test(code));

    if (closes || middle) {
      depth = Math.max(0, depth - 1);
    }

    formatted.push(unit.repeat(depth) + respelled);

    if (middle || OPENS.some((pattern) => pattern.test(code))) {
      depth += 1;
    }

    continuing = continues(code);
  }

  // The line ending is preserved: a module written with one convention should not change to the
  // other because it was formatted.
  return formatted.join(text.includes("\r\n") ? "\r\n" : "\n");
}

/**
 * Registers formatting for VBA.
 *
 * Both a whole-document and a range provider, because the editor uses different ones for "format
 * document" and "format selection" and offering only the first makes the second silently do
 * nothing.
 */
export function registerFormatting(getOptions: () => FormatOptions): void {
  monaco.languages.registerDocumentFormattingEditProvider(VBA_LANGUAGE_ID, {
    provideDocumentFormattingEdits(model) {
      const text = model.getValue();
      const formatted = formatVba(text, getOptions());

      return formatted === text ? [] : [{ range: model.getFullModelRange(), text: formatted }];
    },
  });

  monaco.languages.registerDocumentRangeFormattingEditProvider(VBA_LANGUAGE_ID, {
    provideDocumentRangeFormattingEdits(model, range) {
      // Whole lines, always. Formatting half a line would replace its indentation with the
      // indentation of a statement that starts somewhere the range does not include.
      const whole = new monaco.Range(
        range.startLineNumber,
        1,
        range.endLineNumber,
        model.getLineMaxColumn(range.endLineNumber));

      const text = model.getValueInRange(whole);
      const formatted = formatVba(text, getOptions());

      return formatted === text ? [] : [{ range: whole, text: formatted }];
    },
  });
}
