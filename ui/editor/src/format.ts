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
import { canonicalKeyword } from "xlide-spec/analyzer/lexer/keywordTable";
import { tokenizeCached } from "xlide-spec/analyzer/lexer/tokenize";
import { CANONICAL_KEYWORDS, VBA_LANGUAGE_ID } from "./vba.js";

export interface FormatOptions {
  /** Spaces per indent level. */
  indentSize: number;
  /** Respell keywords in their canonical case. */
  canonicalKeywords: boolean;
}

export const DEFAULT_FORMAT_OPTIONS: FormatOptions = {
  indentSize: 4,
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
  // `EndIf` in one word closes a block If as `End If` does (MS-VBAL 5.4.2.8), and the VBE writes it
  // back as `End If`. Without it the If's body never closed, and everything after it down to the
  // End Sub was formatted one level too deep (upstream #88, measured 2026-09-23).
  /^endif\b/i,
  /^(?:next|loop|wend)\b/i,
];

/** Lines that step out for their own line and back in for the next. */
const MIDDLES = [
  /^(?:else|elseif)\b/i,
  /^case\b/i,
];

/** A label, which VBA puts hard against the left margin. */
/*
 * A LETTER MEANS ANY LETTER, here as everywhere else. See NAME_START below: this said
 * [A-Za-z_]\w* and so did not recognise `étape:` as a line label at all.
 */
const LABEL = /^[\p{L}_][\p{L}\p{M}\p{N}_]*:(?!=)/u;

/*
 * WHERE A WORD STARTS AND WHAT CONTINUES IT, for the keyword canonicaliser below.
 *
 * `\w` is ASCII in JavaScript however the rest is written, so walking a word with it stopped at
 * the first accented letter and handed the canonicaliser the ASCII HEAD of the name. When that
 * head is a keyword the name comes back respelled: measured 2026-08-09, Format Module turned
 * `inés` into `Inés` and `toé` into `Toé`, while `plainName` beside them was untouched. `Inés` is
 * a name, and `to`, `as`, `is`, `on`, `do` and `if` all start ordinary words.
 */
const NAME_START = /[\p{L}_]/u;
const NAME_PART = /[\p{L}\p{M}\p{N}_]/u;

/** Directives, which are never indented. `#EndIf` is one, as `#End If` is. */
const DIRECTIVE = /^#(?:if|elseif|else|end\s*if|const)\b/i;

/** A type-declaration character ending a name, which is no part of the word it spells. */
const TYPE_SUFFIX = /[%&^!#@$]$/;

/**
 * What the spec's own lexer reads in a module, line by line, numbered as `formatVba` splits them.
 *
 * THREE THINGS THIS FORMATTER USED TO WORK OUT BY ITSELF, AND GOT WRONG IN THE SAME WAY: it read
 * one physical line at a time, with no idea what the lines around it made of it.
 *
 * A comment ending in ` _` runs on through the next line (MS-VBAL 3.3.1), which the VBE takes as
 * comment text however much it looks like code. Formatted as code, `' note _` above `end sub`
 * respelled the comment's text to `End Sub` and closed the procedure there (upstream #82). A Rem
 * comment after a colon, `x = 1: rem if not`, had its words respelled because only an apostrophe
 * stopped the respeller. And `step`, `error`, `explicit` and `ptrsafe` are keywords only inside
 * their own statements - a For header, On Error, Option, a Declare - and names everywhere else, so
 * a variable called `step` came out as `Step` (upstream #86). All three measured 2026-09-23 on
 * upstream's 10.7.1 cases.
 *
 * The lexer the page already bundles answers all three the way the analyzer does, because it IS
 * the analyzer's: a comment token spans every line it carries, a comment starts wherever the
 * lexer began one, and a statement-bound keyword standing outside its statement comes back as an
 * identifier. So the formatter asks it rather than keeping a second opinion.
 */
interface LexedLines {
  /** Lines a comment carries on to: its text, left exactly as written. */
  carried: ReadonlySet<number>;
  /** Where each line's comment starts, an apostrophe's or a Rem's. */
  commentAt: ReadonlyMap<number, number>;
  /** Columns of words that spell a keyword but stand where the lexer reads them as a name. */
  names: ReadonlyMap<number, ReadonlySet<number>>;
}

function lexLines(text: string): LexedLines {
  const carried = new Set<number>();
  const commentAt = new Map<number, number>();
  const names = new Map<number, Set<number>>();

  for (const token of tokenizeCached(text)) {
    if (token.kind === "comment") {
      if (!commentAt.has(token.line)) {
        commentAt.set(token.line, token.character);
      }

      const physical = token.rawText.split(/\r\n|\r|\n/).length;
      for (let extra = 1; extra < physical; extra += 1) {
        carried.add(token.line + extra);
      }
    } else if (token.kind === "identifier"
      && canonicalKeyword(token.rawText.replace(TYPE_SUFFIX, "")) !== undefined) {
      const columns = names.get(token.line) ?? new Set<number>();
      columns.add(token.character);
      names.set(token.line, columns);
    }
  }

  return { carried, commentAt, names };
}

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

/**
 * `commentAt` is where the lexer found this line's comment, so a Rem is text as an apostrophe is;
 * `isName` says the lexer reads the word at a column as a name though it spells a keyword.
 */
function respell(line: string, commentAt?: number, isName?: (column: number) => boolean): string {
  // Identifiers only, and never inside a string or a comment. The callback receives each run of
  // word characters; anything not a keyword is returned untouched, which leaves every name the
  // developer chose exactly as they wrote it.
  let inString = false;
  let result = "";
  let index = 0;

  while (index < line.length) {
    if (commentAt !== undefined && index >= commentAt) {
      result += line.slice(index);
      break;
    }

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

    if (ch !== undefined && NAME_START.test(ch)) {
      let end = index;
      while (end < line.length && NAME_PART.test(line.charAt(end))) {
        end += 1;
      }

      const word = line.slice(index, end);

      // A word followed by a dot or preceded by one is a member name, not a keyword: `Range.Type`
      // and `Sheet.Cells` must keep whatever the object model calls them. So is a keyword the
      // lexer reads as a name where it stands: `Dim step As Long`.
      const isMember = line.charAt(index - 1) === "." || line.charAt(end) === ".";
      const canonical = isMember || isName?.(index) ? undefined : CANONICAL.get(word.toLowerCase());

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
  // Spaces, always. VBA's code store will not hold a tab: the editor expands every one it is
  // handed to the next four-column stop, so a module formatted with tabs read back as spaces and
  // the page and the project disagreed for as long as it stayed open.
  const unit = " ".repeat(Math.max(1, options.indentSize));
  const lines = text.split(/\r?\n/);
  const formatted: string[] = [];
  const lexed = lexLines(text);

  let depth = 0;
  // While a statement is continued, its following lines are indented one further and none of
  // them are examined for block structure: they are the middle of one statement.
  let continuing = false;

  for (const [index, original] of lines.entries()) {
    // A line a comment carries on to is the comment's text: not indented, not respelled, and not
    // a block's opener or closer, whatever it says. Exactly as written, which is what the VBE
    // keeps and what upstream's Format Document does since 10.7.1.
    if (lexed.carried.has(index)) {
      formatted.push(original);
      continuing = false;
      continue;
    }

    const body = original.trim();

    if (body.length === 0) {
      formatted.push("");
      continue;
    }

    // Structure is decided on the code BEFORE the comment, wherever the lexer began one: a
    // `: Rem` comment's words are not a block opener, and its ` _` is not a continuation.
    const lead = original.length - original.trimStart().length;
    const commentAt = lexed.commentAt.get(index);
    const code = significant(commentAt === undefined ? body : original.slice(0, commentAt).trim());
    const names = lexed.names.get(index);
    const respelled = options.canonicalKeywords
      ? respell(
        body,
        commentAt === undefined ? undefined : commentAt - lead,
        (column) => names?.has(lead + column) ?? false)
      : body;

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
 * Models with a formatting edit on its way in, so a change event can be told from a keystroke.
 *
 * WHY THE HOST NEEDS TO KNOW. A line being typed does not get squiggled until the caret leaves
 * it - the VBE's own contract, and the reason the analyzer does not complain about `MsgBox `
 * while its arguments are still on their way. The host arms that hold when a change event
 * touches exactly one line and inserts no newline, which is what a keystroke looks like.
 *
 * A format looks like that too, whenever it happens to change one line. Formatting a module
 * whose lines are already indented correctly except one does precisely that: the whole-document
 * edit goes in, the editor reduces it to the single line that actually moved, and the host reads
 * a keystroke. The findings on that line are then hidden until the caret is moved off it, so
 * running Format Module made a red squiggle disappear and stay gone (2026-08-07).
 *
 * The editor gives a change event no reason or source, so the only place that knows a format is
 * responsible is the formatter itself. Armed per model when there is an edit to apply, consumed
 * by the first change event that follows it, and dropped on a turn of the loop if none does.
 */
const formatting = new WeakSet<monaco.editor.ITextModel>();

/**
 * Whether this model's next change is a format's, clearing the mark as it answers. Asked once
 * per change event, so a format that produced an edit is credited to exactly one of them.
 */
export function takeFormattingMark(model: monaco.editor.ITextModel): boolean {
  if (!formatting.has(model)) {
    return false;
  }

  formatting.delete(model);
  return true;
}

/**
 * The edits a format wants, marked as a format's.
 *
 * The mark goes on only when there is something to apply - an already-formatted module produces
 * no change event, and a mark left standing would be spent on whatever the developer typed next.
 * The timer is the backstop for an edit the editor declines to apply at all.
 */
function marked(
  model: monaco.editor.ITextModel,
  edits: monaco.languages.TextEdit[],
): monaco.languages.TextEdit[] {
  if (edits.length > 0) {
    formatting.add(model);
    setTimeout(() => formatting.delete(model), 2000);
  }

  return edits;
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

      return marked(
        model,
        formatted === text ? [] : [{ range: model.getFullModelRange(), text: formatted }]);
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

      return marked(model, formatted === text ? [] : [{ range: whole, text: formatted }]);
    },
  });
}
