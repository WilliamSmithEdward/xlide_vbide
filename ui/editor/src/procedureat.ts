/*
 * Which procedure a line of a module is in, for the status bar and the tree's highlight (#23).
 *
 * The answer the editor itself gives is CodeModule.ProcOfLine, and this is held to it by the
 * folders suite, line by line, against the native editor: a readout that disagreed with the
 * procedure box the developer used to have would be worse than none. The rule that matches it:
 *
 *   - a procedure runs from its header to its End line, and takes the comment and blank lines
 *     ABOVE its header as well, back to the previous procedure's End line - a comment introducing
 *     a Sub belongs to that Sub, which is how the editor's own ProcStartLine counts;
 *   - the declarations section is everything above the first procedure's own leading comments:
 *     up to and including the last line there that is code rather than a comment or a blank;
 *   - lines after the last procedure's End belong to that procedure, since there is no next one
 *     to take them.
 *
 * Scanning is per line and cheap, but a 65,000-line module is scanned once per text version,
 * not once per caret move: the caller keeps the ranges and asks `procedureAt` for each line.
 */

export interface ProcedureRange {
  name: string;
  /** "Sub", "Function", "Property Get", "Property Let" or "Property Set". */
  kind: string;
  /** The line the header is on. */
  header: number;
  /** The first line the procedure claims, its leading comments included. */
  start: number;
  /** The last line it claims: its End line, or the line before the next procedure's start. */
  end: number;
}

// A letter of any script, then letters, digits, marks and underscores: VBA names are not ASCII.
const HEADER = /^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\p{L}[\p{L}\p{N}\p{M}_]*)/iu;
const END = /^\s*End\s+(?:Sub|Function|Property)\b/i;
const BLANK_OR_COMMENT = /^\s*(?:'|Rem\b|$)/i;

/** Splits on any line ending, the way the editor counts lines. */
function linesOf(text: string): string[] {
  return text.split(/\r\n|\r|\n/);
}

/** "Sub", "Function", "Property Get": each word capitalised, one space between. */
function kindOf(raw: string): string {
  return raw.split(/\s+/).map((word) => word[0]!.toUpperCase() + word.slice(1).toLowerCase()).join(" ");
}

/** The procedures of a module's text, in order, each with the lines it claims. */
export function scanProcedures(text: string): ProcedureRange[] {
  const lines = linesOf(text);
  const found: ProcedureRange[] = [];
  let open: ProcedureRange | null = null;

  // Assigned at the call sites rather than inside a closure, so the narrowing below holds.
  const openAt = (header: RegExpExecArray, number: number): ProcedureRange => {
    const range = { name: header[2]!, kind: kindOf(header[1]!), header: number, start: number, end: number };
    found.push(range);
    return range;
  };

  for (let at = 0; at < lines.length; at++) {
    const line = lines[at] ?? "";
    const number = at + 1;

    if (open === null) {
      const header = HEADER.exec(line);
      if (header) {
        open = openAt(header, number);
      }
      continue;
    }

    if (END.test(line)) {
      open.end = number;
      open = null;
      continue;
    }

    // A header inside an unclosed procedure: the text is mid-edit or broken. The new one takes
    // over from its own header, and the old one ends the line before.
    const header = HEADER.exec(line);
    if (header) {
      open.end = number - 1;
      open = openAt(header, number);
    }
  }

  // An unclosed last procedure runs to the end of the text.
  if (open !== null) {
    open.end = lines.length;
  }

  // Now the leading comments. The first procedure starts after the last code line of the
  // declarations; every later one starts right after its predecessor's End.
  for (let i = 0; i < found.length; i++) {
    const procedure = found[i]!;
    if (i === 0) {
      let start = procedure.header;
      while (start > 1 && BLANK_OR_COMMENT.test(lines[start - 2] ?? "")) {
        start -= 1;
      }
      procedure.start = start;
    } else {
      procedure.start = found[i - 1]!.end + 1;
    }
  }

  // And the trailing lines: everything up to the next procedure's start, or the end of the text.
  for (let i = 0; i < found.length; i++) {
    const procedure = found[i]!;
    const next = found[i + 1];
    procedure.end = next ? next.start - 1 : lines.length;
  }

  return found;
}

/** The procedure a 1-based line is in, or null in the declarations section. */
export function procedureAt(procedures: readonly ProcedureRange[], line: number): ProcedureRange | null {
  let low = 0;
  let high = procedures.length - 1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    const candidate = procedures[mid]!;
    if (line < candidate.start) {
      high = mid - 1;
    } else if (line > candidate.end) {
      low = mid + 1;
    } else {
      return candidate;
    }
  }
  return null;
}

/** The status bar's spelling: "Sub Recalculate", "Property Get Total". */
export function describeProcedure(procedure: ProcedureRange | null): string {
  return procedure ? `${procedure.kind} ${procedure.name}` : "";
}
