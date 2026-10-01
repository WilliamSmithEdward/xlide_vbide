// Properties of the page's readers of module text, checked over generated input with fast-check.
//
// Module text is whatever a workbook carries, and a workbook is often somebody else's. The
// formatter, the procedure scanner and the tokenizer all run over every keystroke of it, so each
// must hold its promise for text nobody wrote a case for, not only for the cases in format.mjs,
// procedureat.mjs and tokenizer.mjs. properties.mjs builds this file and runs it.
//
// The text is generated the way the page receives it: a monaco model's value, whose lines are
// joined by one end-of-line sequence and contain no other carriage return or line feed.

import fc from "fast-check";
import { formatVba, type FormatOptions } from "../src/format.js";
import { procedureAt, scanProcedures } from "../src/procedureat.js";
import { buildVbaMonarch } from "../src/vba.js";
import { compile } from "monaco-monarch/monarchCompile.js";
import { MonarchTokenizer } from "monaco-monarch/monarchLexer.js";

export interface Property {
  name: string;
  run: () => void;
}

/** Fragments that steer generated lines into the formatter's and scanner's decisions. */
const FRAGMENTS = [
  "Sub", "Function", "Property Get", "Property Let", "Property Set", "End Sub", "End Function",
  "End Property", "Public", "Private", "Static", "If", "Then", "Else", "ElseIf", "End If", "EndIf",
  "For", "To", "Step", "Next", "Do", "Loop", "While", "Wend", "With", "End With", "Select Case",
  "Case", "End Select", "Type", "End Type", "Enum", "End Enum", "Dim", "As", "Long", "step",
  "error", "Rem", "rem", "'", "\"", "\"\"", " _", "_", ":", ".", "#If", "#End If", "#EndIf",
  "#Const", "label:", "étape:", "inés", "x", "=", "1", "(", ")", ",", " ", "\t", " ",
];

/** One line: fragments, and characters of any kind except a line break. */
const line = fc
  .array(fc.oneof(
    { weight: 4, arbitrary: fc.constantFrom(...FRAGMENTS) },
    { weight: 1, arbitrary: fc.string({ unit: "binary", maxLength: 6 }) },
  ), { maxLength: 8 })
  .map((parts) => parts.join(" ").replace(/[\r\n]/g, ""));

const lines = fc.array(line, { maxLength: 40 });
const eol = fc.constantFrom("\r\n", "\n");
const moduleText = fc.tuple(lines, eol).map(([body, ending]) => body.join(ending));

const options: fc.Arbitrary<FormatOptions> = fc.record({
  indentSize: fc.integer({ min: 0, max: 8 }),
  canonicalKeywords: fc.boolean(),
});

const splitLines = (text: string): string[] => text.split(/\r?\n/);

export function properties(numRuns: number, seed?: number): Property[] {
  const settings = { numRuns, includeErrorInReport: true, ...(seed === undefined ? {} : { seed }) };
  const tokenizer = new MonarchTokenizer(
    null, null, "vba", compile("vba", buildVbaMonarch([], [])),
    { getValue: () => 20000, onDidChangeConfiguration: () => ({ dispose() {} }) },
  );

  return [
    {
      name: "formatting keeps every line, changing only indentation and the case of words",
      run: () => fc.assert(fc.property(moduleText, options, (text, chosen) => {
        const before = splitLines(text);
        const after = splitLines(formatVba(text, chosen));
        if (after.length !== before.length) {
          throw new Error(`${before.length} lines became ${after.length}`);
        }
        after.forEach((formatted, index) => {
          const original = before[index]!;
          // A line a comment carries on to is kept exactly; every other line keeps its text.
          if (formatted === original) {
            return;
          }
          const kept = chosen.canonicalKeywords
            ? formatted.trim().toLowerCase() === original.trim().toLowerCase()
            : formatted.trim() === original.trim();
          if (!kept) {
            throw new Error(`line ${index + 1} changed: ${JSON.stringify(original)} -> ${JSON.stringify(formatted)}`);
          }
        });
      }), settings),
    },
    {
      name: "formatting indents with whole units of spaces",
      run: () => fc.assert(fc.property(moduleText, options, (text, chosen) => {
        const unit = Math.max(1, chosen.indentSize);
        const before = splitLines(text);
        splitLines(formatVba(text, chosen)).forEach((formatted, index) => {
          if (formatted === before[index]) {
            return;
          }
          // Spaces and tabs are whitespace to the VBE; a no-break space is not, and stays.
          const lead = formatted.length - formatted.replace(/^[ \t]+/, "").length;
          if (!/^ *$/.test(formatted.slice(0, lead)) || lead % unit !== 0) {
            throw new Error(`line ${index + 1} is indented ${JSON.stringify(formatted.slice(0, lead))}`);
          }
        });
      }), settings),
    },
    {
      name: "formatting formatted text changes nothing",
      run: () => fc.assert(fc.property(moduleText, options, (text, chosen) => {
        const once = formatVba(text, chosen);
        const twice = formatVba(once, chosen);
        if (twice !== once) {
          throw new Error(`a second format changed ${JSON.stringify(once)} to ${JSON.stringify(twice)}`);
        }
      }), settings),
    },
    {
      name: "procedures tile the module from their first line to its last",
      run: () => fc.assert(fc.property(moduleText, (text) => {
        const count = text.split(/\r\n|\r|\n/).length;
        const found = scanProcedures(text);
        found.forEach((procedure, index) => {
          const previous = found[index - 1];
          const ok = procedure.start >= 1
            && procedure.start <= procedure.header
            && procedure.header <= procedure.end
            && (previous === undefined || procedure.start === previous.end + 1);
          if (!ok) {
            throw new Error(`procedure ${index} is ${JSON.stringify(procedure)} after ${JSON.stringify(previous)}`);
          }
        });
        const last = found[found.length - 1];
        if (last !== undefined && last.end !== count) {
          throw new Error(`the last procedure ends at ${last.end} of ${count} lines`);
        }
      }), settings),
    },
    {
      name: "the procedure at a line is the one whose lines include it",
      run: () => fc.assert(fc.property(moduleText, (text) => {
        const count = text.split(/\r\n|\r|\n/).length;
        const found = scanProcedures(text);
        for (let number = 0; number <= count + 1; number++) {
          const expected = found.find((procedure) => procedure.start <= number && number <= procedure.end) ?? null;
          if (procedureAt(found, number) !== expected) {
            throw new Error(`line ${number} answered ${JSON.stringify(procedureAt(found, number))}`);
          }
        }
      }), settings),
    },
    {
      name: "the tokenizer covers every line with tokens in order",
      run: () => fc.assert(fc.property(lines, (body) => {
        let state = tokenizer.getInitialState();
        body.forEach((text, index) => {
          const result = tokenizer.tokenize(text, true, state);
          const offsets: number[] = result.tokens.map((token: { offset: number }) => token.offset);
          // An empty line has no tokens; any other starts its first at column 0.
          const ordered = (text.length === 0 || offsets[0] === 0)
            && offsets.every((offset, at) => at === 0 || offset > offsets[at - 1]!)
            && offsets.every((offset) => offset <= Math.max(0, text.length - 1));
          if (!ordered) {
            throw new Error(`line ${index + 1} ${JSON.stringify(text)} has token offsets ${JSON.stringify(offsets)}`);
          }
          state = result.endState;
        });
      }), settings),
    },
  ];
}
