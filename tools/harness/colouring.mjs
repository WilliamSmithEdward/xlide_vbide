/*
 * What the editor PAINTS, asked of the rendered span rather than derived from a token type.
 *
 * WHY THIS EXISTS. Semantic colouring has a tokenizer of its own, rebuilt per project from the
 * words the engine sends, and until 2026-08-09 nothing checked a single colour it produced. Two
 * defects were found by eye on the same day:
 *
 *   a name with an accented letter was painted in TWO PIECES, because every declaration rule
 *   spelled the identifier as [A-Za-z_]\w* and stopped at the accent: `RécordAccent` rendered
 *   teal as far as the é and light blue after it.
 *
 *   one procedure was painted two ways: `Helpers.Total(1)` read as a call and
 *   `Helpers.Recalculate "x"` as an ordinary identifier, though a VBA Sub takes its arguments
 *   without parentheses and that is how every Sub call looks.
 *
 * Both were invisible to every other suite, because a colour is not a position, a finding or a
 * text. The colour is read off the rendered span: what a developer means by "is this the wrong
 * colour" is the pixel, and every step between the tokeniser and the pixel can be the wrong one.
 *
 *   node tools\harness\colouring.mjs
 */
import { open } from "file:///F:/GitHub/xlide/xlide_vbide/tools/harness/xlide-api.mjs";

const api = await open();
const project = await api.project();
const name = `Colour${process.pid}`;

let passed = 0;
const failures = [];
const check = (what, ok, detail) => {
  if (ok) { passed += 1; console.log(`ok   ${what}`); }
  else { failures.push(what); console.log(`FAIL ${what}${detail ? `\n     ${detail}` : ""}`); }
};

const CALL = "rgb(220, 220, 170)";
const PLAIN = "rgb(156, 220, 254)";
const TYPE = "rgb(78, 201, 184)";
const KEYWORD = "rgb(197, 134, 192)";
const NAMES = { [CALL]: "call", [PLAIN]: "identifier", [TYPE]: "type", [KEYWORD]: "keyword" };
const nameOf = (colour) => NAMES[colour] ?? colour;

const lines = [
  "Option Explicit",
  "",
  "Public Type PlainRecord",
  "    Field As Long",
  "End Type",
  "",
  "Public Type RécordAccent",
  "    Field As Long",
  "End Type",
  "",
  "Public Sub Recalculate(ByVal label As String)",
  "    Debug.Print label",
  "End Sub",
  "",
  "Public Function CalculérName(ByVal n As Long) As Long",
  "    CalculérName = n",
  "End Function",
  "",
  "Public Sub Caller()",
  "    Recalculate \"bare\"",
  `    ${name}.Recalculate "qualified"`,
  `    Debug.Print ${name}.CalculérName(1)`,
  "    Debug.Print Application.Version",
  "End Sub",
  "",
];

/** The colour at a word's first and last character, on the line that contains the given text. */
async function across(lineText, word) {
  const line = lines.findIndex((one) => one === lineText) + 1;
  const start = lines[line - 1].lastIndexOf(word) + 1;
  const first = await api.ui({ line, column: start });
  const last = await api.ui({ line, column: start + word.length - 1 });
  return {
    word: first.at?.word,
    head: first.at?.colour ?? "(none)",
    tail: last.at?.colour ?? "(none)",
    oneColour: first.at?.colour === last.at?.colour,
  };
}

let made = false;
try {
  await api.component("add", { kind: "module", name, project: project.projectId });
  made = true;
  await api.writeModule(name, lines.join("\r\n"), project.projectId);
  await new Promise((r) => setTimeout(r, 2000));
  await api.pane("open", { module: name, project: project.projectId });
  await new Promise((r) => setTimeout(r, 3000));

  const held = (await api.readModule(name, project.projectId)).text ?? "";
  const accented = held.includes("CalculérName") && held.includes("RécordAccent");
  if (!accented) {
    console.log("     (this machine's code page did not keep the accented names; those rows are skipped)\n");
  }

  // 1. ONE PROCEDURE, ONE COLOUR, however it is reached.
  const bare = await across("    Recalculate \"bare\"", "Recalculate");
  check("a bare call to a project procedure is painted as a call",
    bare.head === CALL, `it is ${nameOf(bare.head)}`);

  const qualified = await across(`    ${name}.Recalculate "qualified"`, "Recalculate");
  check("and the SAME procedure qualified, without parentheses, is too",
    qualified.head === CALL,
    `it is ${nameOf(qualified.head)}. A VBA Sub takes its arguments without parentheses, so this `
    + "is what every Sub call across modules looks like.");

  // 2. AND A LIBRARY MEMBER IS NOT, which is what stops the rule above painting every member.
  const library = await across("    Debug.Print Application.Version", "Version");
  check("a library member is still an ordinary identifier",
    library.head === PLAIN,
    `it is ${nameOf(library.head)}; the project-procedure rule has started painting things it `
    + "does not know.");

  /*
   * 3. THE INTRINSIC OBJECTS ARE OBJECTS, and the keyword beside one is still a keyword.
   *
   * `Debug` was painted as a keyword because it rides in the canonical list, which is a SPELLING
   * list: the formatter is there to turn `debug.print` into `Debug.Print`. Colouring read it as a
   * claim about the word's kind. The companion grammar keeps `Debug` out of all four of its
   * keyword patterns and gives it a rule of its own.
   *
   * `Print` beside it stays a keyword, and that is not an oversight in either product: `Print #1,`
   * is a real VBA statement, and the companion's keyword pattern is ordered ahead of its member
   * pattern, so it wins there too. The pair is checked together because the interesting thing is
   * that they differ.
   */
  const intrinsic = await across("    Debug.Print label", "Debug");
  check("the intrinsic Debug object is an object, not a keyword",
    intrinsic.head === PLAIN && intrinsic.oneColour,
    `it is ${nameOf(intrinsic.head)}. It rides in the canonical list because the formatter respells `
    + "it, and that list is about spelling, not about what a word is.");

  const statementWord = await across("    Debug.Print label", "Print");
  check("and Print beside it is still a keyword, as it is upstream",
    statementWord.head === KEYWORD,
    `it is ${nameOf(statementWord.head)}. Print is a VBA statement in its own right and the `
    + "companion grammar orders keywords ahead of members, so a member spelled like one paints "
    + "as one in both products.");

  // 4. A NAME IS ONE COLOUR ALL THE WAY ALONG.
  const plainType = await across("Public Type PlainRecord", "PlainRecord");
  check("a type name is a type, throughout", plainType.head === TYPE && plainType.oneColour,
    `head ${nameOf(plainType.head)}, tail ${nameOf(plainType.tail)}`);

  if (accented) {
    const accentedType = await across("Public Type RécordAccent", "RécordAccent");
    check("and so is one with an accented letter in it",
      accentedType.head === TYPE && accentedType.oneColour,
      `head ${nameOf(accentedType.head)}, tail ${nameOf(accentedType.tail)}. A name painted in two `
      + "pieces means a tokenizer rule stopped reading at the accent.");

    const accentedCall = await across(`    Debug.Print ${name}.CalculérName(1)`, "CalculérName");
    check("an accented procedure name is a call, throughout",
      accentedCall.head === CALL && accentedCall.oneColour,
      `head ${nameOf(accentedCall.head)}, tail ${nameOf(accentedCall.tail)}`);
  }
} finally {
  if (made) {
    await api.pane("close", { module: name, project: project.projectId, answer: "discard" }).catch(() => {});
    await api.component("remove", { name, project: project.projectId }).catch(() => {});
  }

  console.log(`\n${passed} passed, ${failures.length} failed`);
  for (const one of failures) { console.log(`  ${one}`); }

  process.exitCode = failures.length === 0 ? 0 : 1;
}
