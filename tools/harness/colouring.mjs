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
import { open, waitFor } from "./xlide-api.mjs";

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
  "    Print #1, \"the statement, which keeps its keyword\"",
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

  // WAITED FOR, NOT SLEPT THROUGH. These were `wait(2000)` and `wait(3000)`, which is 5s of a 5.3s
  // suite: the checking took 289ms and the rest was two guesses at how long a machine takes. Both
  // have a condition worth naming, and naming it is also the only way the suite stops being a race
  // that has not lost yet (driving-excel.md).
  await waitFor("the module to hold what was written", async () =>
    ((await api.readModule(name, project.projectId)).text ?? "").includes("End Function"));

  await api.pane("open", { module: name, project: project.projectId });

  /*
   * The tab opening is not the thing; the SEMANTIC TOKENIZER having run over it is, and it is
   * rebuilt per project from the words the engine sends. Until it has, every span reads as a
   * plain identifier - which is a colour, so a check racing this one fails saying `call` was
   * painted `identifier`, and that is exactly what a real defect here looks like.
   *
   * WAITED ON THE DECLARATION, NOT ON ANYTHING ASSERTED BELOW. The first version of this waited
   * for the bare call to be painted as a call, which is precisely what the first check then
   * asserts - so that check could no longer fail, only turn into a timeout. A readiness wait has
   * to name a DIFFERENT observable from the one under test or it launders the assertion into the
   * setup. No check reads the declaration line, and the whole point of this suite is that a
   * declaration, a bare call and a qualified call go through different rules, so a defect in any
   * of those still reports as a failed check rather than being waited away.
   */
  // Advisory, and deliberately so. A wait that THROWS when the thing never arrives turns nine
  // readable colour failures into one timeout, which is a worse report of the same breakage: the
  // sleep it replaced at least let the checks speak. So this gives up quietly and says it did.
  await waitFor("the tokenizer to paint the module", async () =>
    (await across("Public Sub Recalculate(ByVal label As String)", "Recalculate")).head === CALL,
    { budgetMs: 8000 },
  ).catch(() => console.log("     (the semantic pass never arrived; the colours below say what did)"));

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
   * 3. `Debug.Print` IS AN OBJECT AND A METHOD, and used to be two keywords.
   *
   * `Debug` rode in the canonical list, which is a SPELLING list: the formatter is there to turn
   * `debug.print` into `Debug.Print`. Colouring read it as a claim about the word's kind and
   * painted the intrinsic debug object the colour of an If.
   *
   * `Print` was worse, because it was defended. After a dot a keyword is not a keyword: there is
   * no VBA construct where `.Print` is the Print statement. The member rule asked `@keywords`
   * anyway, so `.Close`, `.Open` and `.Type` painted as syntax too.
   */
  const intrinsic = await across("    Debug.Print label", "Debug");
  check("the intrinsic Debug object is painted as an object",
    intrinsic.head === TYPE && intrinsic.oneColour,
    `it is ${nameOf(intrinsic.head)}. It rides in the canonical list because the formatter respells `
    + "it, and that list is about spelling, not about what a word is.");

  const method = await across("    Debug.Print label", "Print");
  check("and Print after the dot is a call, not a statement keyword",
    method.head === CALL,
    `it is ${nameOf(method.head)}. After a dot the word is a member, always: no VBA construct `
    + "puts the Print statement there.");

  // AND THE SAME WORD WITHOUT A DOT IS STILL THE STATEMENT. This is what stops the fix above
  // turning a keyword list into a suggestion.
  const statement = await across("    Print #1, \"the statement, which keeps its keyword\"", "Print");
  check("the bare Print statement is still a keyword",
    statement.head === KEYWORD,
    `it is ${nameOf(statement.head)}. Print #1, is a real VBA statement and nothing about the `
    + "member rule should reach it.");

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
