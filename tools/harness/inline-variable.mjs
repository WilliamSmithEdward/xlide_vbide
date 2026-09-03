/*
 * Inline Variable on the LIVE product: the menu path, and the module the HOST holds afterwards.
 *
 * The pair to extract-variable.mjs, and it checks the same round trip from the other direction:
 * extract then inline should leave the module where it started.
 *
 * Run against the debug fixture:
 *   node tools\harness\inline-variable.mjs
 */

import { open, reporter, wait, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const { check, done } = reporter();
const CRLF = "\r\n";
const NAME = `Inline${process.pid % 10000}`;

const BEFORE = [
  "Option Explicit",              // 1
  "",                            // 2
  "Public Sub Post()",           // 3
  "    Dim limit As Long",       // 4
  "    limit = 10",              // 5
  "    Debug.Print limit",       // 6
  "    Debug.Print limit * 2",   // 7
  "End Sub",                     // 8
];

async function hostText() {
  const answer = await api.readModule(NAME, project.projectId);
  return String(answer.source ?? answer.text ?? "").split(/\r\n|\r|\n/);
}

const has = (lines, text) => lines.some((one) => one.trim() === text);

try {
  await api.component("remove", { name: NAME, project: project.projectId }).catch(() => { });
  await api.component("add", { kind: "module", name: NAME, project: project.projectId });
  await api.writeModule(NAME, BEFORE.join(CRLF), project.projectId);

  await api.caret(4, { module: NAME, project: project.projectId });
  await waitFor("the page to be showing the module", async () =>
    ((await api.ui()).focus.model ?? "").toLowerCase().endsWith(`/${NAME.toLowerCase()}`));

  /* ---- the lightbulb offers it on a name -------------------------------------------------------- */

  const titles = (answer) => (answer.data ?? []).map((one) => one.title);

  // THE CARET, not the range the act asks over: the entry reads where the developer is, which is
  // what a lightbulb does, and `caret` is what puts them there.
  await api.caret(4, { module: NAME, project: project.projectId, column: 9 });
  await wait(300);
  const onName = await api.act("quickFixes", { line: 4, column: 9 });
  check("a caret on the name is offered Inline",
    titles(onName).includes("Inline 'limit'"), titles(onName).join(" | ") || "(none)");

  /* ---- the menu path ---------------------------------------------------------------------------- */

  const made = await api.act("inlineVariable", { word: "limit" });
  check("the menu path inlines it", made.did, made.detail);

  const after = await hostText();
  check("every use took the value",
    has(after, "Debug.Print 10") && has(after, "Debug.Print 10 * 2"),
    after.filter((one) => one.includes("Debug.Print")).map((one) => one.trim()).join(" | "));
  check("the declaration is gone from the host's own module",
    !after.some((one) => one.includes("Dim limit")), after.map((one) => one.trim()).join(" | "));
  check("and so is the assignment",
    !after.some((one) => one.includes("limit =")), after.map((one) => one.trim()).join(" | "));
  check("the answer counted the uses", made.data?.replaced === 2, JSON.stringify(made.data));

  const parity = await waitFor("the surface to agree with the host",
    async () => { const now = await api.inSync(); return now.contentAgrees ? now : null; },
    { budgetMs: 15000 }).catch(() => null);
  check("the surface shows what the host holds", parity !== null,
    parity ? `${parity.nativeLines} line(s)` : "still disagreeing");

  const clean = await waitFor("the findings to settle", async () => {
    const rows = ((await api.problems()).findings ?? [])
      .filter((one) => (one.module ?? "").toLowerCase() === NAME.toLowerCase());
    return rows.length === 0 ? [] : null;
  }, { budgetMs: 20000 }).catch(() => null);
  check("the module still analyses clean", clean !== null,
    clean === null
      ? ((await api.problems()).findings ?? [])
        .filter((one) => (one.module ?? "").toLowerCase() === NAME.toLowerCase())
        .map((one) => one.code).join(",")
      : "(none)");

  /* ---- and back again, which is the pair's whole point ------------------------------------------- */

  // Extract the 10 back out, then inline it again: two refactorings that undo each other should
  // leave the module where it started, and nothing else checks that they agree about anything.
  const line = after.findIndex((one) => one.trim() === "Debug.Print 10") + 1;
  const column = (after[line - 1] ?? "").indexOf("10") + 1;
  const extracted = await api.act("extractVariable", {
    startLine: line, startColumn: column, endLine: line, endColumn: column + 2, name: "limit",
  });
  check("the value extracts back into a variable", extracted.did, extracted.detail);

  const round = await hostText();
  // AS DOUBLE, not As Long, and that is not a round-trip failure: `10` as an EXPRESSION is a
  // Double to the shared inference, which widens every numeric literal. The neighbouring
  // declare-variable quick fix special-cases a whole-number literal to Long and this path does
  // not, which is an inconsistency in the analyzer rather than in either consumer - filed as
  // xlide_vscode#64. Asserted as it behaves, because a suite that asserts what it wishes were
  // true is a suite that goes red on the day the wish is granted.
  check("and the module reads as it did before, bar the widened literal",
    has(round, "Dim limit As Double") && has(round, "limit = 10") && has(round, "Debug.Print limit"),
    round.map((one) => one.trim()).filter(Boolean).join(" | "));

  /* ---- Ctrl+Z ------------------------------------------------------------------------------------ */

  await api.act("undo");
  const back = await waitFor("the extraction to come back", async () => {
    const now = await hostText();
    return now.some((one) => one.includes("Dim limit")) ? null : now;
  }, { budgetMs: 15000 }).catch(() => null);
  check("undo puts the last one back", back !== null,
    back === null ? "the declaration is still there" : `${back.length} line(s)`);

  /* ---- the refusals ------------------------------------------------------------------------------- */

  await api.writeModule(NAME, [
    "Option Explicit", "", "Public Sub Post()",
    "    Dim total As Long", "    total = 2 + 3", "    Debug.Print total", "End Sub",
  ].join(CRLF), project.projectId);
  await wait(600);

  const brackets = await api.act("inlineVariable", { word: "total" });
  check("an expression that would need brackets is refused, and says why",
    !brackets.did && /brackets/i.test(String(brackets.detail))
    && /by value instead of by reference/i.test(String(brackets.detail)),
    brackets.detail);

  const untouched = await hostText();
  check("and a refusal wrote nothing",
    untouched.some((one) => one.includes("Dim total")), untouched.map((one) => one.trim()).join(" | "));
} finally {
  await api.pane("close", { module: NAME, project: project.projectId, answer: "discard" }).catch(() => { });
  await api.component("remove", { name: NAME, project: project.projectId }).catch(() => { });
}

done();
