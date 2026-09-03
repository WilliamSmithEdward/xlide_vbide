/*
 * Extract Variable on the LIVE product: the menu path, and the module the HOST holds afterwards.
 *
 * The one thing this can ask that the engine's suite cannot is whether the declared type and the
 * `Set` survive the round trip into a real project - the type comes from the analyzer's view of
 * THIS workbook's symbols, not of a synthetic two-module project.
 *
 * Run against the debug fixture:
 *   node tools\harness\extract-variable.mjs
 */

import { open, reporter, wait, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const { check, done } = reporter();
const CRLF = "\r\n";
const NAME = `Value${process.pid % 10000}`;

const BEFORE = [
  "Option Explicit",                          // 1
  "",                                         // 2
  "Public Sub Post()",                        // 3
  "    Dim total As Long",                    // 4
  "    total = 2",                            // 5
  "    Debug.Print total * 3",                // 6
  "End Sub",                                  // 7
  "",                                         // 8
  "Public Sub Fill()",                        // 9
  "    Debug.Print New Collection.Count",     // 10
  "End Sub",                                  // 11
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

  await api.caret(6, { module: NAME, project: project.projectId });
  await waitFor("the page to be showing the module", async () =>
    ((await api.ui()).focus.model ?? "").toLowerCase().endsWith(`/${NAME.toLowerCase()}`));

  /* ---- the lightbulb offers it on a selection, and only then ---------------------------------- */

  const titles = (answer) => (answer.data ?? []).map((one) => one.title);

  const bare = await api.act("quickFixes", { line: 6, column: 17 });
  check("a bare caret is not offered Extract variable",
    !titles(bare).includes("Extract variable..."), titles(bare).join(" | ") || "(none)");

  // `total * 3` on line 6, which starts at column 17.
  await api.act("select", { startLine: 6, endLine: 6 });
  await wait(300);
  const selected = await api.act("quickFixes", { line: 6, column: 17 });
  check("and a selection is offered it",
    titles(selected).includes("Extract variable..."), titles(selected).join(" | ") || "(none)");

  /* ---- a value expression --------------------------------------------------------------------- */

  const made = await api.act("extractVariable", {
    startLine: 6, startColumn: 17, endLine: 6, endColumn: 26, name: "scaled",
  });
  check("the menu path extracts an expression", made.did, made.detail);

  const after = await hostText();
  check("the host's own module carries the declaration",
    has(after, "Dim scaled As Double"), after.find((one) => one.includes("Dim scaled")) ?? "(absent)");
  check("assigned without Set, because it is not an object",
    has(after, "scaled = total * 3"), after.find((one) => one.includes("scaled =")) ?? "(absent)");
  check("and the selection is replaced by the name",
    has(after, "Debug.Print scaled"), after.find((one) => one.includes("Debug.Print scaled")) ?? "(absent)");
  check("the declaration keeps the statement's indentation",
    after.includes("    Dim scaled As Double"),
    JSON.stringify(after.find((one) => one.includes("Dim scaled")) ?? ""));

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

  /* ---- an object expression, where the answer is Set ------------------------------------------- */

  // `New Collection` moved down a line by the declaration above it.
  const object = await waitFor("the New Collection line", async () => {
    const now = await hostText();
    const at = now.findIndex((one) => one.includes("New Collection"));
    return at >= 0 ? { at: at + 1, text: now[at] } : null;
  });

  const column = object.text.indexOf("New Collection") + 1;
  const made2 = await api.act("extractVariable", {
    startLine: object.at, startColumn: column, endLine: object.at,
    endColumn: column + "New Collection".length, name: "items",
  });
  check("an object expression extracts", made2.did, made2.detail);

  const both = await hostText();
  check("it is declared as its own class, not as Object",
    has(both, "Dim items As Collection"), both.find((one) => one.includes("Dim items")) ?? "(absent)");
  check("and assigned WITH Set, which is what the analyzer answered",
    has(both, "Set items = New Collection"), both.find((one) => one.includes("items =")) ?? "(absent)");

  /* ---- Ctrl+Z ---------------------------------------------------------------------------------- */

  await api.act("undo");
  const back = await waitFor("the object extraction to come back", async () => {
    const now = await hostText();
    return now.some((one) => one.includes("m_items") || one.includes("Dim items")) ? null : now;
  }, { budgetMs: 15000 }).catch(() => null);
  check("undo puts the last extraction back", back !== null,
    back === null ? "the declaration is still there" : `${back.length} line(s)`);
  check("and leaves the one before it alone",
    back !== null && has(back, "Dim scaled As Double"),
    back ? (back.find((one) => one.includes("scaled")) ?? "(absent)") : "(nothing)");

  /* ---- the refusals ----------------------------------------------------------------------------- */

  // LOCATED BY CONTENT, not by the line it started on: two extractions have added lines above it,
  // and a check that asserts on a refusal has to be sure it provoked the one it names.
  //
  // `total *` and not, say, `New Colle` - which the analyzer accepts, correctly, as a whole `New`
  // expression naming a class that happens not to exist. A binary operator with nothing on its
  // right is the shape that genuinely cannot be an expression.
  const whole = await hostText();
  const line = whole.findIndex((one) => one.includes("total * 3")) + 1;
  const from = (whole[line - 1] ?? "").indexOf("total * 3") + 1;
  const partial = await api.act("extractVariable", {
    startLine: line, startColumn: from, endLine: line, endColumn: from + "total *".length, name: "half",
  });
  check("half an expression is refused, in the dialog's own words",
    !partial.did && /whole expression/i.test(String(partial.detail)), partial.detail);

  const standing = (await api.ui()).extract;
  check("and the dialog was taken down with it", standing === null,
    standing ? JSON.stringify(standing) : "no dialog standing");
} finally {
  await api.act("extractDialog", { press: "cancel" }).catch(() => { });
  await api.pane("close", { module: NAME, project: project.projectId, answer: "discard" }).catch(() => { });
  await api.component("remove", { name: NAME, project: project.projectId }).catch(() => { });
}

done();
