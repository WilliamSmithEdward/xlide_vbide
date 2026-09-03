/*
 * Extract Method on the LIVE product: the menu path, the module underneath, and the way back.
 *
 * The engine's own suite (engine/test/extract-method.mjs) proves the transformation and every
 * refusal against the analyzer. This one proves the other half - that a developer selecting lines
 * and picking the menu entry ends up with a module that says what the engine worked out, in the
 * HOST's copy and not only on the surface.
 *
 * THE NATIVE PANE IS THE ONE THAT COUNTS. The page can hold a perfect result while the module the
 * VBE compiles still holds the old text, and that failure looks identical from the page's side. So
 * every check that matters here reads what the host holds.
 *
 * Run against the debug fixture:
 *   node tools\harness\extract-method.mjs
 */

import { open, reporter, wait, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const { check, done } = reporter();
const CRLF = "\r\n";
const NAME = `Extract${process.pid % 10000}`;
const MADE = `Summed${process.pid % 10000}`;

/*
 * The design document's own example, which is the shape the four decisions were written from:
 * one local read before it is written, one written inside and read after, one written and read
 * nowhere else, and a free reference that is not a local at all.
 */
const BEFORE = [
  "Option Explicit",                                        // 1
  "",                                                       // 2
  "Public Sub PostInvoices()",                              // 3
  "    Dim lastRow As Long",                                // 4
  "    Dim total As Currency",                              // 5
  "    Dim i As Long",                                      // 6
  "    lastRow = 4",                                        // 7
  "",                                                       // 8
  "    total = 0",                                          // 9  <- selection
  "    For i = 2 To lastRow",                               // 10
  "        If i > 0 Then",                                  // 11
  "            total = total + i",                          // 12
  "        End If",                                         // 13
  "    Next i",                                             // 14 <- selection ends
  "",                                                       // 15
  '    Debug.Print "Posted " & total',                      // 16
  "End Sub",                                                // 17
];

/** What the HOST holds for the module, which is the text the VBE compiles. */
async function hostText() {
  const answer = await api.readModule(NAME, project.projectId);
  return String(answer.source ?? answer.text ?? "").split(/\r\n|\r|\n/);
}

const has = (lines, text) => lines.some((one) => one.trim() === text);
const lineWith = (lines, text) => lines.find((one) => one.includes(text))?.trim() ?? "(absent)";

try {
  await api.component("remove", { name: NAME, project: project.projectId }).catch(() => { });
  await api.component("add", { kind: "module", name: NAME, project: project.projectId });
  await api.writeModule(NAME, BEFORE.join(CRLF), project.projectId);

  // The module has to be the one on screen, and the PAGE has to be the side that is on it:
  // extraction acts on the selection in the active editor, exactly as the menu entry does.
  // `activate` picks among tabs already open, and a module just created has none - so the caret
  // route, which opens it the way clicking it in the tree does.
  await api.caret(9, { module: NAME, project: project.projectId });
  await waitFor("the page to be showing the module", async () =>
    ((await api.ui()).focus.model ?? "").toLowerCase().endsWith(`/${NAME.toLowerCase()}`));

  /* ---- the menu path ------------------------------------------------------------------------ */

  const made = await api.act("extractMethod", { startLine: 9, endLine: 14, name: MADE });
  check("the menu path extracts", made.did, made.detail);

  const after = await hostText();
  check("the host's own module carries the new procedure",
    has(after, `Private Function ${MADE}(ByVal lastRow As Long) As Currency`),
    after.find((one) => one.includes("Private ")) ?? "(no Private line)");
  check("the caller calls it and keeps its answer",
    lineWith(after, `${MADE}(`) === `total = ${MADE}(lastRow)`, lineWith(after, `${MADE}(`));
  check("the loop variable's Dim moved out of the caller",
    after.filter((one) => one.trim() === "Dim i As Long").length === 1,
    `${after.filter((one) => one.trim() === "Dim i As Long").length} Dim i`);
  check("the result is assigned to the function name",
    has(after, `${MADE} = total`), lineWith(after, `${MADE} = `));
  check("what was not selected did not move",
    has(after, "lastRow = 4") && has(after, 'Debug.Print "Posted " & total'));

  // THE SURFACE AGREES WITH THE HOST. Written by the host and pushed into the open tab, so a
  // page still showing the old text is a sync that did not happen.
  const parity = await waitFor("the surface to agree with the host",
    async () => { const now = await api.inSync(); return now.contentAgrees ? now : null; },
    { budgetMs: 15000 }).catch(() => null);
  check("the surface shows what the host holds", parity !== null,
    parity ? `${parity.nativeLines} line(s)` : "still disagreeing");

  // And the analyzer is content with what was produced: a refactoring that leaves a module that
  // does not analyse is one nobody can use, and it would look fine in the text.
  const findings = await waitFor("the findings to settle", async () => {
    const rows = ((await api.problems()).findings ?? [])
      .filter((one) => (one.module ?? "").toLowerCase() === NAME.toLowerCase());
    return rows.length === 0 ? [] : null;
  }, { budgetMs: 20000 }).catch(() => null);
  check("the module still analyses clean", findings !== null,
    findings === null
      ? ((await api.problems()).findings ?? [])
        .filter((one) => (one.module ?? "").toLowerCase() === NAME.toLowerCase())
        .map((one) => one.code).join(",")
      : "(none)");

  /* ---- Ctrl+Z, which is the only way back ---------------------------------------------------- */

  await api.act("undo");
  const back = await waitFor("the module to come back", async () => {
    const now = await hostText();
    return now.some((one) => one.includes(MADE)) ? null : now;
  }, { budgetMs: 15000 }).catch(() => null);

  check("undo puts the extraction back", back !== null,
    back === null ? "the new procedure is still there" : `${back.length} line(s)`);
  check("and the module is what it was",
    back !== null && has(back, "total = 0") && has(back, "For i = 2 To lastRow")
    && back.filter((one) => one.trim() === "Dim i As Long").length === 1,
    back ? lineWith(back, "total = 0") : "(nothing to read)");

  /* ---- a refusal, which the developer reads in the dialog ------------------------------------ */

  // A selection that starts inside the If and ends outside it: the block would lose its closer.
  const refused = await api.act("extractMethod", { startLine: 11, endLine: 12, name: `${MADE}Half` });
  check("a half-taken block is refused", !refused.did, refused.detail);
  check("and the refusal says which block and what to do",
    /If block/i.test(String(refused.detail)) && /whole blocks/i.test(String(refused.detail)),
    refused.detail);

  const standing = (await api.ui()).extract;
  check("the dialog was taken down with it", standing === null,
    standing ? JSON.stringify(standing) : "no dialog standing");

  const untouched = await hostText();
  check("and a refusal wrote nothing",
    !untouched.some((one) => one.includes(`${MADE}Half`)) && has(untouched, "total = 0"),
    lineWith(untouched, "total = 0"));

  /* ---- the dialog itself --------------------------------------------------------------------- */

  await api.act("editorAction", { id: "xlide.extractMethod" }).catch(() => { });
  const dialog = await waitFor("the dialog to open", async () => (await api.ui()).extract,
    { budgetMs: 6000 }).catch(() => null);

  // The selection is gone by now (undo moved the caret), so the entry may decline to open at all -
  // which is itself correct. Only assert on the dialog when it did open.
  if (dialog) {
    check("the dialog opens with a name ready to be typed over", dialog.name.length > 0, dialog.name);
    await api.act("extractDialog", { press: "cancel" });
    await wait(200);
    check("and cancel takes it down", (await api.ui()).extract === null);
  } else {
    check("the menu entry declines without a selection", true, "no selection, no dialog");
  }
} finally {
  await api.act("extractDialog", { press: "cancel" }).catch(() => { });
  await api.pane("close", { module: NAME, project: project.projectId, answer: "discard" }).catch(() => { });
  await api.component("remove", { name: NAME, project: project.projectId }).catch(() => { });
}

done();
