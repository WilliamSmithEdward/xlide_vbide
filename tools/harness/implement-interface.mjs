/*
 * Implement Interface on the LIVE product: the menu path, and the class the HOST holds afterwards.
 *
 * The engine's own suite proves the stubs. This proves the half it cannot see - that a developer
 * with the caret on an `Implements` line ends up with a class the VBE would compile, in the
 * workbook and not only on the surface.
 *
 * Run against the debug fixture:
 *   node tools\harness\implement-interface.mjs
 */

import { open, reporter, wait, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const { check, done } = reporter();
const CRLF = "\r\n";
const IFACE = `IStore${process.pid % 10000}`;
const IMPL = `Store${process.pid % 10000}`;

const INTERFACE = [
  "Option Explicit",
  "",
  "Public Count As Long",
  "",
  "Public Sub Save()",
  "End Sub",
  "",
  "Public Function Total(ByVal rate As Double, Optional ByVal since As Date) As Currency",
  "End Function",
  "",
  "Public Property Get Name() As String",
  "End Property",
];

const CLASS = [
  "Option Explicit",
  "",
  `Implements ${IFACE}`,
];

/** What the HOST holds for the class, which is the text the VBE compiles. */
async function hostText(name) {
  const answer = await api.readModule(name, project.projectId);
  return String(answer.source ?? answer.text ?? "").split(/\r\n|\r|\n/);
}

const has = (lines, text) => lines.some((one) => one.trim() === text);

try {
  for (const name of [IFACE, IMPL]) {
    await api.component("remove", { name, project: project.projectId }).catch(() => { });
  }

  await api.component("add", { kind: "class", name: IFACE, project: project.projectId });
  await api.component("add", { kind: "class", name: IMPL, project: project.projectId });
  await api.writeModule(IFACE, INTERFACE.join(CRLF), project.projectId);
  await api.writeModule(IMPL, CLASS.join(CRLF), project.projectId);

  // The caret on the Implements line, which is where the entry lives and where a developer
  // would be standing.
  await api.caret(3, { module: IMPL, project: project.projectId });
  await waitFor("the page to be showing the class", async () =>
    ((await api.ui()).focus.model ?? "").toLowerCase().endsWith(`/${IMPL.toLowerCase()}`));

  /* ---- the lightbulb offers it on the Implements line, and only there ------------------------ */

  const titles = (answer) => (answer.data ?? []).map((one) => one.title);

  const onLine = await api.act("quickFixes", { line: 3, column: 5 });
  check("the Implements line is offered its members",
    titles(onLine).includes(`Implement members of ${IFACE}`), titles(onLine).join(" | ") || "(none)");

  await api.caret(1, { module: IMPL, project: project.projectId });
  await wait(300);
  const elsewhere = await api.act("quickFixes", { line: 1, column: 5 });
  check("and a line that is not one is not",
    !titles(elsewhere).some((one) => one.startsWith("Implement members of")),
    titles(elsewhere).join(" | ") || "(none)");

  /* ---- the menu path ------------------------------------------------------------------------- */

  const written = await api.act("implementInterface", { interfaceName: IFACE });
  check("the members are written", written.did, written.detail);
  check("and the answer names each one",
    (written.data?.added ?? []).join(",") === `${IFACE}_Count,${IFACE}_Count,${IFACE}_Save,${IFACE}_Total,${IFACE}_Name`,
    (written.data?.added ?? []).join(","));

  const after = await hostText(IMPL);
  check("the host's own class carries the Sub",
    has(after, `Private Sub ${IFACE}_Save()`), after.find((one) => one.includes("_Save")) ?? "(absent)");
  check("the Function keeps the interface's parameters word for word",
    has(after, `Private Function ${IFACE}_Total(ByVal rate As Double, Optional ByVal since As Date) As Currency`),
    after.find((one) => one.includes("_Total")) ?? "(absent)");
  check("the public field became a Get and a Let",
    has(after, `Private Property Get ${IFACE}_Count() As Long`)
    && has(after, `Private Property Let ${IFACE}_Count(ByVal RHS As Long)`),
    after.filter((one) => one.includes("_Count")).map((one) => one.trim()).join(" | "));
  check("every stub raises rather than silently doing nothing",
    after.filter((one) => one.trim().startsWith("Err.Raise 5")).length === 5,
    `${after.filter((one) => one.trim().startsWith("Err.Raise 5")).length} raise(s)`);
  check("and the Implements line is still where it was",
    after[2]?.trim() === `Implements ${IFACE}`, after[2]);

  // THE SURFACE AGREES WITH THE HOST, and the analyzer is content with what was written: a class
  // that no longer compiles would read as a success from the text alone.
  const agreed = await waitFor("the surface to agree with the host",
    async () => { const now = await api.inSync(); return now.contentAgrees ? now : null; },
    { budgetMs: 15000 }).catch(() => null);
  check("the surface shows what the host holds", agreed !== null,
    agreed ? `${agreed.nativeLines} line(s)` : "still disagreeing");

  const clean = await waitFor("the findings to settle", async () => {
    const rows = ((await api.problems()).findings ?? [])
      .filter((one) => (one.module ?? "").toLowerCase() === IMPL.toLowerCase());
    return rows.length === 0 ? [] : null;
  }, { budgetMs: 20000 }).catch(() => null);
  check("the class the stubs went into analyses clean", clean !== null,
    clean === null
      ? ((await api.problems()).findings ?? [])
        .filter((one) => (one.module ?? "").toLowerCase() === IMPL.toLowerCase())
        .map((one) => one.code).join(",")
      : "(none)");

  /* ---- asking twice writes nothing twice ------------------------------------------------------ */

  const again = await api.act("implementInterface", { interfaceName: IFACE });
  check("asking again is refused rather than doubling the members",
    !again.did && /already implements every member/i.test(String(again.detail)), again.detail);

  // HEADERS, not mentions: each stub names itself again in its own Err.Raise, so counting
  // mentions counts every member twice and reads as a doubling that did not happen.
  const unchanged = await hostText(IMPL);
  const headers = unchanged.filter((one) => /^Private (Sub|Function|Property)/.test(one));
  check("and the class is exactly as it was",
    headers.length === 5 && headers.filter((one) => one.includes(`${IFACE}_Save`)).length === 1,
    `${headers.length} stub(s), ${headers.filter((one) => one.includes(`${IFACE}_Save`)).length} of them ${IFACE}_Save`);

  /* ---- Ctrl+Z ---------------------------------------------------------------------------------- */

  await api.act("undo");
  const back = await waitFor("the class to come back", async () => {
    const now = await hostText(IMPL);
    return now.some((one) => one.includes(`${IFACE}_Save`)) ? null : now;
  }, { budgetMs: 15000 }).catch(() => null);

  check("undo puts the stubs back", back !== null,
    back === null ? "the stubs are still there" : `${back.length} line(s)`);
  check("and the Implements line survived the undo",
    back !== null && has(back, `Implements ${IFACE}`), back ? back.join(" / ").slice(0, 80) : "(nothing)");

  /* ---- the refusals ---------------------------------------------------------------------------- */

  const wrong = await api.act("implementInterface", { interfaceName: "INotDeclared" });
  check("an interface the class does not declare is refused, naming the ones it does",
    !wrong.did && /does not implement 'INotDeclared'/i.test(String(wrong.detail))
    && String(wrong.detail).includes(IFACE),
    wrong.detail);

  await api.caret(1, { module: IFACE, project: project.projectId });
  await waitFor("the page to be showing the interface", async () =>
    ((await api.ui()).focus.model ?? "").toLowerCase().endsWith(`/${IFACE.toLowerCase()}`));
  const none = await api.act("implementInterface", {});
  check("a class that implements nothing is refused, and told what to add",
    !none.did && /does not declare Implements/i.test(String(none.detail)), none.detail);
} finally {
  for (const name of [IMPL, IFACE]) {
    await api.pane("close", { module: name, project: project.projectId, answer: "discard" }).catch(() => { });
    await api.component("remove", { name, project: project.projectId }).catch(() => { });
  }
}

done();
