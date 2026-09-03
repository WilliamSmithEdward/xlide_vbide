/*
 * Encapsulate Field on the LIVE product: the lightbulb, and the class the HOST holds afterwards.
 *
 * The engine's suite proves the rewrite. This proves the half it cannot see, and one thing neither
 * can prove alone: that the code which USED the field still compiles, because the property took
 * the field's name. A rewrite that renamed the variable would pass every text check in the engine
 * and break every caller in the project.
 *
 * Run against the debug fixture:
 *   node tools\harness\encapsulate-field.mjs
 */

import { open, reporter, wait, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const { check, done } = reporter();
const CRLF = "\r\n";
const NAME = `Holder${process.pid % 10000}`;
const USER = `Uses${process.pid % 10000}`;

const CLASS = [
  "Option Explicit",           // 1
  "",                          // 2
  "Public Label As String",    // 3
  "Public Log As Object",      // 4
  "",                          // 5
  "Public Sub Announce()",     // 6
  "    Debug.Print Label",     // 7
  "End Sub",                   // 8
];

// A SECOND MODULE THAT USES IT, which is the point: the property keeps the name, so this module
// must go on compiling untouched. Nothing here is rewritten by the refactoring.
const CALLER = [
  "Option Explicit",
  "",
  "Public Sub Fill()",
  `    Dim held As ${NAME}`,
  `    Set held = New ${NAME}`,
  '    held.Label = "posted"',
  "    Debug.Print held.Label",
  "End Sub",
];

async function hostText(name) {
  const answer = await api.readModule(name, project.projectId);
  return String(answer.source ?? answer.text ?? "").split(/\r\n|\r|\n/);
}

const has = (lines, text) => lines.some((one) => one.trim() === text);

try {
  for (const name of [NAME, USER]) {
    await api.component("remove", { name, project: project.projectId }).catch(() => { });
  }

  await api.component("add", { kind: "class", name: NAME, project: project.projectId });
  await api.component("add", { kind: "module", name: USER, project: project.projectId });
  await api.writeModule(NAME, CLASS.join(CRLF), project.projectId);
  await api.writeModule(USER, CALLER.join(CRLF), project.projectId);

  await api.caret(3, { module: NAME, project: project.projectId });
  await waitFor("the page to be showing the class", async () =>
    ((await api.ui()).focus.model ?? "").toLowerCase().endsWith(`/${NAME.toLowerCase()}`));

  /* ---- the lightbulb offers it on a declaration, and only there ------------------------------- */

  const titles = (answer) => (answer.data ?? []).map((one) => one.title);

  const onField = await api.act("quickFixes", { line: 3, column: 8 });
  check("a declaration line is offered encapsulation",
    titles(onField).includes("Encapsulate 'Label' in a property"),
    titles(onField).join(" | ") || "(none)");

  await api.caret(7, { module: NAME, project: project.projectId });
  await wait(300);
  const inBody = await api.act("quickFixes", { line: 7, column: 5 });
  check("and a line inside a procedure is not",
    !titles(inBody).some((one) => one.startsWith("Encapsulate ")),
    titles(inBody).join(" | ") || "(none)");

  /* ---- the rewrite ---------------------------------------------------------------------------- */

  const made = await api.act("encapsulateField", { fieldName: "Label" });
  check("the field is encapsulated", made.did, made.detail);

  const after = await hostText(NAME);
  check("the host's own class holds the private variable",
    has(after, "Private m_Label As String") && !has(after, "Public Label As String"),
    after.find((one) => one.includes("m_Label")) ?? "(absent)");
  check("and the property pair over it",
    has(after, "Public Property Get Label() As String")
    && has(after, "Public Property Let Label(ByVal RHS As String)"),
    (made.data?.accessors ?? []).join(" / "));
  check("the declaration is still above the procedures, where VBA insists",
    after.indexOf("Private m_Label As String") < after.findIndex((one) => one.startsWith("Public Sub")),
    `${after.indexOf("Private m_Label As String")} against ${after.findIndex((one) => one.startsWith("Public Sub"))}`);
  check("the class's own use of the name was not touched",
    has(after, "Debug.Print Label"), after.find((one) => one.includes("Debug.Print")) ?? "(absent)");

  // THE OTHER MODULE, byte for byte. The property keeping the name is the whole reason this
  // refactoring needs no call-site rewriting, and the way to prove it is to show that a module
  // full of call sites was not rewritten and still analyses.
  const caller = await hostText(USER);
  check("the module that uses it was not rewritten at all",
    caller.join(CRLF).trim() === CALLER.join(CRLF).trim(),
    caller.filter((one) => one.includes("held.Label")).map((one) => one.trim()).join(" | "));

  const parity = await waitFor("the surface to agree with the host",
    async () => { const now = await api.inSync(); return now.contentAgrees ? now : null; },
    { budgetMs: 15000 }).catch(() => null);
  check("the surface shows what the host holds", parity !== null,
    parity ? `${parity.nativeLines} line(s)` : "still disagreeing");

  const clean = await waitFor("the findings to settle", async () => {
    const rows = ((await api.problems()).findings ?? [])
      .filter((one) => [NAME.toLowerCase(), USER.toLowerCase()].includes((one.module ?? "").toLowerCase()));
    return rows.length === 0 ? [] : null;
  }, { budgetMs: 20000 }).catch(() => null);
  check("both modules analyse clean afterwards", clean !== null,
    clean === null
      ? ((await api.problems()).findings ?? [])
        .filter((one) => [NAME.toLowerCase(), USER.toLowerCase()].includes((one.module ?? "").toLowerCase()))
        .map((one) => `${one.module}:${one.code}`).join(",")
      : "(none)");

  /* ---- an object field, where VBA wants Set ---------------------------------------------------- */

  const object = await api.act("encapsulateField", { fieldName: "Log" });
  check("an Object field gets a Property Set", object.did, object.detail);

  const both = await hostText(NAME);
  check("and it assigns with Set on both sides",
    has(both, "Public Property Set Log(ByVal RHS As Object)")
    && has(both, "Set Log = m_Log") && has(both, "Set m_Log = RHS"),
    both.filter((one) => one.trim().startsWith("Set ")).map((one) => one.trim()).join(" | "));

  /* ---- Ctrl+Z ---------------------------------------------------------------------------------- */

  await api.act("undo");
  const back = await waitFor("the Object field to come back", async () => {
    const now = await hostText(NAME);
    return now.some((one) => one.includes("m_Log")) ? null : now;
  }, { budgetMs: 15000 }).catch(() => null);
  check("undo puts the last encapsulation back", back !== null,
    back === null ? "m_Log is still there" : `${back.length} line(s)`);
  check("and leaves the one before it alone",
    back !== null && has(back, "Private m_Label As String"),
    back ? (back.find((one) => one.includes("m_Label")) ?? "(absent)") : "(nothing)");

  /* ---- the refusals ----------------------------------------------------------------------------- */

  const already = await api.act("encapsulateField", { fieldName: "m_Label" });
  check("a variable that is already private is refused",
    !already.did && /already Private/i.test(String(already.detail)), already.detail);

  const missing = await api.act("encapsulateField", { fieldName: "NoSuchField" });
  check("a name the module does not declare is refused",
    !missing.did && /declares no module-level variable/i.test(String(missing.detail)), missing.detail);
} finally {
  for (const name of [USER, NAME]) {
    await api.pane("close", { module: name, project: project.projectId, answer: "discard" }).catch(() => { });
    await api.component("remove", { name, project: project.projectId }).catch(() => { });
  }
}

done();
