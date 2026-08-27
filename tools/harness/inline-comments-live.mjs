/*
 * The inline suppression directives and ''' doc comments, on the LIVE surface.
 *
 * The engine half is held headlessly by engine/test/inline-comments.mjs; this is the other half,
 * each stage of which could drop what the engine sends: the Problems pane fed by the analysis
 * pass, the page's hover rendering the `documentation` field, and the lightbulb offering the
 * suppression fix. The owner's rule for this library is that everything stays INLINE - there is
 * no sidecar metadata loader to test because none ships.
 *
 * Run against the debug fixture:
 *   node tools\harness\inline-comments-live.mjs
 */

import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const { check, done } = reporter();
const CRLF = "\r\n";
const NAME = `Inline${process.pid % 10000}`;

/*
 * WAITS NAME THE STATE THEY WAIT FOR, and a timeout is a failure, never a settle.
 *
 * The first probe of this waited for the rows to stop changing, and empty-then-empty read as
 * settled: it reported "(none)" for a finding that was plainly there - the lightbulb offered a
 * fix for it one check later - and passed the suppression check vacuously against the same
 * emptiness. (The probe also read `.problems` off a reply whose array is named `findings`,
 * which is the same instrument fault in one more coat.)
 */
async function problemsFor(module, expected) {
  let rows = [];
  const until = Date.now() + 25000;
  while (Date.now() < until) {
    await wait(600);
    const now = (await api.problems().catch(() => ({ findings: [] }))).findings ?? [];
    rows = now.filter((one) => (one.module ?? "").toLowerCase() === module.toLowerCase());
    if (expected(rows)) { return { rows, arrived: true }; }
  }
  return { rows, arrived: false };
}

const codes = (rows) => rows.map((one) => one.code).sort().join(",") || "(none)";

try {
  await api.command("reset").catch(() => {});
  await api.component("remove", { name: NAME, project: project.projectId }).catch(() => {});
  await api.component("add", { kind: "module", name: NAME, project: project.projectId });

  /* ---- the Problems pane, before and after the directive -------------------------------------- */

  await api.writeModule(NAME, [
    "Option Explicit", "",
    "Public Sub Go()", "    missing1 = 1", "End Sub",
  ].join(CRLF), project.projectId);
  const before = await problemsFor(NAME,
    (rows) => rows.some((one) => one.code === "undeclared-variable"));
  check("the finding reaches the Problems pane", before.arrived, codes(before.rows));

  await api.writeModule(NAME, [
    "Option Explicit", "",
    "Public Sub Go()",
    "    ' @xlide-analysis-disable-next-line undeclared-variable -- held by this suite",
    "    missing1 = 1",
    "End Sub",
  ].join(CRLF), project.projectId);
  const after = await problemsFor(NAME, (rows) => rows.length === 0);
  check("a suppression directive empties it", after.arrived, codes(after.rows));

  await api.writeModule(NAME, [
    "Option Explicit", "",
    "Public Sub Go()",
    "    ' @xlide-analysis-disable-next-line bogus-code",
    "    missing1 = 1",
    "End Sub",
  ].join(CRLF), project.projectId);
  const malformed = await problemsFor(NAME, (rows) =>
    rows.some((one) => one.code === "undeclared-variable")
      && rows.some((one) => one.code === "analysis-suppression-directive"));
  check("an unknown code keeps the finding and reports the directive",
    malformed.arrived, codes(malformed.rows));

  /* ---- the lightbulb offers the suppression --------------------------------------------------- */

  await api.writeModule(NAME, [
    "Option Explicit", "",
    "Public Sub Go()", "    missing1 = 1", "End Sub",
  ].join(CRLF), project.projectId);
  await problemsFor(NAME, (rows) => rows.some((one) => one.code === "undeclared-variable"));
  await api.pane("open", { module: NAME, project: project.projectId });
  await wait(800);

  const fixes = await api.act("quickFixes", { line: 4, column: 6 });
  const titles = (fixes.data ?? []).map((one) => one.title);
  check("the lightbulb offers the suppression fix",
    titles.some((title) => /Suppress 'undeclared-variable'/.test(title)),
    titles.join(" | ").slice(0, 120));

  /* ---- ''' docs in the page's own hover ------------------------------------------------------- */

  await api.writeModule(NAME, [
    "Option Explicit", "",
    "''' <summary>Doubles a count for the probe.</summary>",
    "''' <param name=\"n\">how many to start from</param>",
    "Public Function Twice(ByVal n As Long) As Long",
    "    Twice = n * 2",
    "End Function", "",
    "Public Sub Caller()",
    "    Dim answer As Long",
    "    answer = Twice(21)",
    "End Sub",
  ].join(CRLF), project.projectId);
  await problemsFor(NAME, (rows) => rows.length === 0);
  await api.pane("open", { module: NAME, project: project.projectId });
  await wait(800);

  // act("hover") answers Monaco's own Hover - { contents: [{ value }], range } - through the
  // page's provider, which is the rendering a person gets. Line 11 is `answer = Twice(21)`.
  const hover = await api.act("hover", { line: 11, column: 15 });
  const blocks = (hover?.contents ?? hover?.data?.contents ?? [])
    .map((one) => one.value ?? "").join(" ");
  check("hover on the call renders the ''' summary",
    blocks.includes("Doubles a count"), blocks.replace(/\s+/g, " ").slice(0, 120));
  check("and the <param> doc rides along",
    blocks.includes("how many to start from"), blocks.replace(/\s+/g, " ").slice(0, 120));
} finally {
  await api.command("reset").catch(() => {});
  await api.component("remove", { name: NAME, project: project.projectId }).catch(() => {});
}

process.exit(done());
