/*
 * The library a project references, and the fix that gives it one.
 *
 * `Dim wd As Word.Application` in a workbook with no reference to Word does not compile - the VBE
 * stops the whole project - and the analyzer reports `missing-library-reference`. The fix is one
 * tick in Tools > References, so the finding carries it as a quick fix.
 *
 * WHY THIS IS A LIVE SUITE AND NOT AN ENGINE ONE. The engine's own suite
 * (engine/test/project-references.mjs) proves the finding, the fix's shape and the analysis that
 * follows, against the analyzer directly. None of that touches a host. What only a live session
 * can answer is whether the reference actually goes ON - it is a COM call into the VBE's own
 * References collection, not a text edit - and whether the editor notices. The second half is the
 * part that was wrong: nothing in any module's text changes when a reference is added, so every
 * cheap gate between the click and the squiggle sees identical inputs and leaves the findings
 * exactly where they were (2026-09-21).
 *
 * It brings its own module and takes it away, and it takes the reference off again, so it can run
 * against any open workbook and leaves it as it found it.
 *
 *   node tools\harness\references.mjs
 */

import { open, wait, waitFor, reporter } from "./xlide-api.mjs";

const api = await open({});
const { check, done } = reporter();

// Unique per run, for the reason analysis-freshness.mjs records: fixed names let one run inherit
// the last one's answers under the same name.
const RUN = String(process.pid).slice(-5);
const MODULE = `XlideRefs${RUN}`;

// Word, because the suite's own host is Excel and a project's own host library is implicit - a
// reference to it would be there already and prove nothing. The line the finding lands on is 4.
const SOURCE = [
  "Option Explicit",
  "",
  "Public Sub Automate()",
  "    Dim wd As Word.Application",
  "    Set wd = New Word.Application",
  "    wd.Visible = True",
  "End Sub",
  "",
].join("\r\n");

const FINDING_LINE = 4;
const project = await api.project();
const projectId = project.projectId;

const referencesNow = async () => (await api.project(projectId)).references ?? [];
const hasWord = async () => (await referencesNow()).some((one) => one.name === "Word");
const missingNow = async () => {
  const answer = await api.problems(MODULE);
  return (answer.findings ?? answer.problems ?? [])
    .filter((one) => one.code === "missing-library-reference");
};

/** Takes Word off again, whatever state the run ended in. Never throws: this is cleanup. */
async function removeWord() {
  if (!(await hasWord())) { return; }
  await api.immediate(
    'ThisWorkbook.VBProject.References.Remove ThisWorkbook.VBProject.References("Word")')
    .catch(() => null);
}

try {
  await removeWord();
  await api.component("add", { name: MODULE, kind: "standard", project: projectId });
  await api.writeModule(MODULE, SOURCE, projectId);

  await waitFor("the error to be reported", async () => (await missingNow()).length === 1)
    .catch(() => null);

  check("a project that does not reference Word is told so, on the line that names it",
    (await missingNow()).length === 1,
    (await missingNow()).map((one) => `${one.code}@${one.line}`).join(", ") || "(none)");

  check("and the project's reference list says the same thing",
    !(await hasWord()),
    (await referencesNow()).map((one) => one.name).join(", "));

  // THE LIGHTBULB, not the engine's answer: what the developer is actually offered. It reads the
  // ACTIVE editor, so the module comes forward first - the order a developer works in.
  await api.act("activate", { module: MODULE, project: projectId });
  await wait(600);
  const offered = await api.act("quickFixes", { line: FINDING_LINE, column: 15 });
  const titles = (offered.data ?? []).map((one) => one.title);

  check("the lightbulb offers the fix that ADDS the library, not only the one that hides it",
    titles.includes("Add a reference to the Word object library"),
    titles.join(" | ") || "(nothing offered)");

  check("and offers it first, because suppressing a compile error is the last resort",
    (offered.data ?? []).find((one) => one.isPreferred)?.title
      === "Add a reference to the Word object library",
    (offered.data ?? []).map((one) => `${one.title} preferred=${one.isPreferred}`).join(" | "));

  // Applying it, by the finding rather than by naming a library: the route runs the fix offered
  // there, with the identity the engine attached to it.
  const added = await api.addReference(MODULE, FINDING_LINE, { project: projectId });
  check("applying it answers which library went on",
    added.ok === true && added.added === true && added.library === "Word",
    JSON.stringify(added));

  check("the project references Word now, by the identity the fix carried",
    (await referencesNow()).some((one) =>
      one.name === "Word" && one.guid.toUpperCase() === "{00020905-0000-0000-C000-000000000046}"),
    (await referencesNow()).map((one) => one.name).join(", "));

  // THE HALF THAT WAS WRONG. The module was never touched, so nothing that compares text would
  // re-seed or re-analyse, and the error stood on screen until something else was typed.
  await waitFor("the finding to clear", async () => (await missingNow()).length === 0)
    .catch(() => null);
  check("and the error clears with the module never touched",
    (await missingNow()).length === 0,
    (await missingNow()).map((one) => one.code).join(", ") || "(none)");

  const again = await api.addReference(MODULE, FINDING_LINE, { project: projectId })
    .catch((error) => ({ error: String(error.message ?? error) }));
  check("asking a second time is refused in words - the fix is no longer offered there",
    typeof again.error === "string" && again.error.length > 0,
    JSON.stringify(again));
} finally {
  await removeWord();
  await api.component("remove", { name: MODULE, project: projectId }).catch(() => null);

  const left = (await api.project(projectId)).components.some((one) => one.name === MODULE);
  check("the module it brought was taken away", !left, left ? `${MODULE} is still there` : undefined);
  check("and so was the reference", !(await hasWord()),
    (await referencesNow()).map((one) => one.name).join(", "));

  process.exitCode = done();
}
