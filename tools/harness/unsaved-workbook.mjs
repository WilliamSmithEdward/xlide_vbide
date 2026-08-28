/*
 * The never-saved workbook, which no fixture can be: every fixture is a file, and Book1
 * exists only while a session holds it. The state hid two defects until the 2026-08-28 hunt:
 * the tree and every route called the workbook "VBAProject" while Excel's every surface said
 * "Book1", so the name a developer would use resolved nowhere (#14) - and the only route with
 * the reach to close it, an immediate line, was the one route that dies trying: its line runs
 * as a procedure IN the active project, and closing that project's workbook tore the editor
 * down with it, one try, one dead Excel (#13).
 *
 * Runs against whatever fixture is live. The workbook it makes is its own, and closing it
 * goes through the route this suite exists to drive.
 *
 *   node tools\harness\unsaved-workbook.mjs
 */

import { open, reporter, waitFor } from "./xlide-api.mjs";

const api = await open();
const { check, done } = reporter();
const fixture = (await api.project()).projectId;

try {
  await api.immediate("Workbooks.Add");
  await waitFor("the new workbook to reach the tree", async () =>
    (await api.projects()).projects.length >= 2);

  const rows = (await api.projects()).projects;
  const book = rows.find((one) => /^Book\d+$/i.test(one.project));
  check("an unsaved workbook answers by the host's own name for it",
    book !== undefined,
    rows.map((one) => one.project).join(", "));

  // Resolution rides the same axis: the name every native surface speaks is a project
  // argument now. The add is also this suite's carrier, so the close below has something
  // beyond a blank workbook to discard.
  const name = book?.project ?? "Book1";
  const added = await api.component("add", { kind: "module", name: "UnsavedCarrier", project: name })
    .then(() => true, (error) => error.message);
  check(`and '${name}' resolves as a project argument`, added === true, String(added));

  // The evaluator's one sharp edge, refused in words. The red proof is deliberately not
  // runnable here: it is the 2026-08-28 crash itself, and these words are what stand between
  // a session and repeating it.
  const refused = await api.immediate("Workbooks(2).Close False");
  check("an immediate line closing a workbook is refused in words",
    refused.failed === true && /does not survive/.test(refused.text ?? ""),
    JSON.stringify(refused).slice(0, 140));
  check("and the session is alive to say so",
    (await api.stats()).heartbeatAgeMs < 5000,
    `heartbeat ${(await api.stats()).heartbeatAgeMs}ms`);

  // The route the refusal points at: the host's own File Close, the save question answered in
  // the request. Discard, because the workbook is this suite's own scratch.
  const closed = await api.workbook("close", { project: name, saveChanges: 0 });
  check("workbook close through the door closes it",
    closed.closed === true, JSON.stringify(closed));

  await waitFor("the tree to let the workbook go", async () =>
    (await api.projects()).projects.length === 1);
  check("and the fixture is the one project left standing",
    (await api.projects()).projects[0]?.project?.toLowerCase() === fixture.toLowerCase(),
    (await api.projects()).projects.map((one) => one.project).join(", "));
} finally {
  // A failed run may leave Book1 standing. The close is the cleanup AND the feature, so a
  // second attempt costs nothing, and a workbook already gone refuses harmlessly.
  const leftover = (await api.projects().catch(() => ({ projects: [] }))).projects
    .find((one) => /^Book\d+$/i.test(one.project));
  if (leftover) {
    await api.workbook("close", { project: leftover.project, saveChanges: 0 }).catch(() => {});
  }
}

done();
