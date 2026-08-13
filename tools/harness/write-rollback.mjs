/*
 * Does a write the editor refuses cost the module the text it already had?
 *
 * THE DEFECT THIS WAS WRITTEN AGAINST. A whole-module replace is two calls - delete every line,
 * then add the new text - and the editor can refuse the second one. On 2026-08-09 a module holding
 * 2,002 lines of working code was asked to take a body the editor would not have, and came back
 * holding 31,956 lines of the new one: not the old body, not the new body, and the route replied
 * ok. Nothing said a word except one line in the log.
 *
 * The writer now keeps the old text in hand and puts it back, and returns the editor's complaint
 * instead of swallowing it. This proves both, against the real editor, which is the only place the
 * refusal exists.
 *
 * WHAT RUNNING THIS COSTS THE SESSION, and it is not a matter of taste. The only refusal cheap
 * enough to provoke on demand is a module pushed past the editor's identifier budget, and that
 * leaves the WHOLE VBE unable to add a component - "Insufficient memory to continue the
 * execution of the program" - for the rest of the session. Removing the module does not give the
 * memory back. Nothing else refuses: a 200,000-character line is accepted and silently broken
 * into 197, a 60,000-character constant into 60, a null character is taken as text, and 20,001
 * lines of procedures go in without complaint. So this costs an Excel restart, every time, by
 * construction.
 *
 * That is why the gate runs it LAST in its fixture group and nowhere else: the group's session
 * is discarded by a fresh relaunch immediately after, which is the restart this suite demands.
 * For one run it sat mid-group instead, and every suite after it met a session that refused
 * every add (2026-08-12). Anywhere else it runs:
 *
 *   node tools\harness\write-rollback.mjs
 *
 * Then restart Excel. The script says so at the end, and means it.
 */
import { open, reporter } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const name = `RollbackProbe${Date.now().toString().slice(-5)}`;

/** The door gives the host thread three seconds; a large write outlives that and goes on. */
const tolerate = (error) => {
  if (!/did not answer in time|aborted/i.test(String(error?.message))) { throw error; }
  return null;
};

const linesOf = async () =>
  (await api.project().catch(tolerate))?.components?.find((c) => c.name === name)?.lines ?? null;

/** Polls the object model until the count stops moving, so a slow write is waited out. */
async function settle(budgetMs = 180_000) {
  let last = null;
  let still = 0;
  for (let waited = 0; waited <= budgetMs; waited += 1500) {
    await new Promise((r) => setTimeout(r, 1500));
    const now = await linesOf();
    still = now !== null && now === last ? still + 1 : 0;
    last = now;
    if (still >= 2) { break; }
  }

  return last;
}

const { check, done } = reporter();

/** The shape the perf fixture uses, which the editor takes to 65,000 lines without complaint. */
const working = (procedures) => {
  const out = ["Option Explicit", ""];
  for (let n = 0; n < procedures; n++) {
    out.push(`Public Function Keep${n}(ByVal seed As Long) As Long`, `    Keep${n} = seed + ${n}`, "End Function", "");
  }
  return out.join("\r\n");
};

/** The shape the editor refuses: module-level constants past its own identifier budget. */
const refusedBody = (count) => {
  const out = ["Option Explicit"];
  for (let n = 0; n < count; n++) { out.push(`Const K${n} As Long = ${n}`); }
  return out.join("\r\n");
};

let made = false;
try {
  await api.component("add", { kind: "module", name, project: project.projectId });
  made = true;

  const good = working(500);
  const wanted = good.split("\r\n").length;
  await api.mark(`write-rollback: ${wanted} lines the editor accepts`);
  await api.writeModule(name, good, project.projectId).catch(tolerate);
  const held = await settle();

  check("the module took a body the editor accepts", held === wanted,
    `asked for ${wanted} lines and the workbook holds ${held}. The rest of this reads against that, so it stops here.`);

  if (held !== wanted) { throw new Error("the fixture never got its starting body"); }

  await api.mark("write-rollback: a body the editor refuses");
  console.log(`\n  writing a body the editor will refuse, over ${held} working lines`);

  let reported = null;
  await api.writeModule(name, refusedBody(50_000), project.projectId).catch((error) => {
    if (/did not answer in time|aborted/i.test(String(error?.message))) { return; }
    reported = error.message;
  });

  const after = await settle();

  check("a refused write leaves the module holding what it had", after === held,
    `${held} lines became ${after}. The delete landed and the add did not, so what is in the module `
    + "is neither the developer's code nor the text that was sent. This is the defect.");

  const text = (await api.readModule(name, project.projectId).catch(() => null))?.text ?? "";
  check("and holding the same text, not just the same count", text.includes("Public Function Keep0"),
    `the module no longer contains Keep0. It starts: ${text.slice(0, 80)}`);

  // The complaint reaches the caller when the write finishes inside the door's budget. This one
  // does not - it takes some twenty seconds - so the LOG is where the report is read from here.
  const log = await api.log({ match: "was refused", max: 20 }).catch(() => null);
  const said = (log?.lines ?? []).map((l) => (typeof l === "string" ? l : l.text ?? "")).join("\n");
  check("the writer says what the editor said, and that it put the text back",
    /was refused and its previous text was put back/.test(said),
    reported
      ? `the route reported "${reported}" but the log does not say the text was restored`
      : `nothing in the log names the refusal. Last lines:\n     ${said.slice(-300)}`);
} finally {
  if (made) {
    await api.component("remove", { name, project: project.projectId }).catch(() => {});
  }

  process.exitCode = done();
  console.log("\nRESTART EXCEL NOW. The editor has been pushed past its identifier budget and will");
  console.log("refuse to add a component until it is restarted. That is what provoking this costs.");
}
