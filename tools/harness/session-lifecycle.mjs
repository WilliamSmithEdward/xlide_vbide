/*
 * A shutdown begun and then cancelled, driven through the door.
 *
 * THE FIELD FAILURE THIS GUARDS (2026-08-02). Excel says OnBeginShutdown, the session stops, and
 * THEN the host asks about unsaved changes. Cancel abandons the whole shutdown with no callback
 * that ever says so - the host keeps running and the add-in is a corpse, menu-less and dead. The
 * shutdown watchdog is the one thing left alive to notice and stand the session up again, and
 * until this route existed the only way to exercise it was to close Excel by hand and press
 * Cancel, which no gate can do.
 *
 * `session cancelledShutdown` runs the real OnBeginShutdown without a process exit, so the frame
 * stays standing, the watchdog reads a cancellation, and the session revives - a NEW session with
 * a fresh port and startedAt written into the discovery file. This suite drives that and proves
 * the session that comes back is healthy, which is the release-blocking property that had no test.
 *
 * Run against any live session, ALONE (it tears the session down and back up):
 *   node tools\harness\session-lifecycle.mjs
 */
import { discover, reporter, wait } from "./xlide-api.mjs";

const { check, done } = reporter();

const sole = async () => {
  const all = await discover();
  if (all.length === 0) { throw new Error("no live session to drive"); }
  if (all.length > 1) {
    throw new Error(`several sessions are live (pids ${all.map((e) => e.pid).join(", ")}); run this alone`);
  }
  return all[0];
};

console.log("a cancelled shutdown, and the revival that guards it\n");

// SUBSTANTIVE health, not the composite `healthy`. The doctor folds in a build-gap heuristic -
// "the shim and page bundle were built N minutes apart" - which is a fact about the dev
// workflow, not the session: it is routinely true on a machine where the shim was published and
// the page deployed at different times, and it says nothing about whether the session works. The
// session being functionally alive is engine up, surface ready, ghost readers attached, which is
// what a revival has to restore.
const alive = async (api, budgetMs = 25000) => {
  const deadline = Date.now() + budgetMs;
  while (Date.now() < deadline) {
    const d = await api.doctor().catch(() => null);
    if (d && d.engineUp && d.surfaceReady && d.ghostReadersUp) { return d; }
    await wait(500);
  }
  return null;
};

const before = await sole();
console.log(`  session pid ${before.pid}, port ${before.port}, started ${before.startedAt}`);

check("the session is functionally alive before the shutdown", (await alive(before.api)) !== null);

// The reply must arrive BEFORE the teardown - it rides the DebugServer that Stop() disposes.
const said = await before.api.session("cancelledShutdown");
console.log(`  ${JSON.stringify(said)}`);
check("the shutdown was accepted, with a reply that beat the teardown", said.ran === true, JSON.stringify(said));

/*
 * The session goes DOWN, then comes back as a NEW one. Proof it actually cycled, rather than
 * that nothing happened, is the discovery file's startedAt moving on: only a fresh DebugServer
 * writes a fresh one, and only a revival starts a fresh DebugServer. Polled, because the
 * teardown and the two watchdog ticks take a few seconds, and there is a window mid-cycle when
 * the pid's session answers nothing at all.
 */
const revived = await (async () => {
  const deadline = Date.now() + 30000;
  let sawDown = false;
  while (Date.now() < deadline) {
    await wait(500);
    const now = (await discover()).find((e) => e.pid === before.pid);
    if (!now) { sawDown = true; continue; }
    if (now.startedAt !== before.startedAt) { return { instance: now, sawDown }; }
  }
  return null;
})();

check("the session went away and a NEW one came back",
  revived !== null,
  revived ? "" : "no revived session of this pid within 30s; the watchdog did not stand it up");

if (revived) {
  console.log(`  revived: port ${revived.instance.port}, started ${revived.instance.startedAt}`);
  // Waited for, not asked once: a session seconds old is still seeding - the engine is
  // reconnecting and the host thread is busy with its fresh HostStartupComplete, so an
  // immediate doctor call times out on the host thread. This is the SAME readiness the field
  // failure was about: the add-in came back DEAD, and "dead" is precisely engine down, surface
  // not ready, ghosts not attached. That it becomes alive again is the whole property.
  check("the revived session becomes functionally alive again", (await alive(revived.instance.api)) !== null,
    "the session revived but never became healthy - the add-in is back but dead, which is the field failure");
  check("it is a fresh server, not the old one still answering",
    revived.instance.startedAt !== before.startedAt,
    `startedAt ${before.startedAt} -> ${revived.instance.startedAt}`);
  // Not asserted: a fast revival can close the down-window between polls, so seeing it is a
  // bonus rather than a requirement. The startedAt move is the proof that matters.
  console.log(`     (the teardown window was observed: ${revived.sawDown})`);
}

process.exit(done());
