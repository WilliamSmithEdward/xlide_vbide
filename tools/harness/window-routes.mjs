/*
 * The window routes: the palette put away, the editor closed and brought back, and a HIDDEN
 * native pane closed - each through the door, each observed through the door.
 *
 * Three gestures a developer makes daily had no api at all until 2026-08-12, and the probes
 * that tested them reached through Application.VBE or sent window messages from outside the
 * process - the first of which project trust gates, and this machine runs with that trust OFF.
 * These routes are what let their defect classes into the gate:
 *
 *   frame close    lesson 27: closing the editor once killed Excel
 *   frame show     the palette greeting the next Alt+F11 uninvited (2026-08-05)
 *   closeNative    the 2026-08-04 dead tab: a hidden pane's close left its tab standing
 *
 * Runs against the DebugFixture session the live probes share, after objbrowser-live-probe
 * (which leaves a palette existing, hidden) and before Test-ResizeFollow.
 */
import { open, reporter, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();

const { check, done } = reporter();

const state = () => api.state();

console.log("the window routes, against the live session\n");

check("the frame is on screen to begin with", (await state()).frameVisible === true);

// ---- the palette, summoned and put away ----

await api.command("objectBrowser");
await waitFor("the palette to be showing", async () => (await state()).paletteVisible === true);
check("the objectBrowser command summons the palette", (await state()).paletteVisible === true);

const hid = await api.paletteHide();
check("palette hide answers that it did", hid.did === true, JSON.stringify(hid));
await waitFor("the palette to be hidden", async () => (await state()).paletteVisible === false);
const afterHide = await state();
check("and the palette is hidden with its state intact",
  afterHide.paletteVisible === false && afterHide.paletteOpen === true,
  `paletteOpen=${afterHide.paletteOpen} paletteVisible=${afterHide.paletteVisible}`);

// ---- the frame closed, with the palette standing, and the follow contract ----

await api.command("objectBrowser");
await waitFor("the palette to be back for the close test", async () =>
  (await state()).paletteVisible === true);

const closed = await api.frame("close");
check("frame close answers posted, not done", closed.did === true && closed.visible === true,
  `${JSON.stringify(closed)} - the pump delivers SC_CLOSE after the reply, so visible must still read true here`);

await waitFor("the frame to leave the screen", async () => (await state()).frameVisible === false);

// THE PALETTE GOES DOWN SEPARATELY, and waiting for the frame is not waiting for it. The two
// are one gesture but two observations, and reading both from the snapshot that satisfied the
// FRAME's wait is a race the gate lost once (2026-08-29: green standalone and on the next run,
// red in between, with nothing changed).
//
// Waited for, and the wait may NOT throw - which is the whole point. A wait that names the
// assertion launders it into the setup and leaves a check that can only ever time out; this
// one gives the palette a bounded moment and then lets the check below do the judging, so a
// palette that never goes down is still a FAILURE with its state printed, not a stack.
await waitFor("the palette to follow the frame down", async () =>
  (await state()).paletteVisible === false, { budgetMs: 5000 }).catch(() => null);

const whileClosed = await state();
check("the frame is off screen", whileClosed.frameVisible === false);
check("and the palette followed it down, not left floating over the workbook",
  whileClosed.paletteVisible === false,
  "an owned window does not follow a hidden owner by itself; the editor's close path must hide it");

// ---- the frame back, and the palette staying away until summoned ----

const shown = await api.frame("show");
check("frame show answers with the outcome", shown.did === true && shown.visible === true,
  JSON.stringify(shown));
await waitFor("the frame to be back", async () => (await state()).frameVisible === true);
const afterShow = await state();
check("the palette STAYS away when the editor returns",
  afterShow.paletteVisible === false,
  "it returns only when summoned; greeting the next Alt+F11 uninvited was the 2026-08-05 defect");

// ---- a HIDDEN native pane closed through the door ----

const scratch = `WinRoutes${process.pid}`;
await api.component("add", { kind: "module", name: scratch, project: project.projectId });
try {
  await api.pane("open", { module: scratch, project: project.projectId });
  await api.pane("open", { module: "Runner", project: project.projectId });
  await waitFor("Runner in front with the scratch pane hidden behind it", async () => {
    const ui = await api.ui();
    const tabs = ui.workspace.groups.flatMap((g) => g.tabs).map((t) => t.module);
    return tabs.includes(scratch) && tabs.includes("Runner")
      && (await state()).shownModule === "Runner";
  });

  const native = await api.pane("closeNative", { module: scratch, project: project.projectId });
  check("closing the hidden pane's window answers that it closed", native.closed === true,
    JSON.stringify(native));

  // The 2026-08-04 assertion: the tab GOES, because the tracker notices the window go and
  // republishes - not because anything told the page.
  await waitFor("the hidden pane's tab to leave the strip", async () => {
    const ui = await api.ui();
    return !ui.workspace.groups.flatMap((g) => g.tabs).some((t) => t.module === scratch);
  });
  const stripAfter = (await api.ui()).workspace.groups.flatMap((g) => g.tabs).map((t) => t.module);
  check("the hidden pane's tab is gone", !stripAfter.includes(scratch), stripAfter.join(","));
  check("and Runner's tab survived it", stripAfter.includes("Runner"), stripAfter.join(","));
  check("and the shown module did not change - a hidden pane's close steals nothing",
    (await state()).shownModule === "Runner");

  const again = await api.pane("closeNative", { module: scratch, project: project.projectId });
  check("closing a pane that is not there is refused, not shrugged at",
    again.closed === false && /no pane/.test(again.detail ?? ""), JSON.stringify(again));
} finally {
  await api.component("remove", { name: scratch, project: project.projectId }).catch((error) =>
    console.log(`     WARNING: ${scratch} could not be removed (${error.message})`));
}

process.exit(done());
