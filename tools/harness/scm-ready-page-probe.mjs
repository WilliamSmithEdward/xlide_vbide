// The Source Control pane over a repository, headless: served as /index.html?scm=1 the demo
// transport holds a fake repository - two rows, a Folder row, two commits, diff and show rows,
// blame, a commit, an import, a checkout refused while dirty and a past-version tab - so the
// pane's every gesture can be driven without a host. What only a real repository can answer is
// scm.mjs's business; this proves the pane draws, keeps and takes down what the host says.
//
//   node tools\harness\scm-ready-page-probe.mjs      (from PowerShell; see page-probe.mjs)

import { runPageProbe } from "./page-probe.mjs";

const DRIVE = `(async () => {
  const checks = [];
  try {
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });
  const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
  const until = async (what, budget = 5000) => {
    for (let waited = 0; waited < budget; waited += 100) {
      const got = what();
      if (got) { return got; }
      await sleep(100);
    }
    return null;
  };
  const idle = () => until(() => (ui.state().scm.busy ? null : true));

  await until(() => document.querySelectorAll(".tab").length >= 2, 20000);
  const ui = window.xlideUi;
  const root = document.querySelector("#scm");
  document.querySelector('.panel-tab[data-panel="scm"]').click();

  const ready = await until(() => { const s = ui.state().scm; return s && s.state === "ready" && s.commits.length === 2 && !s.busy ? s : null; });
  check("the pane reads the fake repository: ready, two rows, two commits", !!ready, JSON.stringify(ready));
  check("rows are ticked by default", ready && ready.ticked.length === 2, ready && JSON.stringify(ready.ticked));
  check("the message box is pre-filled from the suggestion", ready && ready.message === "developer: Written Module1, Added Module2", ready && ready.message);
  check("the branch select offers both branches and shows main", ready && ready.branch === "main" && ready.branches.join(",") === "main,feature" && root.querySelector("#scm-branch").value === "main");
  check("the Folder section lists the outside row", ready && ready.outside.length === 1 && ready.outside[0].status === "folderNewer");
  check("import is enabled with a Folder row; commit is enabled with ticks and a message", !root.querySelector("#scm-import").disabled && !root.querySelector("#scm-commit").disabled);
  check("blame is enabled over a repository", !root.querySelector("#scm-blame").disabled);

  const untick = ui.act("scmPane", { tick: "Module2", on: false });
  check("act tick off unticks a row", untick.did === true && ui.state().scm.ticked.join(",") === "Module1", untick.detail);

  const showRow = ui.act("scmPane", { module: "Module1" });
  const rowsDrawn = await until(() => root.querySelectorAll("#scm-diff .sync-diff-row").length === 3);
  check("clicking a Changes row draws the host's diff rows with the sync classes", showRow.did === true && !!rowsDrawn, showRow.detail);
  check("ui.scm.showing names the row", ui.state().scm.showing === "Module1");
  check("the showing row carries the selection cue", !!root.querySelector('.scm-entry[data-module="Module1"].scm-entry-showing'));

  // A refresh redraws the list; the comparison being read stays up, re-asked rather than wiped.
  await idle();
  const refreshed = ui.act("scmPane", { press: "refresh" });
  await idle();
  await sleep(300);
  check("a refresh keeps the comparison on screen",
    refreshed.did === true && root.querySelectorAll("#scm-diff .sync-diff-row").length === 3 && ui.state().scm.showing === "Module1",
    JSON.stringify({ rows: root.querySelectorAll("#scm-diff .sync-diff-row").length, showing: ui.state().scm.showing }));
  const whileBusy = (() => { const first = ui.act("scmPane", { press: "refresh" }); const second = ui.act("scmPane", { press: "refresh" }); return { first, second }; })();
  check("a second refresh while the first is in flight is refused, the button being disabled",
    whileBusy.first.did === true && whileBusy.second.did === false, JSON.stringify(whileBusy));
  await idle();

  const openCommit = ui.act("scmPane", { commit: "a1b2c3d" });
  check("act commit opens a commit", openCommit.did === true && ui.state().scm.commits.find((c) => c.short === "a1b2c3d").open === true, openCommit.detail);
  const showFile = ui.act("scmPane", { commit: "a1b2c3d", module: "Module1" });
  const fileHead = await until(() => root.querySelector("#scm-open-version") && root.querySelector("#scm-restore"));
  check("clicking a commit's file draws the comparison with Open this version and Restore", showFile.did === true && !!fileHead, showFile.detail);
  check("ui.scm.showing names the file at the commit", ui.state().scm.showing === "Module1@a1b2c3d");

  const open = ui.act("scmPane", { press: "open" });
  const historyTab = await until(() => {
    const tabs = ui.state().workspace.groups.flatMap((g) => g.tabs);
    return tabs.find((t) => t.face === "history:a1b2c3d") ?? null;
  });
  check("Open this version lands a host-listed history tab", open.did === true && !!historyTab, JSON.stringify(historyTab));
  check("the history tab is labelled module @ short", historyTab && historyTab.label === "Module1 @ a1b2c3d", historyTab && historyTab.label);
  check("the history tab is active", historyTab && historyTab.active === true);
  const stripTab = document.querySelector('.tab[data-face="history:a1b2c3d"]');
  check("the strip draws it with the history class", !!stripTab && stripTab.classList.contains("history"));
  const view = await until(() => document.querySelector('.history-view[data-short="a1b2c3d"]'));
  check("a history view is mounted over the group", !!view && view.isConnected);
  const text = await until(() => { const m = view && view.querySelector(".history-editor .view-lines"); const t = m && m.textContent.replace(/\\u00a0/g, " "); return t && t.includes("at a1b2c3d") ? t : null; });
  check("the read-only editor shows the text the host answered", !!text, text && text.slice(0, 60));
  const census = ui.state().census;
  check("the census excludes the history model", census.models === census.documents, JSON.stringify(census));

  const closed = await ui.act("closeActive", {});
  const gone = await until(() => ui.state().workspace.groups.flatMap((g) => g.tabs).every((t) => !t.face) ? true : null);
  check("closing the history tab removes it through the host's list, and the act says so", closed.did === true && !!gone, JSON.stringify(closed));
  const disposed = await until(() => document.querySelector('.history-view[data-short="a1b2c3d"]') === null ? true : null);
  check("the history view is disposed when the host's list drops it", !!disposed);
  check("the census still balances after the dispose", ui.state().census.models === ui.state().census.documents, JSON.stringify(ui.state().census));

  ui.act("activate", { module: "Module1" });
  await until(() => ui.state().workspace.active && ui.state().workspace.active.module === "Module1" && !ui.state().workspace.active.face ? true : null);
  const blamed = await ui.act("blame", { which: "on" });
  const painted = await until(() => { const b = ui.state().blame; return b.on && b.lines === 2 && b.uncommitted === 2 && b.painted === 4 ? b : null; });
  check("blame on paints two committed lines and two uncommitted", blamed.did === true && !!painted, JSON.stringify(painted ?? ui.state().blame));
  // The hints are drawn for the lines on screen; a render is asked for before the DOM is read.
  ui.act("reveal", { line: 1 });
  await sleep(300);
  const hint = await until(() => { const spans = [...document.querySelectorAll(".xlide-blame")]; return spans.length >= 2 ? spans.map((s) => s.textContent.replace(/\\u00a0/g, " ")) : null; });
  check("the injected hints read author, date, short", !!hint && hint.some((h) => h.includes("Demo Developer, 2026-09-01, a1b2c3d")), JSON.stringify(hint));
  const notCommitted = [...document.querySelectorAll(".xlide-blame-uncommitted")].map((s) => s.textContent.replace(/\\u00a0/g, " "));
  check("uncommitted lines say not committed", notCommitted.length === 2 && notCommitted.every((h) => h.includes("not committed")), JSON.stringify(notCommitted));
  check("the pane's Blame button reads pressed", root.querySelector("#scm-blame").getAttribute("aria-pressed") === "true");

  ui.act("activate", { module: "Module2" });
  await until(() => ui.state().workspace.active && ui.state().workspace.active.module === "Module2" ? true : null);
  check("blame is per document: Module2 is off", ui.state().blame.on === false);
  ui.act("activate", { module: "Module1" });
  await until(() => ui.state().workspace.active && ui.state().workspace.active.module === "Module1" ? true : null);
  check("and Module1 is still on after the switch", ui.state().blame.on === true && ui.state().blame.lines === 2);
  await ui.act("blame", { which: "off" });
  check("blame off clears the hints", document.querySelectorAll(".xlide-blame, .xlide-blame-uncommitted").length === 0 && ui.state().blame.on === false);

  await idle();
  ui.act("scmPane", { message: "Demo commit from the probe" });
  const commit = ui.act("scmPane", { press: "commit" });
  const committed = await until(() => { const s = ui.state().scm; return s.commits.length === 3 && s.rows.length === 1 && !s.busy ? s : null; });
  check("commit sends the ticked row: one row left, three commits", commit.did === true && !!committed, JSON.stringify(committed && { rows: committed.rows.map((r) => r.module), commits: committed.commits.length, detail: committed.detail }));
  check("the notice carries the action's words", committed && committed.detail.startsWith("committed 1 module"), committed && committed.detail);
  check("the message box was cleared and refilled from the new suggestion", committed && committed.message === "developer: Written Module1, Added Module2", committed && committed.message);

  await idle();
  const branch = ui.act("scmPane", { branch: "feature" });
  const refused = await until(() => { const s = ui.state().scm; return !s.busy && /unsaved/.test(s.detail) ? s : null; });
  check("a checkout over a dirty workbook is refused in words and the select snaps back", branch.did === true && !!refused && root.querySelector("#scm-branch").value === "main", refused && refused.detail);

  await idle();
  const imported = ui.act("scmPane", { press: "import" });
  const clean = await until(() => { const s = ui.state().scm; return !s.busy && s.outside.length === 0 ? s : null; });
  check("import clears the Folder rows and disables Import", imported.did === true && !!clean && root.querySelector("#scm-import").disabled, imported.detail);

  await idle();
  ui.act("scmPane", { commit: "a1b2c3d", module: "Module1" });
  await until(() => root.querySelector("#scm-restore"));
  const restorePress = ui.act("scmPane", { press: "restore" });
  const outcome = await until(() => document.querySelector("#scm-restore-outcome") && document.querySelector("#scm-restore-outcome").textContent.includes("written") ? document.querySelector("#scm-restore-summary").textContent : null);
  check("restore asks, runs and shows the outcome in its card", restorePress.did === true && !!outcome, outcome);
  document.querySelector("#scm-restore-done")?.click();
  check("the card closes", document.querySelector("#scm-restore-card") === null);

  return { pass: checks.every((one) => one.ok), checks };
  } catch (error) {
    checks.push({ name: "the drive threw", ok: false, detail: String(error && error.stack || error) });
    return { pass: false, checks };
  }
})()`;

await runPageProbe({ label: "xlide-scm-ready", drive: DRIVE, path: "/index.html?scm=1" });
