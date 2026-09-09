/*
 * Source control: git.exe behind the folder a project's modules are exported to, and the pane
 * that works it.
 *
 * WHAT IT IS FOR. Rubberduck's source control panel was its most requested feature and never
 * shipped working. This product has the pieces every earlier attempt lacked - a faithful export,
 * a change log that already models rounds and restore, a comparison renderer the Changes pane
 * shares - and git turns the session's log into something durable, branchable and shared. The
 * questions only a live repository can answer:
 *
 *   - are the rows the LIVE project against the branch head, and does a commit clear them?
 *   - does a commit take the ticked modules and leave the rest, and does the log say so?
 *   - does blame map the committed file's lines onto the editor's lines, and mark the rest?
 *   - does a restore from a commit land as a change-log round, so it can be restored away?
 *   - does an edit made in the FOLDER show up on its own, and import when asked?
 *   - does a checkout refuse over a dirty workbook and import once it is clean?
 *   - does the pane draw what the route answers, or has one of the two drifted?
 *
 * The repository is a temporary folder of this suite's own, initialised through the route,
 * committed, branched, edited and deleted here. The workbook is ScmFixture.xlsm, whose three
 * modules name themselves in a string so a blame or a comparison that fetched the wrong module's
 * text reads as the wrong module. Nothing here touches any other workbook, and the fixture is put
 * back at the end.
 *
 * IT NEEDS git.exe, and it says so rather than skipping: a suite that goes green on a machine
 * that cannot run the feature is a gate that cannot fail.
 *
 * Run against ScmFixture.xlsm with the editor open:
 *   tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\ScmFixture.xlsm
 *   node tools\harness\scm.mjs
 */

import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, realpathSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { open, waitFor, comparingReporter } from "./xlide-api.mjs";

const api = await open();
const { check, done } = comparingReporter();

const project = (await api.projects()).projects[0];
console.log(`project: ${project.project}\n`);

// IT WRITES, COMMITS, IMPORTS AND SAVES, so it checks whose workbook it is in. Every check
// below rewrites a module, saves the workbook, or checks a branch out over it - safe in the
// fixture built for it and somebody's actual work anywhere else.
if (!/scmfixture/i.test(project.project)) {
  console.log("FAIL this suite commits, restores, imports and saves the modules of whatever workbook"
    + `\n     is open, and this one is ${project.project}, which is not the fixture it was built`
    + "\n     for. Refusing."
    + "\n\n     powershell tools\\harness\\Start-Excel.ps1 -Fresh -Workbook artifacts\\fixtures\\ScmFixture.xlsm");
  console.log("\n0 passed, 1 failed");
  process.exit(1);
}

// A project that is running or stopped refuses writes - rightly - and the first refusal would
// otherwise be an unhandled throw under a libuv teardown assertion rather than a sentence.
const modeNow = (await api.state()).debugMode ?? "design";
if (modeNow !== "design") {
  console.log(`FAIL the project is in ${modeNow} mode, so nothing here can write. Is a form or`
    + " a run standing? Stop it (or POST command?name=reset), then run this again.");
  console.log("\n0 passed, 1 failed");
  process.exit(1);
}

// GIT, OUT LOUD. The product answers `noGit` and carries on; this suite cannot, because every
// question it asks is about a repository. Named rather than skipped: the gate has no notion of
// a suite that is excused when a tool is missing, and it should not - a green run on a machine
// without git would be a run that measured nothing.
//
// FOUND THE WAY THE PRODUCT FINDS IT: PATH first, then Git for Windows' three standard folders,
// so a machine where the product works is a machine where this suite runs.
const gitCandidates = ["git", ...[
  process.env.ProgramFiles,
  process.env["ProgramFiles(x86)"],
  process.env.LOCALAPPDATA ? join(process.env.LOCALAPPDATA, "Programs") : undefined,
].filter(Boolean).map((root) => join(root, "Git", "cmd", "git.exe"))];
let gitExe = null;
let gitVersion = "";
for (const candidate of gitCandidates) {
  try {
    gitVersion = execFileSync(candidate, ["--version"], { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] }).trim();
    gitExe = candidate;
    break;
  } catch {
    // Not there; the next place is tried.
  }
}
if (gitExe === null) {
  console.log("FAIL git.exe was not found on PATH or in Git for Windows' standard folders, and this"
    + "\n     suite is about a git repository. Install Git for Windows (git-scm.com/download/win),"
    + "\n     per user, no administrator rights, then run this again.");
  console.log("\n0 passed, 1 failed");
  process.exit(1);
}
console.log(`${gitVersion} (${gitExe})\n`);

const git = (...args) => execFileSync(gitExe, ["-C", folder, ...args], { encoding: "utf8" }).trim();

// THE LONG PATH, not the short one. %TEMP% on Windows is routinely spelt with a ~1 name, and
// `git rev-parse --show-toplevel` answers the real one - so the folder is canonicalised once
// here and sent to the product already long, and every path comparison below goes through the
// same door.
const canonical = (path) => {
  try {
    return realpathSync.native(path).toLowerCase().replace(/\//g, "\\");
  } catch {
    return resolve(path).toLowerCase().replace(/\//g, "\\");
  }
};

const scratch = join(tmpdir(), `xlide-scm-${process.pid}`);
rmSync(scratch, { recursive: true, force: true });
mkdirSync(scratch, { recursive: true });
const folder = realpathSync.native(scratch);
console.log(`folder: ${folder}\n`);

// A bare repository beside the folder stands in for a remote: no credentials, no network, and
// push, fetch and pull run for real against it.
const remote = `${folder}-remote`;
rmSync(remote, { recursive: true, force: true });

const held = (name) => api.readModule(name, project.projectId).then((one) => one.text);
const write = (name, text, by) => api.writeModule(name, text, project.projectId, { by });
const scm = (args = {}) => api.scm({ ...args, project: project.projectId });
const log = (args = {}) => api.changes({ ...args, project: project.projectId });
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Windows refuses to delete a folder that is some process's working directory, and git runs
// with its own in the repository: a round the pane started before the forget can still be out
// with git when the cleanup reaches the folder. rmSync's own retries never fire for that: asked
// for ten of them, it threw EPERM after 0ms against a folder a process sat in (2026-09-08), so
// this asks again itself, for up to ten seconds.
const removeWhenFree = async (path) => {
  for (let attempt = 0; ; attempt += 1) {
    try {
      rmSync(path, { recursive: true, force: true });
      return null;
    } catch (error) {
      if (attempt >= 40) {
        return error.message;
      }

      await sleep(250);
    }
  }
};

// Every action answers the status after acting; commit and import carry it under `status`,
// which is how the pane redraws after them. One reader for both shapes.
const statusOf = (reply) => reply.status ?? reply;
const rowsOf = (status) => (status.rows ?? []).map((row) => `${row.module}:${row.status}`).sort();
const sameText = (a, b) => (a ?? "").replace(/\r\n/g, "\n").trimEnd() === (b ?? "").replace(/\r\n/g, "\n").trimEnd();
const paneShown = async () => (await api.ui()).scm ?? null;

// A press on a control the pane has disabled answers did:false, and Refresh is disabled for as
// long as a read is in flight - which it is the moment the pane comes forward, and again after
// every stamp the host taps it with. So a press waits for the pane to be idle first: the check
// is about the control, not about whether the read it was racing had finished.
const paneIdle = () => waitFor("the pane to be idle", async () => {
  const pane = await paneShown();
  return pane && pane.busy === false ? pane : false;
}, { budgetMs: 30000 });

// The fixture's own three modules, checked by name because the checks below name them.
const inside = await api.project(project.project);
const names = (inside.components ?? []).map((one) => one.name);
for (const wanted of ["Ledger", "Reports", "Account"]) {
  if (!names.includes(wanted)) {
    console.log(`FAIL this workbook has no module named ${wanted}, so the checks cannot run.`
      + "\n     Rebuild the fixture: powershell tools\\New-ScmFixture.ps1");
    console.log("\n0 passed, 1 failed");
    process.exit(1);
  }
}

// What every check below is measured against, and what the cleanup puts back.
const ledgerWas = await held("Ledger");
const reportsWas = await held("Reports");
const accountWas = await held("Account");

const identity = { name: `scm.mjs ${process.pid}`, email: `scm-${process.pid}@example.invalid` };

try {
  // ---- no repository, then one --------------------------------------------------------------

  // Whatever a previous run left behind, forgotten first: a Repository pointing at a folder that
  // no longer exists would answer noRepository where this expects noFolder.
  await scm({ action: "forget" });
  const bare = await scm();
  check("a project under no source control says so", bare.state, "noFolder");
  check("and the words the pane uses ride the reply", typeof bare.covers === "string" && bare.covers.length > 0);

  const pointed = await scm({ action: "settings", folder });
  check("remembering a folder answers the status for it", canonical(pointed.folder), canonical(folder));
  check("with no repository above it, the state says so", pointed.state, "noRepository");
  check("and git was found", typeof pointed.gitVersion === "string" && /git/i.test(pointed.gitVersion));

  const initialised = await scm({ action: "init" });
  check("init makes the folder the repository root", canonical(initialised.repository), canonical(folder));
  check("and sets core.autocrlf off, so a module round-trips byte for byte",
    git("config", "--get", "core.autocrlf"), "false");
  check("and writes a .gitignore for the export's lock and partial files",
    existsSync(join(folder, ".gitignore")) && readFileSync(join(folder, ".gitignore"), "utf8").includes(".xlide-sync.lock"));
  check("and the state is now about identity or ready",
    ["noIdentity", "ready"].includes(initialised.state), true);

  // Written whether or not the machine has a global identity, so every commit below carries a
  // name the blame check can hold it to.
  const named = await scm({ action: "identity", name: identity.name, email: identity.email });
  check("identity writes the repository's user.name and user.email", named.identity, identity);
  check("and the repository is ready", named.state, "ready");
  const trunk = named.branch;
  check("on a branch with a name", typeof trunk === "string" && trunk.length > 0);

  // ---- every module is an added row, and one commit clears them ---------------------------------

  const fresh = await scm();
  check("with nothing committed, every module is an added row",
    (fresh.rows ?? []).every((row) => row.status === "added") && fresh.rows.length === inside.components.length, true);
  check("and the fixture's three are among them",
    ["Ledger", "Reports", "Account"].every((name) => fresh.rows.some((row) => row.module === name)), true);
  check("a class is a .cls and a standard module a .bas",
    [fresh.rows.find((row) => row.module === "Account")?.file, fresh.rows.find((row) => row.module === "Ledger")?.file],
    ["Account.cls", "Ledger.bas"]);
  check("and nothing is committed yet", fresh.lastCommit, null);

  const first = await scm({ action: "commit", message: "the fixture as built", by: "scm.mjs" });
  check("a commit answers its hash", /^[0-9a-f]{40}$/.test(first.hash ?? ""), true);
  check("and its short form", first.hash.startsWith(first.short) && first.short.length >= 7, true);
  check("committing everything commits the three",
    ["Ledger", "Reports", "Account"].every((name) => (first.committed ?? []).includes(name)), true);
  check("and skips nothing", first.skipped ?? [], []);
  check("and the rows are clear", rowsOf(statusOf(first)), []);
  check("and the last commit is this one", statusOf(first).lastCommit?.short, first.short);
  check("git agrees there is one commit", git("rev-list", "--count", "HEAD"), "1");
  check("and the files are on disk",
    ["Ledger.bas", "Reports.bas", "Account.cls"].every((file) => existsSync(join(folder, file))), true);

  // ---- a write shows as one modified row, and the suggestion names the hand ---------------------

  await write("Ledger", `${ledgerWas}\r\n' by claude ${process.pid}`, "claude");
  const oneWrite = await scm();
  check("a write through the door is one modified row", rowsOf(oneWrite), ["Ledger:modified"]);
  check("and the suggested message names the hand that wrote it",
    /claude/i.test(oneWrite.suggestedMessage ?? "") && /Ledger/.test(oneWrite.suggestedMessage ?? ""), true);

  await write("Reports", `${reportsWas}\r\n' by developer ${process.pid}`, "developer");
  const twoWrites = await scm();
  check("a second module written is a second row", rowsOf(twoWrites), ["Ledger:modified", "Reports:modified"]);

  // ---- the pane draws what the route answers, and commits one row through its own controls ----

  await api.ask(
    `(() => { const tab = document.querySelector('.panel-tab[data-panel="scm"]');`
    + ` if (tab) { tab.click(); } return !!tab; })()`);
  await api.act("scmPane", { file: project.project });

  // Pressed BEFORE the wait: a layout that persisted with this pane showing loaded it at page
  // ready, under no repository, and nothing above pushed a repaint through the pane's door.
  await paneIdle();
  const refreshed = await api.act("scmPane", { press: "refresh" });
  check("the pane's refresh presses", refreshed.did, true);

  const drawn = await waitFor("the pane to read the status", async () => {
    const pane = await paneShown();
    return pane && pane.project === project.project && pane.state === "ready" && pane.busy === false
      && (pane.rows ?? []).length === 2
      ? pane
      : false;
  }, { budgetMs: 30000 });

  check("the pane and the route agree about the rows", rowsOf(drawn), rowsOf(twoWrites));
  check("and about the branch", drawn.branch, trunk);

  const tickedOn = await api.act("scmPane", { tick: "Ledger", on: true });
  const tickedOff = await api.act("scmPane", { tick: "Reports", on: false });
  check("rows tick and untick from the pane", [tickedOn.did, tickedOff.did], [true, true]);

  const typed = await api.act("scmPane", { message: `just Ledger ${process.pid}` });
  check("the message types", typed.did, true);
  check("and reads back", (await paneShown())?.message, `just Ledger ${process.pid}`);

  await paneIdle();
  const pressedCommit = await api.act("scmPane", { press: "commit" });
  check("the pane's commit presses", pressedCommit.did, true);

  const afterPaneCommit = await waitFor("the pane's commit to land", async () => {
    const now = await scm();
    return JSON.stringify(rowsOf(now)) === JSON.stringify(["Reports:modified"]) ? now : false;
  }, { budgetMs: 40000 });
  check("a commit of the ticked row leaves the unticked one", rowsOf(afterPaneCommit), ["Reports:modified"]);
  check("git agrees there are two commits", git("rev-list", "--count", "HEAD"), "2");
  check("and the second one carries the message the pane typed",
    git("log", "-1", "--format=%s"), `just Ledger ${process.pid}`);

  // ---- history, comparisons, and the text at a commit -----------------------------------------

  const history = await scm({ action: "log" });
  check("the log shows two commits, newest first",
    (history.commits ?? []).map((one) => one.subject), [`just Ledger ${process.pid}`, "the fixture as built"]);
  check("with the files each touched",
    [history.commits[0]?.files.map((file) => file.file), history.commits[1]?.files.map((file) => file.file).sort()],
    [["Ledger.bas"], fresh.rows.map((row) => row.file).sort()]);
  check("and who made them", history.commits.every((one) => one.author === identity.name), true);
  check("and the first is the one the route answered", history.commits[1]?.hash, first.hash);
  const second = history.commits[0];

  check("a log limited to one commit answers one", (await scm({ action: "log", limit: 1 })).commits.length, 1);
  check("and a log for one module answers the commits that touched it",
    (await scm({ action: "log", module: "Reports" })).commits.map((one) => one.short), [first.short]);

  const lined = await scm({ action: "diff", module: "Reports" });
  check("a diff lines the live text against HEAD, and only the new line is added",
    (lined.rows ?? []).filter((row) => row.kind === "added" || row.kind === "removed").map((row) => `${row.kind}:${row.right}`),
    [`added:' by developer ${process.pid}`]);

  const shown = await scm({ action: "show", module: "Ledger", ref: first.short });
  check("show answers the text a module held at a commit",
    typeof shown.text === "string" && shown.text.includes('"Ledger.Balance"') && !shown.text.includes("by claude"), true);
  check("and lines it up against the live text",
    (shown.rows ?? []).some((row) => row.kind === "added" && /by claude/.test(row.right ?? "")), true);

  // ---- a branch: refused over a dirty workbook, then checked out and imported -----------------
  //
  // The branch is cut from the FIRST commit, so checking it out moves Ledger.bas back to the
  // fixture's text - which is what makes the import after it observable. Reports carries its
  // uncommitted line across, the way git carries a change that does not collide.

  git("branch", `feature-${process.pid}`, first.hash);

  await write("Reports", `${reportsWas}\r\n' by developer ${process.pid}\r\n' and again`, "developer");
  check("a write leaves the workbook dirty", (await scm()).dirty, true);

  const refused = await scm({ action: "checkout", ref: `feature-${process.pid}` })
    .then(() => "(answered)")
    .catch((error) => error.message);
  check("a checkout over a dirty workbook is refused in words", /dirty|unsaved|save/i.test(refused), true);
  check("and the branch has not moved", (await scm()).branch, trunk);

  await api.command("save");
  await waitFor("the save to land", async () => (await scm()).dirty === false, { budgetMs: 30000 });

  const switched = await scm({ action: "checkout", ref: `feature-${process.pid}` });
  check("after a save the checkout goes through", switched.branch, `feature-${process.pid}`);
  check("and git agrees", git("rev-parse", "--abbrev-ref", "HEAD"), `feature-${process.pid}`);
  check("and it imported the branch's Ledger into the project", sameText(await held("Ledger"), ledgerWas), true);
  check("while Reports kept the edit that did not collide", (await held("Reports")).includes("' and again"), true);
  check("and the branches list both, with the current one marked",
    [(switched.branches ?? []).some((one) => one.name === trunk && !one.current),
      (switched.branches ?? []).some((one) => one.name === `feature-${process.pid}` && one.current)],
    [true, true]);

  await paneIdle();
  await api.act("scmPane", { press: "refresh" });
  const paneBranch = await waitFor("the pane to follow the branch", async () => {
    const pane = await paneShown();
    return pane && pane.branch === `feature-${process.pid}` && pane.busy === false ? pane : false;
  }, { budgetMs: 20000 });
  check("and the pane's branch select shows it", paneBranch.branch, `feature-${process.pid}`);

  // ---- blame: committed lines carry their commit, the rest are uncommitted -------------------

  const ledgerNow = await held("Ledger");
  const ledgerLines = ledgerNow.split(/\r?\n/).length;
  await write("Ledger", `${ledgerNow}\r\n' blame probe ${process.pid}`, "claude");

  const blamed = await scm({ action: "blame", module: "Ledger" });
  check("blame maps every committed line to the editor's numbering",
    (blamed.lines ?? []).map((line) => line.line), Array.from({ length: ledgerLines }, (_, at) => at + 1));
  check("and each carries the commit that wrote it",
    blamed.lines.every((line) => line.hash === first.hash && line.short === first.short && line.author === identity.name), true);
  check("and the line written since is uncommitted", blamed.uncommitted, [ledgerLines + 1]);

  await api.caret(1, { module: "Ledger", project: project.projectId });
  await waitFor("Ledger to be the active module", async () =>
    ((await api.ui()).focus.model ?? "").toLowerCase().endsWith("/ledger"), { budgetMs: 15000 });

  // `ui.blame` is three counts for the active model - on, hinted lines, uncommitted lines - so
  // the layer is held to the route's numbers rather than to its rows.
  const toggled = await api.act("blame", { which: "toggle" });
  check("the editor's blame toggle presses", toggled.did, true);
  const painted = await waitFor("the blame layer to paint", async () => {
    const layer = (await api.ui()).blame;
    return layer && layer.on === true && layer.lines > 0 ? layer : false;
  }, { budgetMs: 20000 });
  check("and the layer is on, painting the committed lines", painted.lines, ledgerLines);
  check("with the new line marked uncommitted", painted.uncommitted, 1);

  const off = await api.act("blame", { which: "off" });
  check("and off again", off.did && (await api.ui()).blame?.on === false, true);

  // ---- restore: the text at a commit, landing as a change-log round -------------------------

  const restored = await scm({ action: "restore", module: "Ledger", ref: first.hash, by: "claude" });
  check("restoring a module from a commit writes it", restored.did, "written");
  check("and the module holds the commit's text", sameText(await held("Ledger"), ledgerWas), true);

  const rounds = (await log()).rounds ?? [];
  const restoreRound = rounds.find((round) => {
    const label = round.label ?? "";
    return label.startsWith("restore Ledger from ") && first.hash.startsWith(label.slice("restore Ledger from ".length));
  });
  check("and it landed in the change log as a round labelled with the commit",
    restoreRound ? `${restoreRound.by}` : "(no such round)", "claude");

  const again = await scm({ action: "restore", module: "Ledger", ref: first.hash, by: "claude" });
  check("restoring what is already there says so", again.did, "unchanged");

  // ---- the past-version tab --------------------------------------------------------------------

  const opened = await scm({ action: "open", module: "Ledger", ref: first.short });
  check("open answers the status", opened.state, "ready");

  const tabOf = (ui) => (ui.workspace?.groups ?? []).flatMap((group) => group.tabs ?? [])
    .find((tab) => tab.label === `Ledger @ ${first.short}`);
  const historyTab = await waitFor("the past-version tab", async () => {
    const tab = tabOf(await api.ui());
    return tab && tab.active ? tab : false;
  }, { budgetMs: 20000 });
  check("open puts a tab labelled with the module and the short hash on the strip, active",
    { label: historyTab.label, active: historyTab.active }, { label: `Ledger @ ${first.short}`, active: true });

  await api.act("closeActive");
  await waitFor("the past-version tab to close", async () => tabOf(await api.ui()) === undefined, { budgetMs: 15000 });
  check("and closing the active tab removes it", tabOf(await api.ui()) === undefined, true);

  // ---- the folder: an edit made outside shows up on its own, and imports when asked -----------

  const accountFile = join(folder, "Account.cls");
  const accountOnDisk = readFileSync(accountFile, "utf8");
  writeFileSync(accountFile, `${accountOnDisk.replace(/\r?\n$/, "")}\r\n' edited in the folder ${process.pid}\r\n`, "utf8");

  const noticed = await waitFor("the pane to notice the folder edit on its own", async () => {
    const pane = await paneShown();
    const row = (pane?.outside ?? []).find((one) => one.module === "Account");
    return row ? row : false;
  }, { budgetMs: 30000 });
  check("an edit made in the folder appears in the Folder section, nothing pressed", noticed.status, "folderNewer");
  check("and the route agrees",
    ((await scm()).outside ?? []).some((one) => one.module === "Account" && one.status === "folderNewer"), true);
  check("while the module itself is untouched", (await held("Account")).includes("edited in the folder"), false);

  const imported = await scm({ action: "import", modules: ["Account"], by: "scm.mjs" });
  check("import writes the folder's text into the module", imported.imported, ["Account"]);
  check("and the module holds it", (await held("Account")).includes(`' edited in the folder ${process.pid}`), true);
  check("and the Folder section no longer names it",
    (statusOf(imported).outside ?? []).some((one) => one.module === "Account"), false);

  // ---- the other presses, and the remote path answering in words --------------------------

  await paneIdle();
  const exported = await api.act("scmPane", { press: "export" });
  check("the pane's export presses", exported.did, true);

  await paneIdle();
  const rowOpened = await api.act("scmPane", { module: "Reports" });
  check("a row opens its comparison from the pane", rowOpened.did, true);
  const diffRows = await waitFor("the comparison to draw", async () => {
    const said = await api.ask(`[...document.querySelectorAll('#scm-diff .sync-diff-row')].length`);
    const count = typeof said === "number" ? said : Number(said);
    return count > 0 ? count : false;
  }, { budgetMs: 15000 });
  check("and it draws rows", diffRows > 0, true);

  await paneIdle();
  const commitOpened = await api.act("scmPane", { commit: first.short });
  check("a commit opens from the pane's history", commitOpened.did, true);

  // ---- a remote: a bare repository beside the folder, then one that cannot answer -------------

  execFileSync(gitExe, ["init", "--bare", remote], { encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] });
  check("before a remote is attached the status says so", (await scm()).remoteUrl, "");
  const attached = await scm({ action: "remote", url: remote });
  check("a remote attaches through the route, as origin", [attached.remote, attached.remoteUrl], ["origin", remote]);
  check("and git holds it", git("remote", "get-url", "origin"), remote);
  await api.act("scmPane", { press: "refresh" });
  await paneIdle();
  check("and the pane shows it", (await paneShown()).remoteUrl, remote);
  const onBranch = (await scm()).branch;
  const pushed = await scm({ action: "push" });
  check("a first push sets the upstream itself", pushed.upstream, `origin/${onBranch}`);
  check("and leaves the branch level with it", [pushed.ahead, pushed.behind], [0, 0]);
  check("and git agrees the remote holds the branch",
    git("ls-remote", "--heads", "origin").includes(`refs/heads/${onBranch}`), true);
  const fetched = await scm({ action: "fetch" });
  check("fetch answers the status", fetched.state, "ready");
  const pulled = await scm({ action: "pull" });
  check("pull with nothing to pull says so in git's words", /up to date/i.test(pulled.detail ?? ""), true);

  // Then a remote that cannot answer, so the failure is git's own words at once rather than a
  // hang on a hidden prompt. With no remote at all the route refuses before git runs, which
  // proves nothing about the prompt; and git itself, asked to fetch with no remote, prints
  // nothing and exits 0.
  // Re-pointed THROUGH THE PANE: the remote line's input and button, the way a wrong paste is
  // corrected, which is the same verb the route ran above with `set-url` behind it this time.
  const nowhere = join(folder, "nowhere");
  await paneIdle();
  const urlTyped = await api.act("scmPane", { url: nowhere });
  check("the pane's remote line takes a URL", urlTyped.did, true);
  const urlPressed = await api.act("scmPane", { press: "remote" });
  check("and its button re-points the remote", urlPressed.did, true);
  await paneIdle();
  check("so git holds the new URL", git("remote", "get-url", "origin"), nowhere);
  check("and the status says so", (await scm()).remoteUrl, nowhere);
  const broken = await scm({ action: "fetch" }).catch((error) => ({ detail: error.message }));
  check("fetch against a remote that cannot answer fails in git's own words, without hanging",
    /does not appear to be a git repository|could not read from remote/i.test(broken.detail ?? ""), true);

  // ---- and the way out ------------------------------------------------------------------------

  // RACED ON PURPOSE: a status round is out with git for a few hundred milliseconds, and a forget
  // that lands inside that window used to be undone by the round's last step, which armed the
  // watcher for the folder it had gathered. The forgotten folder then refused to delete and
  // stamped the page until the workbook closed. The door answers concurrently, so the forget can
  // overtake the round; a reply that already says noFolder means it did not, and the race is run
  // again from a fresh pointing.
  let overtaken = null;
  let forgotten = null;
  let sinceRace = 0;
  for (let attempt = 0; attempt < 5 && overtaken === null; attempt += 1) {
    if (attempt > 0) {
      await scm({ action: "settings", folder });
      await paneIdle();
    }

    sinceRace = (await api.log({ max: 1 })).next;
    const inFlight = scm();
    await sleep(30);
    forgotten = await scm({ action: "forget" });
    const raced = await inFlight;
    overtaken = raced.state === "noFolder" ? null : raced.state;
  }

  check("forgetting the folder puts the project back under no source control", forgotten.state, "noFolder");
  check("the forget overtook a status round that had gathered the folder", overtaken, "ready");

  // THE LOG IS THE WITNESS, because the fault is invisible from outside: a watcher holds its
  // folder with delete sharing, so the folder still deletes, and the stamps it sends land on a
  // pane that answers "no folder". The host logs "watching" only when it arms a new watcher,
  // and the watcher for this folder was standing before the race, so any such line after it is
  // the round the forget overtook arming one for a folder the project no longer names.
  await paneIdle();
  const armed = (await api.log({ since: sinceRace, match: "scm: watching", max: 20 })).lines ?? [];
  check("the round the forget overtook does not arm a watcher for the forgotten folder", armed.length, 0);
} finally {
  // THE FIXTURE GOES BACK, whatever happened above: the texts as captured, saved to the file,
  // the repository forgotten, and the folder gone. Each step is tolerant of the others, because
  // half a cleanup is better than none and a cleanup that dies hides the finding it follows.
  for (const [name, text] of [["Ledger", ledgerWas], ["Reports", reportsWas], ["Account", accountWas]]) {
    try {
      await write(name, text, "scm.mjs cleanup");
    } catch (error) {
      console.log(`     WARNING: ${name} was left holding this run's text (${error.message})`);
    }
  }

  try {
    await api.act("blame", { which: "off" });
  } catch {
    // The layer may never have been built; nothing to turn off.
  }

  try {
    // WAITED FOR ON DISK, the fixture driver's rule: the api's save answers before Excel has
    // written, and the gate's next group kills Excel, so a save only asked for is a save lost.
    // `dirty` is answered in every state, so the folder's absence does not cut the wait short.
    const wasWritten = statSync(project.projectId).mtimeMs;
    await api.command("save");
    await waitFor("the cleanup save to land", async () =>
      (await scm()).dirty === false && statSync(project.projectId).mtimeMs > wasWritten,
    { budgetMs: 30000 });
  } catch (error) {
    console.log(`     WARNING: the fixture may be left unsaved (${error.message})`);
  }

  try {
    await scm({ action: "forget" });
  } catch (error) {
    console.log(`     WARNING: the folder is still remembered as the repository (${error.message})`);
  }

  // GUARDED like every step above it: a throw here would take the verdict with it.
  for (const gone of [folder, remote]) {
    const trouble = await removeWhenFree(gone);
    if (trouble !== null) {
      console.log(`     WARNING: ${gone} was left behind (${trouble})`);
    }
  }
}

process.exit(done());
