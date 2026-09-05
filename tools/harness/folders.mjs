/*
 * THE FOLDER LAYOUT AND THE CURRENT PROCEDURE, against real Excel (#23).
 *
 * The headless probe (folders-page-probe.mjs) proves what the tree does with folders it is told
 * about. This proves the telling: that the host reads '@Folder("Parent.Child") out of every
 * module, opened or not, in every spelling the fixture writes it; that an annotation typed into a
 * module moves it on the next analysis pass, and back; that two workbooks with the same module
 * names and the same folder names keep their folders apart; that the layout is a setting the
 * host keeps; and that the status bar's "current procedure" agrees with the editor's own
 * ProcOfLine on EVERY line of a module built to disagree with a lazy rule.
 *
 * Runs against FolderFixture.xlsm and FolderTwinFixture.xlsm, open together:
 *
 *   tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\FolderFixture.xlsm,artifacts\fixtures\FolderTwinFixture.xlsm
 *   node tools\harness\folders.mjs
 *
 * Puts back everything it changes: the layout setting, the folded folders, and Loose's text.
 */
import { open, reporter, waitFor } from "./xlide-api.mjs";

const api = await open();
const { check, done } = reporter();

const MAIN = "FolderFixture.xlsm";
const TWIN = "FolderTwinFixture.xlsm";

/** The folder each module's annotation names, as the fixture wrote them. */
const MAIN_FOLDERS = {
  Ledger: "Accounts.Ledger",
  Posting: "Accounts.Ledger",
  Invoice: "Accounts.Billing",
  Helpers: "Shared",
  Tools: "shared",
  Bare: "Accounts.Billing.Reminders",
  Loose: null,
  Late: null,
  Twice: "Accounts",
  Procedures: "Shared",
  Sheet1: "Accounts",
  ReminderForm: "Accounts.Billing.Reminders",
  ThisWorkbook: null,
};

const TWIN_FOLDERS = {
  Helpers: "Accounts.Ledger",
  Ledger: null,
  TwinOnly: "Shared",
  Posting: "Twin.Only",
  Sheet1: null,
  ThisWorkbook: null,
};

const book = (ui, name) => ui.explorer.workbooks.find((one) => one.name === name);
/** "A.B.C" -> ["a", "a.b", "a.b.c"]: the folders a module in that path is inside, lower-cased. */
const ancestorsOf = (path) => {
  const segments = (path ?? "").split(".").filter((one) => one.length > 0);
  return segments.map((_, depth) => segments.slice(0, depth + 1).join(".").toLowerCase());
};
/** Whether a workbook's folders are exactly those on the way to the module: open on the path, folded elsewhere. */
const onlyThePath = (ui, workbook, module) => {
  const inside = ancestorsOf(book(ui, workbook)?.modules.find((one) => one.name === module)?.folder);
  return (book(ui, workbook)?.folders ?? []).every((one) => one.expanded === inside.includes(one.path.toLowerCase()));
};
const foldersOf = (ui, name) => book(ui, name)?.folders.map((one) => one.path) ?? [];
const folderMap = (rows) => Object.fromEntries(rows.map((row) => [row.name, row.folder ?? null]));
const differences = (expected, actual) =>
  Object.entries(expected).filter(([name, folder]) => actual[name] !== folder)
    .map(([name, folder]) => `${name}: ${JSON.stringify(actual[name])} not ${JSON.stringify(folder)}`);

/** The tree's rows between one workbook's row and the next, tagged by what each is. */
const rowsOf = (workbook) => api.ask(`(() => {
  const all = [...document.querySelectorAll('#sidebar-tree > *')];
  const start = all.findIndex((row) => row.dataset.project === ${JSON.stringify(workbook)});
  if (start < 0) return null;
  const out = [];
  for (const row of all.slice(start + 1)) {
    if (row.dataset.project !== undefined) break;
    if (row.dataset.folder !== undefined) out.push('folder:' + row.dataset.folder);
    else if (row.dataset.component) out.push('module:' + row.dataset.component);
  }
  return out;
})()`);

/** A wait that answers false on timing out, for a wait that feeds a check rather than ends the run. */
const settles = (what, predicate, options) => waitFor(what, predicate, options).then(() => true, () => false);

const originalSettings = await api.settings();
let looseText = null;

// Whether this session has already run this suite: the fresh-launch checks below are about the
// order things arrive in at a session's start, which a rerun cannot reproduce, so a rerun asks
// the same question by driving the state instead. Never skipped, never vacuous.
const rerun = ((await api.log({ match: "folders.mjs: ran", max: 5 })).lines ?? []).length > 0;

try {
  // ---- what the host read ----
  const main = await api.project(MAIN);
  const twin = await api.project(TWIN);
  const mainMissed = differences(MAIN_FOLDERS, folderMap(main.components));
  check("the project route reports every annotation in the main workbook, in every spelling",
    mainMissed.length === 0 && main.components.length === Object.keys(MAIN_FOLDERS).length,
    mainMissed.join("; ") || `${main.components.length} components`);
  const twinMissed = differences(TWIN_FOLDERS, folderMap(twin.components));
  check("and every annotation in the twin, where the same names sit elsewhere",
    twinMissed.length === 0, twinMissed.join("; "));

  // ---- what the tree holds, in either layout ----
  let ui = await api.ui();
  check("the tree carries each module's folder whichever layout is showing",
    differences(MAIN_FOLDERS, folderMap(book(ui, MAIN)?.modules ?? [])).length === 0
      && differences(TWIN_FOLDERS, folderMap(book(ui, TWIN)?.modules ?? [])).length === 0,
    JSON.stringify(folderMap(book(ui, MAIN)?.modules ?? [])));
  check("the main workbook's folders are the annotations' paths, parents first, two spellings as one",
    foldersOf(ui, MAIN).join(",") === "Accounts,Accounts.Billing,Accounts.Billing.Reminders,Accounts.Ledger,Shared",
    foldersOf(ui, MAIN).join(","));
  check("the twin's folders are its own, a Shared of its own included",
    foldersOf(ui, TWIN).join(",") === "Accounts,Accounts.Ledger,Shared,Twin,Twin.Only",
    foldersOf(ui, TWIN).join(","));

  // ---- the layout is a setting the host keeps ----
  // The developer's own choice is recorded and put back, not asserted: the page shows whatever
  // the host holds, and the suite starts from the tree layout by asking for it.
  check("the page shows the layout the host holds", ui.explorer.view === originalSettings.explorerView,
    `${originalSettings.explorerView} / ${ui.explorer.view}`);
  if (originalSettings.explorerView !== "tree") {
    await api.settings({ explorerView: "tree" });
    await settles("the tree layout", async () => (await api.ui()).explorer.view === "tree");
  }
  const switched = await api.act("explorerView", { view: "folders" });
  check("explorerView asks the host for the folder layout", switched.did, switched.detail);
  const landed = await settles("the folder layout", async () => (await api.ui()).explorer.view === "folders");
  check("and the layout lands on the host's echo", landed);
  check("the host keeps it as a setting", (await api.settings()).explorerView === "folders");

  // ---- the folders follow the editor like the workbooks do ----
  // In the workbook being edited, the folders on the way to the shown module are open and the
  // others folded; the other workbook, where the attention never was, stands as it started. On a
  // FRESH launch this is the state the session started in - the active module is announced
  // before the tree arrives, and the folders have to catch up when it does - and which
  // workbook the host shows first is the host's choice, so it is read rather than assumed. On a
  // rerun the same rule is driven in the main workbook: out to the root, then back to Ledger.
  if (rerun) {
    await api.caret(1, { module: "Loose", project: MAIN });
    await settles("Loose active", async () => (await api.ui()).explorer.active === "Loose");
    await api.caret(1, { module: "Ledger", project: MAIN });
    await settles("Ledger active", async () => (await api.ui()).explorer.active === "Ledger");
    ui = await api.ui();
    check("the folders on the way to the module being edited are open and the others folded (driven)",
      onlyThePath(ui, MAIN, "Ledger")
        && book(ui, MAIN).folders.find((one) => one.path === "Accounts.Ledger")?.expanded === true
        && book(ui, MAIN).folders.find((one) => one.path === "Shared")?.expanded === false,
      JSON.stringify(book(ui, MAIN).folders));
  } else {
    ui = await api.ui();
    const shownIn = ui.explorer.attentionWorkbook;
    const other = shownIn === MAIN ? TWIN : MAIN;
    check("the folders on the way to the module being edited are open and the others folded (fresh launch)",
      ui.explorer.active !== null && (shownIn === MAIN || shownIn === TWIN)
        && onlyThePath(ui, shownIn, ui.explorer.active)
        && book(ui, other).folders.every((one) => one.expanded),
      `${ui.explorer.active} in ${shownIn}: ${JSON.stringify(book(ui, shownIn ?? MAIN)?.folders)}; other: ${JSON.stringify(book(ui, other)?.folders)}`);
  }

  // ---- the shape, per workbook ----
  await api.act("expandWorkbook", { workbook: MAIN, open: true });
  await api.act("expandWorkbook", { workbook: TWIN, open: true });
  for (const workbook of [MAIN, TWIN]) {
    for (const path of foldersOf(ui, workbook)) {
      await api.act("expandFolder", { workbook, path, open: true });
    }
  }
  const mainRows = await rowsOf(MAIN);
  check("the main workbook draws folders before modules, nested, with the root modules last",
    (mainRows ?? []).join(" ") === [
      "folder:Accounts", "folder:Accounts.Billing", "folder:Accounts.Billing.Reminders",
      "module:ReminderForm", "module:Bare", "module:Invoice",
      "folder:Accounts.Ledger", "module:Ledger", "module:Posting",
      "module:Sheet1", "module:Twice",
      "folder:Shared", "module:Helpers", "module:Procedures", "module:Tools",
      "module:ThisWorkbook", "module:Late", "module:Loose",
    ].join(" "),
    (mainRows ?? []).join(" "));
  const twinRows = await rowsOf(TWIN);
  check("the twin draws its own folders, with the colliding names where ITS annotations put them",
    (twinRows ?? []).join(" ") === [
      "folder:Accounts", "folder:Accounts.Ledger", "module:Helpers",
      "folder:Shared", "module:TwinOnly",
      "folder:Twin", "folder:Twin.Only", "module:Posting",
      "module:Sheet1", "module:ThisWorkbook", "module:Ledger",
    ].join(" "),
    (twinRows ?? []).join(" "));

  // ---- folding one workbook's folder leaves the other's alone ----
  const folded = await api.act("expandFolder", { workbook: MAIN, path: "Shared", open: false });
  ui = await api.ui();
  check("folding Shared in the main workbook folds only that one",
    folded.did
      && book(ui, MAIN).folders.find((one) => one.path === "Shared")?.expanded === false
      && book(ui, TWIN).folders.find((one) => one.path === "Shared")?.expanded === true,
    JSON.stringify([book(ui, MAIN).folders, book(ui, TWIN).folders]));
  check("and its modules leave the tree while the twin's Shared keeps its module",
    !(await rowsOf(MAIN)).includes("module:Helpers") && (await rowsOf(TWIN)).includes("module:TwinOnly"));
  await api.act("expandFolder", { workbook: MAIN, path: "Shared", open: true });

  // ---- the follow opens the folders above the module being edited ----
  // From a module at the root first: the tree follows a CHANGE of module, and a folder folded
  // around the module already being edited stays folded, which is the developer's own hand.
  await api.caret(1, { module: "Loose", project: MAIN });
  await settles("Loose active", async () => (await api.ui()).explorer.active === "Loose");
  check("moving to a module at the root folds every folder of that workbook",
    book(await api.ui(), MAIN).folders.every((one) => !one.expanded),
    JSON.stringify(book(await api.ui(), MAIN).folders));
  await api.caret(1, { module: "Bare", project: MAIN });
  const followed = await settles("the follow", async () => {
    const now = await api.ui();
    return now.explorer.active === "Bare"
      && ["Accounts", "Accounts.Billing", "Accounts.Billing.Reminders"]
        .every((path) => book(now, MAIN).folders.find((one) => one.path === path)?.expanded);
  });
  check("editing a module inside a folded folder opens the folders above it, and only those", followed
    && book(await api.ui(), MAIN).folders.find((one) => one.path === "Shared")?.expanded === false,
    JSON.stringify(book(await api.ui(), MAIN).folders));

  // ---- an annotation typed into a module moves it, and back ----
  looseText = (await api.readModule("Loose", MAIN)).text;
  check("Loose starts with no annotation", looseText !== undefined && !/@Folder/i.test(looseText));
  await api.writeModule("Loose", `'@Folder("Moved.Here")\r\n${looseText}`, MAIN);
  const moved = await settles("Loose to move", async () => {
    const now = await api.ui();
    return book(now, MAIN)?.modules.find((one) => one.name === "Loose")?.folder === "Moved.Here"
      && foldersOf(now, MAIN).includes("Moved.Here");
  }, { budgetMs: 30000 });
  check("writing an annotation into a module moves it into a new folder on the next pass", moved,
    JSON.stringify(book(await api.ui(), MAIN)?.modules.find((one) => one.name === "Loose")));
  check("the project route agrees",
    (await api.project(MAIN)).components.find((one) => one.name === "Loose")?.folder === "Moved.Here");
  check("and the new folder is drawn", (await rowsOf(MAIN)).includes("folder:Moved.Here"));
  await api.writeModule("Loose", looseText, MAIN);
  const back = await settles("Loose to come back", async () => {
    const now = await api.ui();
    return book(now, MAIN)?.modules.find((one) => one.name === "Loose")?.folder === null
      && !foldersOf(now, MAIN).includes("Moved.Here");
  }, { budgetMs: 30000 });
  check("taking the annotation out puts it back at the root, and the empty folder goes", back);
  looseText = null;

  // ---- the current procedure, held to the editor's own answer on every line ----
  const procedures = (await api.readModule("Procedures", MAIN)).text ?? "";
  const lineCount = procedures.split(/\r\n|\r|\n/).length;
  check("the parity module has the shapes it exists for", lineCount > 20 && /Property Let Total/.test(procedures), `${lineCount} lines`);

  const disagreements = [];
  let inProcedure = 0;
  let inDeclarations = 0;
  for (let line = 1; line <= lineCount; line++) {
    await api.caret(line, { module: "Procedures", project: MAIN });
    const settled = await settles(`line ${line}`, async () => {
      const [now, state] = await Promise.all([api.ui(), api.state()]);
      return now.statusPosition === `Ln ${line}, Col 1` && state.caretLine === line && state.shownModule === "Procedures";
    }, { budgetMs: 10000 });
    const [now, state] = await Promise.all([api.ui(), api.state()]);
    const pageName = now.statusProcedure === "(Declarations)" ? null : now.statusProcedure.split(" ").pop() ?? null;
    const hostName = state.procedureAtCaret;
    if (!settled || pageName !== hostName) {
      disagreements.push(`line ${line}: page "${now.statusProcedure}" host ${JSON.stringify(hostName)}${settled ? "" : " (unsettled)"}`);
    }
    if (hostName === null) { inDeclarations += 1; } else { inProcedure += 1; }

    // The tree's mark, on the row of the same name, while the module is the unfolded one.
    const marked = now.explorer.currentProcedure;
    if (hostName !== null && (marked?.name !== hostName || marked?.module !== "Procedures")) {
      disagreements.push(`line ${line}: the tree marks ${JSON.stringify(marked)} not ${hostName}`);
    }
    if (hostName === null && marked !== null) {
      disagreements.push(`line ${line}: the tree marks ${JSON.stringify(marked)} in the declarations`);
    }
  }
  check("the status bar's current procedure agrees with ProcOfLine on every line, and the tree marks the same row",
    disagreements.length === 0, disagreements.slice(0, 6).join("; ") + (disagreements.length > 6 ? ` +${disagreements.length - 6} more` : ""));
  // The instrument cannot pass vacuously: the module has both kinds of line, and both were seen.
  check("both kinds of line were measured", inProcedure >= 12 && inDeclarations >= 8, `${inProcedure} in procedures, ${inDeclarations} in the declarations`);

  // ---- back to the flat tree ----
  await api.act("explorerView", { view: "tree" });
  const flat = await settles("the tree layout", async () => (await api.ui()).explorer.view === "tree");
  check("the tree layout comes back on request", flat);
  check("and draws no folder rows", !(await rowsOf(MAIN)).some((row) => row.startsWith("folder:")));
} finally {
  // Put back, and CHECKED: a restore posted while the host thread was busy answered an error,
  // was swallowed, and the next run started in the folder layout and failed its first check.
  if (looseText !== null) {
    await api.writeModule("Loose", looseText, MAIN).catch(() => {});
  }
  await api.mark("folders.mjs: ran").catch(() => {});
  await settles("the layout put back", async () => {
    await api.settings({ explorerView: originalSettings.explorerView });
    return (await api.settings()).explorerView === originalSettings.explorerView;
  });
  for (const path of ["Accounts", "Shared"]) {
    await api.act("expandFolder", { workbook: MAIN, path, open: true }).catch(() => {});
  }
}

done();
