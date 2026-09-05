// Drives the BUILT editor page (ui/editor/dist) in a headless browser and walks the explorer's
// FOLDER LAYOUT against the page's demo transport (#23).
//
// The demo's Book1 carries two annotated modules - Module1 in "Sales.Import", SalesRow in
// "sales" - and two at the root, which is enough to ask everything the layout promises: the two
// tabs, the folders drawn before the modules, one folder for two spellings, one level of indent
// per folder, a folder that folds and says what it hides, the api that opens it again, the
// follow that unfolds a folded folder when the module being edited is inside it, the flat tree
// back on the other tab, and the caret's mark on the procedure row beneath the unfolded module.
//
// Prints a JSON verdict {pass, checks} on stdout and exits nonzero when any check fails.
// Everything goes through `xlideUi.act` and `xlideUi.state`, so the api a script would use is
// what is under test here alongside the feature. NO BACKTICKS in the drive: it is spliced into
// a template literal.

import { runPageProbe } from "./page-probe.mjs";

const DRIVE = `(async () => {
  const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
  const act = (name, args) => window.xlideUi.act(name, args ?? {});
  const state = () => window.xlideUi.state();
  const tabs = () => [...document.querySelectorAll('.tab')].map((tab) => tab.dataset.module);
  const rows = () => [...document.querySelectorAll('#sidebar-tree > *')].map((row) =>
    row.dataset.folder !== undefined ? 'folder:' + row.dataset.folder
      : row.dataset.component ? 'module:' + row.dataset.component
      : row.dataset.procModule ? 'proc:' + row.textContent.trim()
      : row.dataset.project ? 'workbook:' + row.dataset.project
      : 'other');
  const depthOf = (selector) => document.querySelector(selector)?.style.getPropertyValue('--tree-depth') || '0';
  const waitFor = async (what, until) => {
    for (let waited = 0; waited < 10000; waited += 100) {
      if (until()) return true;
      await sleep(100);
    }
    return false;
  };

  const checks = [];
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });

  // The demo publishes the tree a beat after load; nothing below means anything until it has.
  await waitFor('tabs', () => tabs().length >= 2);
  act('expandWorkbook', { workbook: 'Book1.xlsm', open: true });
  await waitFor('rows', () => rows().includes('module:Module1'));

  // ---- the tabs ----
  const tabsDrawn = [...document.querySelectorAll('#explorer-views .explorer-view')];
  check('two layout tabs sit above the tree, Tree first',
    tabsDrawn.map((tab) => tab.textContent).join(',') === 'Tree,Folders', tabsDrawn.map((tab) => tab.textContent).join(','));
  check('the tree layout is the one showing and its tab says so',
    state().explorer.view === 'tree' && tabsDrawn[0].getAttribute('aria-selected') === 'true',
    state().explorer.view + ' / ' + tabsDrawn[0].getAttribute('aria-selected'));
  check('the flat tree draws no folder rows', !rows().some((row) => row.startsWith('folder:')), rows().join(' '));
  check('the folders exist as data whichever layout is showing',
    state().explorer.workbooks[0].folders.map((one) => one.path).join(',') === 'Sales,Sales.Import',
    JSON.stringify(state().explorer.workbooks[0].folders));

  // ---- switching by the api, which is a settings change echoed back ----
  const asked = act('explorerView', { view: 'folders' });
  check('explorerView asks for the layout', asked.did, asked.detail);
  check('the folder layout lands on the echo', await waitFor('view', () => state().explorer.view === 'folders'), state().explorer.view);
  check('the Folders tab is now the selected one',
    tabsDrawn[1].getAttribute('aria-selected') === 'true' && tabsDrawn[0].getAttribute('aria-selected') === 'false');

  // ---- the shape ----
  const book1 = () => {
    const all = rows();
    const start = all.indexOf('workbook:Book1.xlsm');
    const end = all.indexOf('workbook:Book2.xlsm');
    return all.slice(start + 1, end < 0 ? undefined : end);
  };
  check('folders come first, parents before children, then the root modules by kind',
    book1().filter((row) => !row.startsWith('proc:')).join(' ')
      === 'folder:Sales folder:Sales.Import module:Module1 module:SalesRow module:Sheet1 module:ThisWorkbook',
    book1().join(' '));
  check('two spellings of one folder are one folder, spelled as the first module spelled it',
    document.querySelectorAll('[data-folder]').length === 2
      && document.querySelector('[data-folder="Sales"] .tree-folder-name').textContent === 'Sales');
  check('a module sits one level deeper per folder above it',
    depthOf('[data-component="Module1"]') === '2' && depthOf('[data-component="SalesRow"]') === '1'
      && depthOf('[data-component="Sheet1"]') === '0' && depthOf('[data-folder="Sales.Import"]') === '1',
    [depthOf('[data-component="Module1"]'), depthOf('[data-component="SalesRow"]'), depthOf('[data-component="Sheet1"]')].join(','));
  check('a folder row carries the annotation that makes it, as its tooltip',
    document.querySelector('[data-folder="Sales.Import"]').title === '\\'@Folder("Sales.Import")',
    document.querySelector('[data-folder="Sales.Import"]').title);
  check('the module rows keep the attributes the flat tree gives them',
    document.querySelector('[data-component="Module1"]').dataset.workbook === 'Book1.xlsm'
      && document.querySelector('[data-component="Module1"]').dataset.kind === '1');

  // ---- folding ----
  document.querySelector('[data-folder="Sales"]').click();
  await sleep(50);
  check('clicking a folder folds it and everything under it',
    !rows().includes('folder:Sales.Import') && !rows().includes('module:Module1') && !rows().includes('module:SalesRow'),
    rows().join(' '));
  check('a folded folder says how many modules it hides',
    document.querySelector('[data-folder="Sales"] .tree-kind')?.textContent === '2 modules'
      && document.querySelector('[data-folder="Sales"]').getAttribute('aria-expanded') === 'false',
    document.querySelector('[data-folder="Sales"]')?.textContent);
  check('the snapshot says which folders are shut',
    JSON.stringify(state().explorer.workbooks[0].folders) === JSON.stringify([
      { path: 'Sales', expanded: false, modules: 2 }, { path: 'Sales.Import', expanded: true, modules: 1 }]),
    JSON.stringify(state().explorer.workbooks[0].folders));

  const reopened = act('expandFolder', { workbook: 'Book1.xlsm', path: 'sales', open: true });
  check('expandFolder opens it again, by path without regard to case', reopened.did && rows().includes('module:Module1'), reopened.detail);
  const missing = act('expandFolder', { workbook: 'Book1.xlsm', path: 'Nowhere', open: true });
  check('and says so for a folder that does not exist', !missing.did, missing.detail);

  // ---- the follow ----
  // Module2 is at the root and open in a tab. Move there, fold Sales, then come back to Module1:
  // the tree following the editor opens the folders above the module it is following.
  act('activate', { module: 'Module2' });
  await waitFor('active', () => state().explorer.active === 'Module2');
  check('moving to a module outside the folder folds it, the way leaving a workbook folds the workbook',
    !rows().includes('module:Module1') && state().explorer.workbooks[0].folders.every((one) => !one.expanded),
    JSON.stringify(state().explorer.workbooks[0].folders));
  act('activate', { module: 'Module1' });
  await waitFor('active', () => state().explorer.active === 'Module1');
  check('following the editor back into it opens the folders above the module',
    rows().includes('module:Module1') && state().explorer.workbooks[0].folders.every((one) => one.expanded),
    rows().join(' '));
  check('and the accordion unfolds the module there, under its folder',
    state().explorer.unfolded?.module === 'Module1' && rows().indexOf('proc:Sub Recalculate') > rows().indexOf('module:Module1'),
    rows().join(' '));
  check('the procedure rows sit one level deeper than their module',
    depthOf('[data-proc-module="Module1"]') === '2', depthOf('[data-proc-module="Module1"]'));

  // ---- the caret's mark ----
  // Line 10 is Sub Recalculate's header; line 3 is Option Explicit's neighbour, in no procedure.
  act('select', { startLine: 12 });
  await sleep(50);
  const bar = () => document.getElementById('status-procedure').textContent;
  check('the status bar names the procedure the caret is in', bar() === 'Sub Recalculate', bar());
  check('and the snapshot reads the same', state().statusProcedure === 'Sub Recalculate', state().statusProcedure);
  const marked = () => [...document.querySelectorAll('.tree-proc.current')].map((row) => row.textContent.trim());
  check('the tree marks that procedure\\'s row, and only it',
    marked().join(',') === 'Sub Recalculate' && document.querySelector('.tree-proc.current').getAttribute('aria-current') === 'location',
    marked().join(','));
  check('the snapshot names the marked row',
    state().explorer.currentProcedure?.name === 'Recalculate' && state().explorer.currentProcedure?.module === 'Module1',
    JSON.stringify(state().explorer.currentProcedure));
  act('select', { startLine: 3 });
  await sleep(50);
  check('above the first procedure the bar says (Declarations) and no row is marked',
    bar() === '(Declarations)' && marked().length === 0, bar() + ' / ' + marked().join(','));
  act('select', { startLine: 44 });
  await sleep(50);
  check('a later procedure moves the mark', bar() === 'Function Describe' && marked().join(',') === 'Function Describe', bar());

  // ---- back to the flat tree ----
  act('explorerView', { view: 'tree' });
  check('the tree layout comes back on the echo', await waitFor('view', () => state().explorer.view === 'tree'), state().explorer.view);
  check('and draws the flat list again, no folders, no indent',
    !rows().some((row) => row.startsWith('folder:')) && depthOf('[data-component="Module1"]') === '0'
      && book1().filter((row) => !row.startsWith('proc:')).join(' ') === 'module:Sheet1 module:ThisWorkbook module:Module1 module:SalesRow',
    book1().join(' '));
  check('the mark survives the layout', marked().join(',') === 'Function Describe', marked().join(','));
  const refused = act('explorerView', { view: 'sideways' });
  check('a layout that does not exist is refused', !refused.did, refused.detail);

  return { pass: checks.every((one) => one.ok), checks };
})()`;

await runPageProbe({ label: "xlide-folders", drive: DRIVE });
