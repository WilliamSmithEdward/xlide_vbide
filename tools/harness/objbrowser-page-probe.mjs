// Drives the BUILT Object Browser page (ui/editor/dist, ?view=objbrowser) in a headless
// browser against its demo transport and walks the pinned behaviours: the palette page
// boots without the editor shell, the three search scopes act on the right panes, All
// pulls a whole matched group in eagerly, the details pane fills its signature, context,
// and description rows (hiding the empty ones), and the splitter answers the keyboard.
// Prints a JSON verdict {pass, checks} on stdout and exits nonzero when any check fails.
//
// No dependencies: Node's own http server serves the dist, Edge provides the browser, and
// the DevTools protocol runs over Node's global WebSocket. Invoked by Test-ObjectBrowser.ps1.


import { runPageProbe } from "./page-probe.mjs";

const DRIVE = `(async () => {
  const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
  const rows = (pane) => [...document.querySelectorAll('#objbrowser-' + pane + ' .objbrowser-row')];
  const names = (pane) => rows(pane).map((row) => row.querySelector('.objbrowser-name').textContent);
  const picker = () => document.getElementById('objbrowser-library');
  const detailRow = (part) => document.getElementById('objbrowser-detail-' + part);
  const type = (text) => {
    const box = document.getElementById('objbrowser-search');
    box.value = text;
    box.dispatchEvent(new Event('input', { bubbles: true }));
  };
  const mode = (value) => {
    const pick = document.getElementById('objbrowser-scope');
    pick.value = value;
    pick.dispatchEvent(new Event('change', { bubbles: true }));
  };
  const clickRow = (pane, name) =>
    rows(pane).find((row) => row.querySelector('.objbrowser-name').textContent === name).click();

  const checks = [];
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });

  for (let waited = 0; (picker()?.options.length ?? 0) < 3 && waited < 20000; waited += 200) await sleep(200);
  check('the palette page sheds the editor shell', !document.getElementById('shell'));
  check('the demo lists the project and the libraries', picker().options.length === 3,
    [...picker().options].map((option) => option.textContent).join('|'));

  for (let waited = 0; rows('modules').length < 2 && waited < 5000; waited += 100) await sleep(100);
  check('the project scope loads its modules', names('modules').join(',') === 'Module1,ThisWorkbook',
    names('modules').join(','));
  check('the details pane and its splitter exist',
    !!document.getElementById('objbrowser-splitter')
      && !!detailRow('signature') && !!detailRow('context') && !!detailRow('description'));

  // Typing settles before the panes rebuild (the search debounces at 150ms, like the find
  // widget), so a query's effect is polled for rather than read on the next line.
  const settled = async (condition) => {
    for (let waited = 0; waited < 3000; waited += 100) {
      if (condition()) { return true; }
      await sleep(100);
    }
    return condition();
  };

  // Group: filters the left pane by name and leaves the members pane alone.
  type('module');
  await settled(() => names('modules').join(',') === 'Module1');
  check('Group filters the types pane', names('modules').join(',') === 'Module1', names('modules').join(','));
  check('Group leaves the members pane alone', rows('members').length === 0);

  // All: a group whose own name matches brings its whole membership along, loading it.
  mode('all');
  for (let waited = 0; rows('members').length < 2 && waited < 5000; waited += 100) await sleep(100);
  check('All pulls the whole matched group in', names('members').join(',') === 'Greet,Total',
    names('members').join(','));
  check('spanning members name their group',
    rows('members').every((row) => row.querySelector('.objbrowser-context').textContent === 'Module1'));

  // Object: filters the selected type's members and leaves the list alone.
  mode('object');
  type('');
  await settled(() => names('modules').length === 2);
  clickRow('modules', 'Module1');
  for (let waited = 0; rows('members').length < 2 && waited < 5000; waited += 100) await sleep(100);
  type('gr');
  await settled(() => names('members').join(',') === 'Greet');
  check('Object filters the selected types members', names('members').join(',') === 'Greet',
    names('members').join(','));
  check('Object leaves the types pane alone', names('modules').length === 2);

  // Details: a project member fills signature and context; the empty description hides.
  clickRow('members', 'Greet');
  check('the signature row carries the declaration',
    detailRow('signature').textContent === 'Public Sub Greet(name As String)',
    detailRow('signature').textContent);
  check('the context row names the module and line',
    detailRow('context').textContent === 'Member of scratch.xlsm.Module1, line 3',
    detailRow('context').textContent);
  check('an empty description row hides', detailRow('description').hidden);

  // Description: a library member that carries one shows it in the third row.
  type('');
  await settled(() => names('members').length >= 2);
  picker().selectedIndex = 1;
  picker().dispatchEvent(new Event('change', { bubbles: true }));
  for (let waited = 0; names('modules').length < 3 && waited < 5000; waited += 100) await sleep(100);
  clickRow('modules', 'Range');
  for (let waited = 0; rows('members').length < 2 && waited < 5000; waited += 100) await sleep(100);
  clickRow('members', 'Address');
  check('a populated description row shows',
    !detailRow('description').hidden && detailRow('description').textContent === 'Returns the address.',
    detailRow('description').textContent);

  // The splitter answers the keyboard.
  const detail = document.getElementById('objbrowser-detail');
  const splitter = document.getElementById('objbrowser-splitter');
  const before = detail.getBoundingClientRect().height;
  splitter.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true, cancelable: true }));
  check('ArrowUp on the splitter grows the details pane', detail.getBoundingClientRect().height > before);

  return { pass: checks.every((one) => one.ok), checks };
})()`;

await runPageProbe({
  label: "xlide-objbrowser",
  path: "/index.html?view=objbrowser",
  drive: DRIVE,
});
