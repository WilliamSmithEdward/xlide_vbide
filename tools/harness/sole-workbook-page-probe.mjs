// Closing the last tab folds the tree back, and the sole workbook is the exception.
//
// With several workbooks open, folding them away when nothing is left open is real tidying. With
// one, folding it leaves a tree of a single closed row: every module in the project behind a click
// that has one possible answer (the developer, 2026-08-10).
//
// Two branches, and the demo could only reach the first until it learned `?books=1`, so this is
// the half that had no headless coverage at all. The other half is checked in tree-page-probe.
//
// Prints a JSON verdict {pass, checks} on stdout and exits nonzero when any check fails.

import { runPageProbe } from "./page-probe.mjs";

const DRIVE = `(async () => {
  const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
  const act = (name, args) => window.xlideUi.act(name, args ?? {});
  const books = () => window.xlideUi.state().explorer.workbooks;
  const tabs = () => [...document.querySelectorAll('.tab')].length;

  const checks = [];
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });

  for (let waited = 0; books().length === 0 && waited < 20000; waited += 200) await sleep(200);

  check('the demo is serving a single workbook', books().length === 1,
    books().map((one) => one.name).join(', '));

  act('expandWorkbook', { workbook: books()[0].name, open: true });
  await sleep(400);
  check('and it is open to start with', books()[0].expanded === true, JSON.stringify(books()[0].expanded));

  // Every tab, however many the demo opened.
  for (let i = 0; i < 12 && tabs() > 0; i++) {
    const closed = act('closeActive');
    await sleep(250);
    if (!closed.did) {
      // A dirty tab asks first; the demo's tabs are dirty on purpose.
      act('answerCloseConfirm', { answer: 'discard' });
      await sleep(250);
    }
  }
  await sleep(600);

  check('every tab is closed', tabs() === 0, tabs() + ' left');

  // WITHOUT THIS THE CHECK BELOW IS VACUOUS. Folding happens when the workspace reports itself
  // empty; if the demo never reports it, nothing folds and "it stayed open" is true for the wrong
  // reason.
  check('and the workspace reported itself empty, so folding was actually attempted',
    document.getElementById('shell').classList.contains('empty'),
    'the shell is not in its empty state; nothing would have folded either way');

  check('THE SOLE WORKBOOK STAYS OPEN', books()[0].expanded === true,
    'it folded away, leaving the whole project behind one closed row');

  return { pass: checks.every((one) => one.ok), checks };
})()`;

await runPageProbe({
  label: "xlide-sole-workbook",
  path: "/index.html?books=1",
  drive: DRIVE,
});
