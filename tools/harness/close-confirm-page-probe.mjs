// Drives the BUILT editor page (ui/editor/dist) in a headless browser and walks the
// close-confirm flow end to end against the page's demo transport: a dirty close asks,
// Escape and Cancel keep the tab, Don't Save and Save close it, questions queue one at a
// time, and a repeated ask is deduplicated. Prints a JSON verdict {pass, checks} on stdout
// and exits nonzero when any check fails.
//
// No dependencies: Node's own http server serves the dist, Edge provides the browser, and
// the DevTools protocol runs over Node's global WebSocket. Invoked by Test-CloseConfirm.ps1.


import { runPageProbe } from "./page-probe.mjs";

// What the page is asked to do, and what must be true after each step. Runs inside the
// browser; the return value is the probe's verdict. The queue step presses a second tab
// while the question is up - a user cannot click through the backdrop, but this is exactly
// the shape a Close Others produces, so the synthetic press stands in for that message.

const DRIVE = `(async () => {
  const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
  const tabs = () => [...document.querySelectorAll('.tab')].map((tab) => tab.dataset.module);
  const press = (name) => {
    const tab = [...document.querySelectorAll('.tab')].find((one) => one.dataset.module === name);
    const x = tab.querySelector('.tab-close');
    const opts = { bubbles: true, cancelable: true, button: 0, pointerId: 1 };
    x.dispatchEvent(new PointerEvent('pointerdown', opts));
    x.dispatchEvent(new PointerEvent('pointerup', opts));
  };
  const answer = (label) =>
    [...document.querySelectorAll('#close-confirm-buttons button')].find((b) => b.textContent === label);
  const title = () => document.getElementById('close-confirm-title')?.textContent ?? '';
  const asking = () => !!document.getElementById('close-confirm-backdrop');

  const checks = [];
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });

  for (let waited = 0; tabs().length < 2 && waited < 20000; waited += 200) await sleep(200);
  check('the demo opens two dirty tabs', tabs().length === 2, tabs().join(','));

  press('Module2');
  check('a dirty close asks instead of closing', asking() && tabs().length === 2);
  check('the question names the module', title().includes('Module2'), title());
  check('the answers are Save, Do not Save, Cancel',
    [...document.querySelectorAll('#close-confirm-buttons button')].map((b) => b.textContent).join('|')
      === "Save|Don't Save|Cancel");
  check('Save holds the focus', document.activeElement?.textContent === 'Save');

  document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
  check('Escape cancels and the tab stays', !asking() && tabs().length === 2);

  press('Module2');
  press('Module2');
  press('Module1');
  check('a repeated ask is asked once, the other queued',
    document.querySelectorAll('#close-confirm-backdrop').length === 1 && title().includes('Module2'));

  answer("Don't Save").click();
  check('Do not Save closes the tab', tabs().join(',') === 'Module1', tabs().join(','));
  check('the queued question follows', asking() && title().includes('Module1'), title());

  answer('Cancel').click();
  check('Cancel keeps the tab', !asking() && tabs().join(',') === 'Module1');

  const first = [...document.querySelectorAll('.tab')].find((tab) => tab.dataset.module === 'Module1');
  first.dispatchEvent(new MouseEvent('auxclick', { bubbles: true, cancelable: true, button: 1 }));
  check('a middle-click close asks too', asking());

  answer('Save').click();
  check('Save closes the tab', tabs().length === 0 && !asking());

  return { pass: checks.every((one) => one.ok), checks };
})()`;

await runPageProbe({
  label: "xlide-close-confirm",
  drive: DRIVE,
});
