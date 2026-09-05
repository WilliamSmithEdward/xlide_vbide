// Drives the BUILT editor page (ui/editor/dist) in a headless browser and reads the count the
// Problems tab wears against the page's demo transport.
//
// The demo publishes one finding, a warning on Module1, so the tab should carry a 1; pressing the
// warnings toggle out should take the number away, since the tab counts what the list shows and
// nothing else (the owner, 2026-09-05: "use sum of whatever is enabled with the three toggled");
// pressing it back should bring the 1 back. The demo's tree is the host's tree in miniature, so
// the same code runs here as in Excel.
//
// Prints a JSON verdict {pass, checks} on stdout and exits nonzero when any check fails. NO
// BACKTICKS in the drive: it is spliced into a template literal.
import { runPageProbe } from "./page-probe.mjs";

const DRIVE = `(async () => {
  const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
  const tab = () => document.querySelector('.panel-tab[data-panel="problems"]');
  const badge = () => tab()?.querySelector('.tab-badge')?.textContent ?? null;
  const toggle = (group) => document.querySelector('[data-severity-filter="' + group + '"]');
  const waitFor = async (until) => {
    for (let waited = 0; waited < 10000; waited += 100) {
      if (until()) return true;
      await sleep(100);
    }
    return false;
  };
  const checks = [];
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });

  check('the Problems tab exists', !!tab());
  const arrived = await waitFor(() => badge() === '1');
  check('the tab wears the count of the findings the pane lists', arrived && badge() === '1', badge());
  check('and says so to a screen reader', (tab()?.getAttribute('aria-label') ?? '').includes('1 shown'), tab()?.getAttribute('aria-label'));

  toggle('warnings').click();
  await waitFor(() => badge() === null);
  check('pressing the warnings toggle out takes the number off the tab, since the list shows none', badge() === null, badge());
  check('and the toggle itself still says how many it hides', (toggle('warnings').textContent ?? '').includes('1 Warning'), toggle('warnings').textContent);

  toggle('warnings').click();
  await waitFor(() => badge() === '1');
  check('pressing it back brings the number back', badge() === '1', badge());

  return { pass: checks.every((one) => one.ok), checks };
})()`;

await runPageProbe({ label: "xlide-problems-badge", drive: DRIVE });
