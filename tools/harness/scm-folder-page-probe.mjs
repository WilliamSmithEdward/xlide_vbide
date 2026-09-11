// The Source Control pane's folder choice, headless: served as /index.html?scm=noFolder the demo
// transport answers with git present and no folder chosen, suggesting a folder whose path is far
// longer than a button should be. The owner's screenshot (2026-09-10) showed the whole path
// stretching "Use <folder>" across the pane; the button now draws it shortened from the middle -
// the drive and the folder's own name - and keeps the whole path in its tooltip.
//
//   node tools\harness\scm-folder-page-probe.mjs      (from PowerShell; see page-probe.mjs)

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

  const WHOLE = "C:\\\\Users\\\\Demo Developer\\\\OneDrive - Contoso\\\\Documents\\\\Finance\\\\Quarterly Close\\\\Book1";

  await until(() => document.querySelectorAll(".tab").length >= 2, 20000);
  const ui = window.xlideUi;
  const root = document.querySelector("#scm");
  document.querySelector('.panel-tab[data-panel="scm"]').click();

  const state = await until(() => { const s = ui.state().scm; return s && s.state === "noFolder" ? s : null; });
  check("the demo answers noFolder and the pane adopts it", !!state, JSON.stringify(ui.state().scm));

  const use = await until(() => root.querySelector("#scm-use-folder"));
  check("the pane offers the suggested folder as a button", !!use);

  const label = use ? use.textContent : "";
  check("the button draws the drive and the folder's own name, the middle given way",
    label === "Use C:\\\\Users\\\\...\\\\Quarterly Close\\\\Book1", label);
  check("so the label is a fraction of the path it stands for",
    label.length <= 44 && label.length < WHOLE.length / 2, label.length + " of " + WHOLE.length);
  check("and the tooltip keeps the path whole", use && use.title === WHOLE, use && use.title);

  const choose = root.querySelector("#scm-browse");
  check("Choose folder stands beside it", !!choose && choose.textContent === "Choose folder",
    choose && choose.textContent);

  return { pass: checks.every((one) => one.ok), checks };
  } catch (error) {
    checks.push({ name: "the drive threw", ok: false, detail: String(error && error.stack || error) });
    return { pass: false, checks };
  }
})()`;

await runPageProbe({ label: "xlide-scm-folder", drive: DRIVE, path: "/index.html?scm=noFolder" });
