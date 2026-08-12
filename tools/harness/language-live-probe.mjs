/*
 * Non-ASCII through COM, the door, and the page: the paths engine/test/language.mjs cannot reach.
 *
 * The engine matrix runs headless over a pipe and proves the analyzer's half. This proves the
 * rest of the chain, which only exists when Excel is running:
 *
 *   the door     module text as an HTTP body, in and out
 *   COM          the VBA object model's reader and writer, which trade in BSTRs
 *   the page     the same text as a JSON message, a monaco model, and a rendered tab
 *   the tree     component names in the explorer
 *
 * What it MEASURES rather than asserts: VBA stores module text in the system ANSI code page, so
 * a character outside that page cannot survive a write at all. It comes back as a question mark,
 * from Excel, before xlide sees it. Which characters survive therefore depends on the machine,
 * and this reports what this machine does instead of pretending otherwise.
 */

import { open } from "./xlide-api.mjs";

const moduleArg = process.argv.indexOf("--module");
const target = moduleArg >= 0 ? process.argv[moduleArg + 1] : "HelpersExtra";

const api = await open({});
const wait = (ms) => new Promise((r) => setTimeout(r, ms));

// The same samples the engine matrix uses, which are the companion product's samples.
const LANGUAGES = [
  ["Western European", "déjà vu € œuvre Straße"],
  ["Central European", "Příliš žluťoučký kůň"],
  ["Turkish", "Türkçe deneme ğüşiöç İı"],
  ["Baltic", "Lietuviškas ąčęėįšųū"],
  ["Vietnamese", "Tiếng Việt thử nghiệm"],
  ["Cyrillic", "Проверка русского текста"],
  ["Greek", "Δοκιμή ελληνικού"],
  ["Hebrew", "בדיקת עברית"],
  ["Arabic", "اختبار العربية"],
  ["Thai", "ทดสอบภาษาไทย"],
  ["Japanese", "テスト用モジュール"],
  ["Chinese, Simplified", "中文测试模块"],
  ["Korean", "한국어 테스트"],
  ["Astral (emoji)", "🎯 target 🇯🇵"],
];

// THE PROJECT THAT HOLDS THE TARGET, not the one in front. `api.project()` answers about the
// front workbook, and with two open, which is front depends on nothing this probe controls:
// launched fresh on the gate's double session the front one was the TWIN, which has no
// HelpersExtra, and the run died on its second line (2026-08-12). Asked of every open project
// instead, and refused when the answer is ambiguous - two workbooks holding the target is the
// exact state where writing into "whichever" rewrites a stranger's module.
//
// The DISPLAY name is what everything below passes. projects() hands out projectId as a full
// path, which every route tolerates (the server resolves it) - but the page leg's
// documents.get is a raw client-side map keyed by what the host publishes, the workbook file
// name, so the path form answered null for a model the store was listing in the same breath
// (2026-08-12). The file name is the canonical project= form per the api reference.
const holders = [];
for (const one of (await api.projects()).projects) {
  const held = await api.readModule(target, one.projectId).catch(() => ({ text: null }));
  if (held.text !== null) { holders.push({ name: one.project, text: held.text }); }
}
if (holders.length !== 1) {
  console.log(holders.length === 0
    ? `no open workbook holds a module named ${target}`
    : `${holders.length} open workbooks hold a ${target} (${holders.map((h) => h.name).join(", ")}); refusing to guess`);
  console.log("\n0 passed, 1 failed");
  process.exit(1);
}
const projectName = holders[0].name;
const original = holders[0].text;
console.log(`writing into ${target} of ${projectName}; ${original.length} chars will be restored\n`);

let survived = 0;
const lost = [];
let restoredClean = false;

try {
  for (const [label, sample] of LANGUAGES) {
    // One language at a time, so a character the host cannot store cannot be blamed on another.
    const source = [
      "Option Explicit",
      "",
      `' ${sample}`,
      "",
      "Public Sub Probe()",
      `    Debug.Print "${sample}"`,
      "End Sub",
      "",
    ].join("\r\n");

    const wrote = await api.writeModule(target, source, projectName);
    if (!wrote.ran) {
      console.log(`  ${label.padEnd(20)} WRITE REFUSED: ${JSON.stringify(wrote)}`);
      continue;
    }

    await wait(900);

    // 1. COM: back out through the session's own reader.
    const back = (await api.readModule(target, projectName)).text ?? "";
    const throughCom = back.includes(sample);

    // 2. The page: the same module's text as the surface holds it, and its tab label. The
    // model is WAITED FOR, not slept for: the session's first pane open creates it
    // asynchronously, and a fixed 700ms lost that race on the gate's fresh double session
    // (2026-08-12). Only a script that SURVIVED the host is worth waiting on - its sample is
    // coming, and the question is whether our chain delivers it. One the code page already
    // dropped gets a single look for the printed line; waiting ten seconds for text that
    // cannot arrive would cost the lost scripts two minutes between them.
    await api.pane("open", { module: target, project: projectName });

    let page = { inModel: null, tab: false };
    const pageBudget = throughCom ? 10000 : 0;
    for (let waited = 0; ; waited += 250) {
      const seen = await api.eval(
        `JSON.stringify((() => {
          const model = window.xlideBridge.documents.get(${JSON.stringify(target)}, ${JSON.stringify(projectName)});
          const ui = window.xlideUi.state();
          return {
            inModel: model ? model.getValue().includes(${JSON.stringify(sample)}) : null,
            tab: ui.workspace.groups.flatMap((g) => g.tabs).map((t) => t.label).includes(${JSON.stringify(target)}),
          };
        })())`,
        "page");
      page = typeof seen.value === "string" ? JSON.parse(seen.value) : seen.value;
      if (page && page.inModel === true) { break; }
      if (waited >= pageBudget) { break; }
      await wait(250);
    }

    // 3. The analyzer, through the door: it must still find the procedure past the text.
    const outline = await api.outline(target, projectName);
    const foundProbe = (outline.procedures ?? []).some((one) => one.name === "Probe");

    const verdict = throughCom
      ? (page.inModel && foundProbe ? "ok" : "LOST AFTER COM")
      : "not storable by the host";

    if (throughCom && page.inModel && foundProbe) {
      survived++;
    } else if (!throughCom) {
      lost.push(label);
    }

    console.log(
      `  ${label.padEnd(20)} com=${String(throughCom).padEnd(5)} page=${String(page.inModel).padEnd(5)} ` +
      `outline=${String(foundProbe).padEnd(5)} ${verdict}`);
  }
} finally {
  await api.writeModule(target, original, projectName);
  await wait(900);
  const restored = (await api.readModule(target, projectName)).text ?? "";
  restoredClean = restored.trim() === original.trim();
  console.log(`\n${target} restored: ${restoredClean}`);
}

console.log(`\n${survived} of ${LANGUAGES.length} survive every path on this machine.`);
if (lost.length > 0) {
  console.log(
    `Not storable by this host's VBA (its module text is in the system ANSI code page, not \n` +
    `Unicode, so these become question marks before xlide sees them): ${lost.join(", ")}.`);
}

// The failure worth failing on is text that reaches COM and is then lost by OUR code. A
// character the host itself cannot store is the host's limit and not a regression. A module
// left holding the last sample is also ours: the exit used to shrug at a failed restore, and
// any suite after this one would then read the leftovers as the fixture (2026-08-12).
const corrupted = LANGUAGES.length - survived - lost.length;
const failures = corrupted + (restoredClean ? 0 : 1);
console.log(`\n${survived + lost.length + (restoredClean ? 1 : 0)} passed, ${failures} failed `
  + `(${survived} survived, ${lost.length} at the host's own limit, restore ${restoredClean ? "clean" : "FAILED"})`);
process.exit(failures === 0 ? 0 : 1);
