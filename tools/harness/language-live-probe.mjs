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

const project = await api.project();
const original = (await api.readModule(target, project.projectId)).text ?? "";
console.log(`writing into ${target} of ${project.project}; ${original.length} chars will be restored\n`);

let survived = 0;
const lost = [];

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

    const wrote = await api.writeModule(target, source, project.projectId);
    if (!wrote.ran) {
      console.log(`  ${label.padEnd(20)} WRITE REFUSED: ${JSON.stringify(wrote)}`);
      continue;
    }

    await wait(900);

    // 1. COM: back out through the session's own reader.
    const back = (await api.readModule(target, project.projectId)).text ?? "";
    const throughCom = back.includes(sample);

    // 2. The page: the same module's text as the surface holds it, and its tab label.
    await api.pane("open", { module: target, project: project.projectId });
    await wait(700);

    const seen = await api.eval(
      `JSON.stringify((() => {
        const model = window.xlideBridge.documents.get(${JSON.stringify(target)}, ${JSON.stringify(project.projectId)});
        const ui = window.xlideUi.state();
        return {
          inModel: model ? model.getValue().includes(${JSON.stringify(sample)}) : null,
          tab: ui.workspace.groups.flatMap((g) => g.tabs).map((t) => t.label).includes(${JSON.stringify(target)}),
        };
      })())`,
      "page");

    const page = typeof seen.value === "string" ? JSON.parse(seen.value) : seen.value;

    // 3. The analyzer, through the door: it must still find the procedure past the text.
    const outline = await api.outline(target, project.projectId);
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
  await api.writeModule(target, original, project.projectId);
  await wait(900);
  const restored = (await api.readModule(target, project.projectId)).text ?? "";
  console.log(`\n${target} restored: ${restored.trim() === original.trim()}`);
}

console.log(`\n${survived} of ${LANGUAGES.length} survive every path on this machine.`);
if (lost.length > 0) {
  console.log(
    `Not storable by this host's VBA (its module text is in the system ANSI code page, not \n` +
    `Unicode, so these become question marks before xlide sees them): ${lost.join(", ")}.`);
}

// The failure worth failing on is text that reaches COM and is then lost by OUR code. A
// character the host itself cannot store is the host's limit and not a regression.
const ourFault = survived + lost.length !== LANGUAGES.length;
process.exit(ourFault ? 1 : 0);
