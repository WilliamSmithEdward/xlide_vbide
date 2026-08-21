import { open, wait } from "./xlide-api.mjs";
const api = await open();
const doctor = await api.doctor();
console.log("page stamp now:", doctor.pageBuildStamp);
const project = await api.project();
const first = (await api.projects()).projects[0].project;
await api.pane("open", { module: "HelpersExtra", project: project.projectId }).catch(() => {});
await wait(2500);
const probe = [
  '(async () => {',
  '  const ed = globalThis.xlideBridge.workspace.activeEditor();',
  '  const model = ed.getModel();',
  '  const settle = () => new Promise((go) => setTimeout(go, 40));',
  '  const lines = model.getLineCount();',
  '  let at = -1;',
  '  for (let i = 1; i <= lines; i++) {',
  '    const one = model.getLineContent(i);',
  '    if (/^\s+\S/.test(one) && !/\b(if|for|do|while|with|select|sub|function)\b/i.test(one)) { at = i; break; }',
  '  }',
  '  if (at < 0) { return { error: "no indented line" }; }',
  '  ed.setPosition({ lineNumber: at, column: model.getLineMaxColumn(at) });',
  '  ed.focus();',
  '  ed.trigger("keyboard", "type", { text: String.fromCharCode(10) });',
  '  await settle();',
  '  const afterOne = model.getLineContent(at + 1);',
  '  ed.trigger("keyboard", "type", { text: String.fromCharCode(10) });',
  '  await settle();',
  '  return { line: at, source: JSON.stringify(model.getLineContent(at)),',
  '           afterFirstEnter: JSON.stringify(afterOne),',
  '           afterSecondEnter: JSON.stringify(model.getLineContent(at + 1)) };',
  '})()',
].join("\n");
console.log(JSON.stringify(await api.ask(probe)));
