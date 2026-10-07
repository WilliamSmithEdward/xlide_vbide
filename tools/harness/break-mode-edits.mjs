// The contributor's exact Stop / edit 2 to 100 / continue / reopen case in a real Excel host.
import { open, waitFor } from './xlide-api.mjs';

const api = await open({ pid: Number(process.argv[2]) });
const phase = process.argv[3] ?? 'safe';
await waitFor('the disposable procedure to stop', async () =>
  (await api.state()).debugMode === 'break', { budgetMs: 30000 });
const original = await api.native();
if (original.activeModule !== 'Issue71Probe')
  throw new Error(`Stopped in ${original.activeModule}, not Issue71Probe`);

if (phase === 'reset') {
  const pending = await api.ask(`(() => {
    const model = window.xlideBridge.workspace.activeEditor()?.getModel();
    const old = model?.getValue();
    if (!old?.includes('Debug.Print 100')) return old;
    model.setValue(old.replace('Debug.Print 100', 'Static extra As Long\\n    Debug.Print 100'));
    return model.getValue();
  })()`);
  if (!pending?.includes('Static extra As Long')) throw new Error(`Structural edit missing: ${pending}`);
  await waitFor('the reset choice to appear', async () =>
    await api.ask(`!!document.getElementById('break-edit-backdrop')`), { budgetMs: 10000 });
  const spacing = await api.ask(`(() => {
    const card = document.getElementById('break-edit-card');
    const buttons = card?.querySelector('.modal-buttons');
    return { padding: parseFloat(getComputedStyle(card).paddingLeft),
      buttonGap: parseFloat(getComputedStyle(buttons).marginTop) };
  })()`);
  if (spacing.padding < 16 || spacing.buttonGap < 12)
    throw new Error(`Reset dialog content is crowded: ${JSON.stringify(spacing)}`);
  if ((await api.state()).debugMode !== 'break') throw new Error('Structural edit reset VBA without consent');
  if ((await api.native({ text: true })).nativeText?.includes('Static extra As Long'))
    throw new Error('Structural edit reached VBA before Reset');
  await api.ask(`document.querySelector('#break-edit-card .modal-button.primary')?.click()`);
  await api.ask(`window.xlideBridge.activateModule('CleanModule')`);
  await waitFor('another module to open', async () =>
    (await api.state()).shownModule === 'CleanModule', { budgetMs: 5000 });
  await api.ask(`window.xlideBridge.activateModule('Issue71Probe')`);
  await waitFor('the pending module to reopen', async () =>
    (await api.state()).shownModule === 'Issue71Probe', { budgetMs: 5000 });
  const reopened = await api.ask(`window.xlideBridge.workspace.activeEditor()?.getModel()?.getValue()`);
  if (!reopened?.includes('Static extra As Long'))
    throw new Error(`Refused edit disappeared after tab switch: ${reopened}`);
  if ((await api.native({ text: true })).nativeText?.includes('Static extra As Long'))
    throw new Error('Pending edit reached VBA before Reset');
  const run = await api.command('run');
  if (run.ran || (await api.state()).debugMode !== 'break')
    throw new Error(`Run used stale code while an edit was pending: ${JSON.stringify(run)}`);
  await api.ask(`(() => {
    const model = window.xlideBridge.workspace.activeEditor()?.getModel();
    // Ask for the choice again after the first was dismissed.
    model.setValue(model.getValue().replace('extra As Long', 'extra2 As Long'));
  })()`);
  await waitFor('the reset choice to reappear', async () =>
    await api.ask(`!!document.getElementById('break-edit-backdrop')`), { budgetMs: 10000 });
  await api.ask(`document.querySelector('#break-edit-card .modal-button:not(.primary)')?.click()`);
  await waitFor('Reset to apply retained code', async () =>
    (await api.state()).debugMode === 'design'
      && (await api.native({ text: true })).nativeText?.includes('Static extra2 As Long'),
  { budgetMs: 10000 });
  console.log('PASS: reset-required edit stayed visible, blocked stale Run, and applied after explicit Reset.');
  process.exit(0);
}

const edited = await api.ask(`(() => {
  const model = window.xlideBridge.workspace.activeEditor()?.getModel();
  if (!model) return null;
  const old = model.getValue();
  if (!old.includes('Debug.Print 2')) return old;
  model.setValue(old.replace('Debug.Print 1', 'Debug.Print 10').replace('Debug.Print 2', 'Debug.Print 100'));
  return model.getValue();
})()`);
if (!edited?.includes('Debug.Print 100'))
  throw new Error(`The XLIDE editor did not accept the edit: ${edited}`);

await waitFor('the changed line to reach the native module', async () => {
  const now = await api.native();
  return now.nativeContent === now.surfaceContent && now.nativeContent !== original.nativeContent;
}, { budgetMs: 10000 });
if ((await api.state()).debugMode !== 'break')
  throw new Error('Writing the changed statement reset the debugger');

await api.ask(`(() => {
  const model = window.xlideBridge.workspace.activeEditor()?.getModel();
  model.setValue(model.getValue().replace('Debug.Print 100', 'Debug.Print 200\\n    Debug.Print 100'));
})()`);
await waitFor('the new statement to reach VBA', async () =>
  (await api.native({ text: true })).nativeText?.includes('Debug.Print 200'), { budgetMs: 10000 });
await api.ask(`(() => {
  const model = window.xlideBridge.workspace.activeEditor()?.getModel();
  model.setValue(model.getValue().split('\\r\\n').filter(line => !line.includes('Debug.Print 200')).join('\\r\\n'));
})()`);
await waitFor('the removed statement to leave VBA', async () =>
  !(await api.native({ text: true })).nativeText?.includes('Debug.Print 200'), { budgetMs: 10000 });
if ((await api.state()).debugMode !== 'break')
  throw new Error('Statement insertion or deletion reset the debugger');

await api.ask(`window.xlideBridge.activateModule('CleanModule')`);
await waitFor('another module to open', async () =>
  (await api.state()).shownModule === 'CleanModule', { budgetMs: 5000 });
await api.ask(`window.xlideBridge.activateModule('Issue71Probe')`);
await waitFor('the edited module to reopen', async () =>
  (await api.state()).shownModule === 'Issue71Probe', { budgetMs: 5000 });
const reopened = await api.ask(`window.xlideBridge.workspace.activeEditor()?.getModel()?.getValue()`);
if (!reopened?.includes('Debug.Print 100') || reopened.includes('Debug.Print 2'))
  throw new Error(`The edit reverted after switching modules: ${reopened}`);

const ran = await api.command('run');
if (!ran.ran) throw new Error(`Continue was refused: ${JSON.stringify(ran)}`);
await waitFor('the changed statement to execute', async () => {
  const output = (await api.immediate()).text ?? '';
  return /(?:^|\r?\n)\s*100\s*(?:\r?\n|$)/.test(output);
}, { budgetMs: 10000 });
const output = (await api.immediate()).text ?? '';
if (/(?:^|\r?\n)\s*2\s*(?:\r?\n|$)/.test(output))
  throw new Error(`The old statement also executed: ${output}`);
await waitFor('the first run to finish', async () =>
  (await api.state()).debugMode === 'design', { budgetMs: 10000 });
console.log('PASS: statement replacements before and after Stop, insertion, deletion, tab switch, and Continue.');
