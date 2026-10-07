// The contributor's exact Stop / edit 2 to 100 / continue / reopen case in a real Excel host.
import { open, waitFor } from './xlide-api.mjs';

const api = await open({ pid: Number(process.argv[2]) });
await waitFor('the disposable procedure to stop', async () =>
  (await api.state()).debugMode === 'break', { budgetMs: 30000 });
const original = await api.native();
if (original.activeModule !== 'Issue71Probe')
  throw new Error(`Stopped in ${original.activeModule}, not Issue71Probe`);

const edited = await api.ask(`(() => {
  const model = window.xlideBridge.workspace.activeEditor()?.getModel();
  if (!model) return null;
  const old = model.getValue();
  if (!old.includes('Debug.Print 2')) return old;
  model.setValue(old.replace('Debug.Print 2', 'Debug.Print 100'));
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
console.log('PASS: XLIDE edit reached VBA while stopped, survived reopening, and Continue printed 100.');
