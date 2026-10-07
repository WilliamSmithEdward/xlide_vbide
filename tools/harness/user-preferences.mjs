// Issue #76: verify visible pane state in fresh Office processes, not only JSON serialization.
import { readFile, writeFile } from 'node:fs/promises';
import { open, waitFor } from './xlide-api.mjs';

const [pid, phase, expectedFile] = process.argv.slice(2);
const api = await open({ pid: Number(pid) });
await waitFor('the editor surface', async () => (await api.state()).surfaceReady, { budgetMs: 15000 });

const snapshot = () => api.ask(`(() => {
  const layout = JSON.parse(localStorage.getItem('xlide.docks.v1'));
  const rendered = Object.fromEntries(['left', 'right', 'top', 'bottom'].map(side => {
    const element = document.getElementById('dock-' + side);
    return [side, { hidden: element.hidden, size: element.style.getPropertyValue('--dock-size'),
      tabs: [...element.querySelectorAll('.panel-tab')].map(tab => tab.dataset.panel),
      active: [...element.querySelectorAll('.panel-tab.active')].map(tab => tab.dataset.panel) }];
  }));
  return { layout, rendered, sourceControl: JSON.parse(localStorage.getItem('xlide.scm.v1') ?? 'null'),
    editorGroups: JSON.parse(localStorage.getItem('xlide.workspace.v1') ?? 'null'),
    editors: { count: document.querySelectorAll('.editor-group').length,
      active: [...document.querySelectorAll('.editor-group')].findIndex(group => group.classList.contains('active-group')),
      splitters: [...document.querySelectorAll('.group-splitter')].map(splitter => splitter.getAttribute('aria-orientation')) } };
})()`);

if (phase === 'seed') {
  await api.settings({ formatIndentSize: 3 });
  await api.ask(`(() => {
    const docks = window.xlideBridge.shell.docks;
    docks.movePaneTo('immediate', 'right');
    docks.movePaneTo('properties', 'top');
    docks.close('watch');
    docks.reveal('scm');
    const left = document.getElementById('dock-left-splitter');
    for (let i = 0; i < 3; i++) left.dispatchEvent(new KeyboardEvent('keydown', { key:'ArrowLeft', bubbles:true, cancelable:true }));
    const rail = document.getElementById('scm-splitter');
    rail.dispatchEvent(new KeyboardEvent('keydown', { key:'ArrowRight', bubbles:true, cancelable:true }));
    docks.reveal('locals');
    return true;
  })()`);
  await api.ask(`window.xlideBridge.activateModule('Helper')`);
  await waitFor('a second code module', async () => (await api.state()).shownModule === 'Helper', { budgetMs: 5000 });
  await api.ask(`(() => {
    window.xlideBridge.workspace.splitActive('right');
    document.querySelector('.group-splitter')?.dispatchEvent(new KeyboardEvent('keydown', { key:'ArrowLeft', bubbles:true, cancelable:true }));
  })()`);
  const expected = await snapshot();
  if (expected.rendered.right.hidden || expected.rendered.top.hidden || !expected.layout.closed.includes('watch'))
    throw new Error(`The seed did not change the visible layout: ${JSON.stringify(expected)}`);
  if (!expected.rendered.bottom.active.includes('locals')) throw new Error('Locals did not become the active pane');
  if (!expected.sourceControl) throw new Error('Source Control divider was not saved');
  if (expected.editors.count !== 2 || expected.editors.active !== 1) throw new Error('The editor group geometry was not customized');
  await writeFile(expectedFile, JSON.stringify(expected, null, 2));
  console.log('PASS: customized docking, active panes, visibility, sizes, divider, and settings.');
} else {
  const expected = JSON.parse(await readFile(expectedFile, 'utf8'));
  try {
    await waitFor('the shared layout to restore in the visible UI', async () =>
      JSON.stringify(await snapshot()) === JSON.stringify(expected), { budgetMs: 15000 });
  } catch (error) {
    console.error('Expected:', JSON.stringify(expected), '\nActual:', JSON.stringify(await snapshot()));
    throw error;
  }
  const settings = await api.settings();
  if (settings.formatIndentSize !== 3) throw new Error(`The shared setting did not restore: ${JSON.stringify(settings)}`);
  console.log(`PASS: process ${pid} restored visible pane state, geometry, divider, and settings.`);
}
