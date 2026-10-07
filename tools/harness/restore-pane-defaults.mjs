// Verify the toolbar's confirmation gate against an actual hosted editor and a fresh process.
import assert from 'node:assert/strict';
import { open, waitFor } from './xlide-api.mjs';
const api = await open({ pid: Number(process.argv[2]) });
await waitFor('editor surface', async () => (await api.state()).surfaceReady, { budgetMs:15000 });
await api.ask(`window.xlideBridge.activateModule('Helper')`);
await waitFor('Helper editor', async () => (await api.state()).shownModule === 'Helper', { budgetMs:10000 });
const snapshot = () => api.ask(`JSON.parse(localStorage.getItem('xlide.docks.v1'))`);
const checkDefaults = async () => {
  const layout = await snapshot();
  assert.deepEqual(layout.closed, []);
  assert.deepEqual(layout.sizes, {left:260,right:340,top:200,bottom:200});
  assert.equal(layout.sides.top, null);
  assert.equal(layout.sides.right, null);
  assert.deepEqual(layout.sides.left.children.map(g=>g.tabs), [['explorer'],['properties']]);
  assert.deepEqual(layout.sides.left.sizes, [0.62,0.38]);
  assert.equal(layout.sides.bottom.active, 'problems');
  assert.ok(layout.sides.bottom.tabs.includes('watch'));
  const rendered = await api.ask(`({ topHidden:document.getElementById('dock-top').hidden, groups:document.querySelectorAll('#dock-bottom .dock-group').length, active:document.querySelector('#dock-bottom .panel-tab.active')?.dataset.panel })`);
  assert.deepEqual(rendered,{topHidden:true,groups:1,active:'problems'});
  assert.equal((await api.settings()).formatIndentSize,3,'editor setting survives the pane reset');
};
if (process.argv[3] === 'restore') {
  await checkDefaults();
  console.log('PASS: confirmed pane defaults persist into a fresh Excel process.');
} else {
  await api.settings({formatIndentSize:3});
  await api.ask(`(() => { const d=window.xlideBridge.shell.docks; d.movePaneTo('immediate','top'); d.close('watch'); document.getElementById('dock-left-splitter').dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowRight',bubbles:true,cancelable:true})); })()`);
  const custom = await snapshot();
  assert.ok(custom.closed.includes('watch') && custom.sides.top && custom.sizes.left!==260);
  const open = () => api.ask(`document.querySelector('#toolbar [data-command="restorePaneDefaults"]').click()`);
  const choose = label => api.ask(`([...document.querySelectorAll('#restore-panes-card button')].find(b=>b.textContent===${JSON.stringify(label)})).click()`);
  await open();
  const modal = await api.ask(`({role:document.getElementById('restore-panes-card')?.getAttribute('role'), focus:document.activeElement?.textContent, text:document.getElementById('restore-panes-card')?.textContent})`);
  assert.equal(modal.role,'alertdialog');
  assert.equal(modal.focus,'Cancel');
  assert.ok(modal.text.includes('Are you sure'));
  const spacing = await api.ask(`(() => {
    const card=document.getElementById('restore-panes-card'), style=getComputedStyle(card), box=card.getBoundingClientRect();
    const title=card.querySelector('.modal-title').getBoundingClientRect(), buttons=card.querySelector('.modal-buttons').getBoundingClientRect();
    return { padding:[style.paddingTop,style.paddingRight,style.paddingBottom,style.paddingLeft].map(parseFloat),
      titleInset:title.left-box.left, topInset:title.top-box.top, buttonInset:box.right-buttons.right, bottomInset:box.bottom-buttons.bottom,
      withinViewport:box.width<=innerWidth && box.height<=innerHeight };
  })()`);
  assert.ok(spacing.padding.every(p=>p>=16),JSON.stringify(spacing));
  assert.ok(spacing.titleInset>=16 && spacing.topInset>=16 && spacing.buttonInset>=16 && spacing.bottomInset>=16,JSON.stringify(spacing));
  assert.ok(spacing.withinViewport);
  assert.deepEqual(await snapshot(),custom,'opening confirmation does not reset');
  await choose('Cancel');
  assert.deepEqual(await snapshot(),custom,'Cancel preserves the layout');
  await open();
  await api.ask(`document.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true,cancelable:true}))`);
  assert.deepEqual(await snapshot(),custom,'Escape preserves the layout');
  assert.equal(await api.ask(`!!document.getElementById('restore-panes-card')`),false);
  await open();
  await api.ask(`document.getElementById('restore-panes-backdrop').dispatchEvent(new MouseEvent('mousedown',{bubbles:true}))`);
  assert.deepEqual(await snapshot(),custom,'backdrop dismissal preserves the layout');
  await open();
  await choose('Restore defaults');
  await checkDefaults();
  assert.equal((await api.state()).shownModule,'Helper','reset keeps the active code module');
  console.log('PASS: toolbar confirmation, Cancel/Escape/backdrop safety, confirmed pane reset, and preserved editor settings.');
}
