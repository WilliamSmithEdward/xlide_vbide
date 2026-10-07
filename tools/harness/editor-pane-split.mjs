// Real pointer regression: Immediate below the editor, above a separate lower pane group.
import assert from 'node:assert/strict';
import { writeFile } from 'node:fs/promises';
import { open, waitFor } from './xlide-api.mjs';
import { loopbackUrl } from './loopback-url.mjs';

const api = await open({ pid: Number(process.argv[2]) });
await waitFor('editor surface', async () => (await api.state()).surfaceReady, { budgetMs: 15000 });
await api.ask(`window.xlideBridge.activateModule('Helper')`);
await waitFor('a code editor', async () => (await api.state()).shownModule === 'Helper', { budgetMs: 10000 });
const geometry = `(() => {
  const rect = el => { const r = el.getBoundingClientRect(); return { x:r.x, y:r.y, width:r.width, height:r.height }; };
  const groups = [...document.querySelectorAll('#dock-bottom .dock-group')];
  return { groups: groups.map(g => ({ tabs:[...g.querySelectorAll('.panel-tab')].map(t => t.dataset.panel), rect:rect(g) })),
    editor:rect(document.getElementById('editor-area')), divider:!!document.querySelector('#dock-bottom .group-splitter') };
})()`;
const verify = async () => {
  const state = await api.ask(geometry);
  assert.equal(state.groups.length, 2, JSON.stringify(state));
  assert.deepEqual(state.groups[0].tabs, ['immediate']);
  assert.ok(state.groups[1].tabs.includes('problems'));
  assert.ok(!state.groups[1].tabs.includes('immediate'));
  assert.ok(state.editor.height > 90);
  assert.ok(state.groups[0].rect.y >= state.editor.y + state.editor.height - 1);
  assert.ok(state.groups[1].rect.y >= state.groups[0].rect.y + state.groups[0].rect.height);
  assert.ok(state.divider);
  return state;
};
if (process.argv[3] === 'restore') {
  await waitFor('two restored bottom pane groups', async () => {
    try { await verify(); return true; } catch { return false; }
  }, { budgetMs: 15000 });
  console.log('PASS: a fresh Excel process restored Immediate below code and above the separate lower group.');
}

if (process.argv[3] !== 'restore') {
await api.ask(`window.xlideBridge.shell.docks.movePaneTo('immediate','top')`);
const targets = await (await fetch(loopbackUrl(api.devtoolsPort, '/json'))).json();
const target = targets.find(t => t.title === 'xlide editor');
assert.ok(target, 'live editor DevTools target');
const socket = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((resolve,reject) => { socket.onopen=resolve; socket.onerror=reject; });
let next=0;
const pending=new Map();
socket.onmessage=({data}) => {
  const reply=JSON.parse(data), waiter=pending.get(reply.id);
  if (!waiter) return;
  pending.delete(reply.id);
  if (reply.error) waiter.reject(new Error(reply.error.message)); else waiter.resolve(reply.result);
};
const send=(method,params) => new Promise((resolve,reject) => {
  const id=++next; pending.set(id,{resolve,reject}); socket.send(JSON.stringify({id,method,params}));
});
const mouse=(type,point,buttons=1) => send('Input.dispatchMouseEvent', {type,...point,button:'left',buttons,clickCount:type==='mouseMoved'?0:1});
try {
  const drag = async () => {
  const points=await api.ask(`(() => {
    const tab=document.querySelector('[data-panel="immediate"]'), a=tab.getBoundingClientRect(), b=document.getElementById('editor-area').getBoundingClientRect();
    return { from:{x:a.x+a.width/2,y:a.y+a.height/2}, middle:{x:b.x+b.width/2,y:b.y+b.height/2} };
  })()`);
  await mouse('mousePressed',points.from);
  await mouse('mouseMoved',{x:points.from.x+12,y:points.from.y});
  await mouse('mouseMoved',points.middle);
  const petal=await api.ask(`(() => { const r=document.querySelector('.drop-compass .drop-petal-bottom').getBoundingClientRect(); return {x:r.x+r.width/2,y:r.y+r.height/2}; })()`);
  await mouse('mouseMoved',petal);
  const preview = await api.ask(`(() => { const p=document.querySelector('.drop-petal-bottom'), c=document.querySelector('.drop-compass'), o=document.querySelector('.drop-overlay'); return { lit:p?.classList.contains('lit'), compassHidden:c?.hidden, overlay:o?.className, dragging:!!document.querySelector('.panel-tab.dragging'), editor:document.getElementById('editor-area').getBoundingClientRect().toJSON() }; })()`);
  assert.ok(preview.lit && preview.overlay === 'drop-overlay drop-overlay-new', `bottom compass offers a separate group: ${JSON.stringify({points,petal,preview})}`);
  await mouse('mouseReleased',petal,0);
  };
  await drag();
  const before=await verify();
  await api.ask(`document.querySelector('#dock-bottom .group-splitter').dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowDown',bubbles:true,cancelable:true}))`);
  const after=await verify();
  assert.ok(after.groups[0].rect.height > before.groups[0].rect.height, 'Immediate has an independent resize divider');
  assert.ok(after.groups[1].rect.height < before.groups[1].rect.height);
  await drag();
  assert.deepEqual(await verify(), after, 'dropping the adjacent pane again leaves its geometry unchanged');
  const shot = await send('Page.captureScreenshot', { format:'png' });
  await writeFile(new URL('../../artifacts/pane-split-preview.png', import.meta.url), Buffer.from(shot.data,'base64'));
  console.log('PASS: real pointer drop put Immediate below code, separately above Problems; independent resizing works.');
} finally { socket.close(); }
}
