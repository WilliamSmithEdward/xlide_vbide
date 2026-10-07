// End-to-end Locals tree check for issue #70, run by Test-LocalsObjectExpansion.ps1.
import { open, waitFor } from './xlide-api.mjs';
import { loopbackUrl } from './loopback-url.mjs';
const api = await open({ pid: Number(process.argv[2]) });
const targets = await (await fetch(loopbackUrl(api.devtoolsPort, '/json'))).json();
const target = targets.find(row => row.title === 'xlide editor');
if (!target) throw new Error('No live XLIDE editor target');
const socket = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
let nextId = 0;
const pending = new Map();
socket.onmessage = ({ data }) => {
  const reply = JSON.parse(data);
  const waiter = pending.get(reply.id);
  if (!waiter) return;
  pending.delete(reply.id);
  if (reply.error || reply.result?.exceptionDetails) waiter.reject(new Error(JSON.stringify(reply)));
  else waiter.resolve(reply.result?.result?.value);
};
function page(expression) {
  return new Promise((resolve, reject) => {
    const id = ++nextId;
    pending.set(id, { resolve, reject });
    socket.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression, returnByValue: true } }));
  });
}
const rows = () => page(`Array.from(document.querySelectorAll('#locals-table .locals-row:not(.locals-header)'),
  row => ({ text: row.textContent, depth: row.querySelector('.locals-name')?.style.paddingLeft,
    toggle: row.querySelector('.locals-toggle')?.getAttribute('aria-expanded') }))`);
const immediateRows = () => page(`document.querySelectorAll('#immediate-log > div').length`);
try {
  const immediateBefore = await immediateRows();
  await page(`window.xlideBridge.runCommand({id:'stepInto',target:'host'})`);
  await waitFor('Locals rows after debugger step', async () =>
    (await rows()).some(row => row.text.includes('sample')),
    { budgetMs: 10000 });
  if (!(await rows()).some(row => row.text.includes('sampleCollection/Collection')))
    throw new Error('The live panel did not show the Collection local');
  const clicked = await page(`(() => { const button = [...document.querySelectorAll('#locals-table .locals-toggle')]
    .find(x => x.getAttribute('aria-label')?.includes('sample')); button?.click(); return !!button; })()`);
  if (!clicked) throw new Error('No Collection expander in the live Locals panel');
  await waitFor('Collection members in Locals panel', async () =>
    (await rows()).some(row => row.text.includes('Item 1') && row.depth === '14px'),
    { budgetMs: 10000 });
  const expanded = await rows();
  for (const member of ['Item 1', 'Item 2']) {
    const matches = expanded.filter(row => row.text.includes(member) && row.depth === '14px');
    if (matches.length !== 1) throw new Error(`Expected one ${member} member, found ${matches.length}: ${JSON.stringify(expanded)}`);
  }
  const childClicked = await page(`(() => { const button = [...document.querySelectorAll('#locals-table .locals-toggle')]
    .find(x => x.getAttribute('aria-label')?.includes('Item 1')); button?.click(); return !!button; })()`);
  if (!childClicked) throw new Error('No nested object expander');
  await waitFor('nested Collection member', async () =>
    (await rows()).some(row => row.text.includes('nested value') && row.depth === '28px'),
    { budgetMs: 10000 });
  const nested = await rows();
  if (nested.filter(row => row.text.includes('nested value') && row.depth === '28px').length !== 1)
    throw new Error(`Expected one nested value: ${JSON.stringify(nested)}`);
  await page(`(() => { const button = [...document.querySelectorAll('#locals-table .locals-toggle')]
    .find(x => x.getAttribute('aria-label')?.includes('Item 1')); button.click(); return true; })()`);
  await waitFor('nested member collapse', async () =>
    !(await rows()).some(row => row.text.includes('nested value')),
    { budgetMs: 10000 });
  await page(`(() => { const button = [...document.querySelectorAll('#locals-table .locals-toggle')]
    .find(x => x.getAttribute('aria-label')?.includes('sample')); button.click(); return true; })()`);
  await waitFor('Collection members collapse', async () =>
    !(await rows()).some(row => row.text.includes('Item 1')),
    { budgetMs: 10000 });
  if (!(await rows()).some(row => row.text.includes('sample') && row.toggle === 'false'))
    throw new Error('The Collection did not return to its collapsed state');
  const modelClicked = await page(`(() => { const button = [...document.querySelectorAll('#locals-table .locals-toggle')]
    .find(x => x.getAttribute('aria-label')?.includes('model')); button?.click(); return !!button; })()`);
  if (!modelClicked) throw new Error('No user class expander in the live Locals panel');
  await waitFor('user class members', async () =>
    (await rows()).some(row => row.depth === '14px' && row.text.includes('ready')),
    { budgetMs: 10000 });
  const classRows = await rows();
  if (!classRows.some(row => row.depth === '14px' && /Count\s*7\s*Long/.test(row.text)))
    throw new Error(`Class property Count was not shown: ${JSON.stringify(classRows)}`);
  await page(`(() => { const button = [...document.querySelectorAll('#locals-table .locals-toggle')]
    .find(x => x.getAttribute('aria-label')?.includes('model')); button.click(); return true; })()`);
  await waitFor('user class collapse', async () =>
    !(await rows()).some(row => row.depth === '14px' && row.text.includes('ready')),
    { budgetMs: 10000 });
  const interfaceClicked = await page(`(() => { const button = [...document.querySelectorAll('#locals-table .locals-toggle')]
    .find(x => x.getAttribute('aria-label')?.includes('iface')); button?.click(); return !!button; })()`);
  if (!interfaceClicked) throw new Error('No interface-typed object expander');
  await waitFor('interface property', async () =>
    (await rows()).some(row => row.depth === '14px' && row.text.includes('StatusreadyString')),
    { budgetMs: 10000 });
  await page(`(() => { const button = [...document.querySelectorAll('#locals-table .locals-toggle')]
    .find(x => x.getAttribute('aria-label')?.includes('iface')); button.click(); return true; })()`);
  await waitFor('interface collapse', async () =>
    !(await rows()).some(row => row.depth === '14px' && row.text.includes('StatusreadyString')),
    { budgetMs: 10000 });
  if ((await api.breakpoints()).mode !== 'break') throw new Error('Expansion left the debugger break');
  if (await immediateRows() !== immediateBefore) {
    const history = await page(`Array.from(document.querySelectorAll('#immediate-log > div'), row => row.textContent)`);
    throw new Error(`Locals expansion added output to the Immediate panel: ${JSON.stringify(history)}`);
  }
  console.log('PASS: live Locals panel expands and collapses nested Collections, class and interface properties.');
} finally {
  socket.close();
}
