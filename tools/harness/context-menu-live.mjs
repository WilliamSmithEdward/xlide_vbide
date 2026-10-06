// Focused live regression for #51. Run against a Debug Excel/VBE session:
//   node tools/harness/context-menu-live.mjs <Excel PID>
// Checks the actual Monaco menu inside its shadow root, rather than document CSS.
import { open } from './xlide-api.mjs';
import { loopbackUrl } from './loopback-url.mjs';

const pid = Number(process.argv[2] ?? process.env.XLIDE_PID);
if (!Number.isInteger(pid) || pid <= 0) throw new Error('Pass the disposable Excel PID.');
const api = await open({ pid });
if (!api.devtoolsPort) throw new Error('The Debug add-in must expose DevTools for color-scheme emulation.');

const targets = await (await fetch(loopbackUrl(api.devtoolsPort, '/json'))).json();
const target = targets.find((row) => row.title === 'xlide editor');
if (!target) throw new Error('The live editor DevTools target is unavailable.');
const socket = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
let nextId = 0;
const pending = new Map();
socket.onmessage = ({ data }) => {
  const reply = JSON.parse(data);
  const waiter = pending.get(reply.id);
  if (!waiter) return;
  pending.delete(reply.id);
  if (reply.error) waiter.reject(new Error(reply.error.message));
  else waiter.resolve(reply.result);
};
function cdp(method, params) {
  return new Promise((resolve, reject) => {
    const id = ++nextId;
    pending.set(id, { resolve, reject });
    socket.send(JSON.stringify({ id, method, params }));
  });
}

const inspect = `(() => {
  const line = document.querySelector('.monaco-editor .view-lines');
  if (!line) throw new Error('No live Monaco editor is open');
  const rect = line.getBoundingClientRect();
  line.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true,
    clientX: rect.x + 100, clientY: rect.y + 40, button: 2 }));
  const root = document.querySelector('.shadow-root-host')?.shadowRoot;
  const menu = root?.querySelector('.context-view.monaco-menu-container > .monaco-scrollable-element');
  const label = root?.querySelector('.monaco-menu .action-menu-item');
  return { visible: !!menu, background: menu && getComputedStyle(menu).backgroundColor,
    foreground: label && getComputedStyle(label).color };
})()`;

try {
  for (const [scheme, background, foreground] of [
    ['dark', 'rgb(37, 37, 38)', 'rgb(240, 240, 240)'],
    ['light', 'rgb(255, 255, 255)', 'rgb(27, 27, 31)'],
  ]) {
    await cdp('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-color-scheme', value: scheme }] });
    const found = await api.ask(inspect);
    if (!found?.visible || found.background !== background || found.foreground !== foreground) {
      throw new Error(`${scheme} context menu: ${JSON.stringify(found)}; expected ${background} and ${foreground}`);
    }
    console.log(`${scheme} context menu: opaque ${found.background}, readable ${found.foreground}`);
  }
  console.log('RESULT: PASS - live context menu in dark and light themes');
} finally {
  await cdp('Emulation.setEmulatedMedia', { features: [] });
  socket.close();
}
