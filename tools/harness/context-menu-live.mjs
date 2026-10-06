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
  const frame = root?.querySelector('.monaco-menu');
  const label = root?.querySelector('.monaco-menu .action-menu-item');
  const labelRect = label?.getBoundingClientRect();
  return { visible: !!menu, background: menu && getComputedStyle(menu).backgroundColor,
    foreground: label && getComputedStyle(label).color,
    borderWidth: frame && getComputedStyle(frame).borderTopWidth,
    borderColor: frame && getComputedStyle(frame).borderTopColor,
    hoverAt: labelRect && { x: labelRect.x + 20, y: labelRect.y + labelRect.height / 2 } };
})()`;
const inspectHover = `(() => {
  const root = document.querySelector('.shadow-root-host')?.shadowRoot;
  const item = root?.querySelector('.action-item.focused .action-menu-item');
  const style = item && getComputedStyle(item);
  return { selected: !!item, background: style?.backgroundColor, foreground: style?.color,
    outlineWidth: style?.outlineWidth, outlineColor: style?.outlineColor };
})()`;

try {
  for (const [scheme, background, foreground, border, hoverBackground, hoverForeground, hoverOutline] of [
    ['dark', 'rgb(37, 37, 38)', 'rgb(240, 240, 240)',
      'rgb(107, 116, 125)', 'rgb(23, 109, 165)', 'rgb(255, 255, 255)', 'rgb(138, 200, 245)'],
    ['light', 'rgb(255, 255, 255)', 'rgb(27, 27, 31)',
      'rgb(133, 140, 150)', 'rgb(145, 200, 243)', 'rgb(27, 27, 31)', 'rgb(49, 120, 184)'],
  ]) {
    await cdp('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-color-scheme', value: scheme }] });
    const found = await api.ask(inspect);
    if (!found?.visible || found.background !== background || found.foreground !== foreground
      || found.borderWidth !== '1px' || found.borderColor !== border) {
      throw new Error(`${scheme} context menu surface: ${JSON.stringify(found)}`);
    }
    // Move out first so a previous scheme's pointer position cannot keep the same row hovered.
    await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: 20, y: 20, button: 'none' });
    await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', ...found.hoverAt, button: 'none' });
    const hovered = await api.ask(inspectHover);
    if (!hovered?.selected || hovered.background !== hoverBackground
      || hovered.foreground !== hoverForeground || hovered.outlineWidth !== '1px'
      || hovered.outlineColor !== hoverOutline) {
      throw new Error(`${scheme} context menu hover: ${JSON.stringify(hovered)}`);
    }
    console.log(`${scheme} context menu: opaque ${found.background}, ${found.borderWidth} border,`
      + ` hover ${hovered.background} with ${hovered.outlineColor} outline`);
  }
  console.log('RESULT: PASS - live context menu surface, border, and hover in both themes');
} finally {
  await cdp('Emulation.setEmulatedMedia', { features: [] });
  const closed = new Promise((resolve) => { socket.onclose = resolve; });
  socket.close();
  await closed;
}
