// Regression suite for #21, through the same evaluation path as the Immediate panel.
import { open, wait, waitFor, reporter } from './xlide-api.mjs';

const pid = Number(process.argv[2] ?? process.env.XLIDE_PID);
if (!Number.isInteger(pid) || pid <= 0) throw new Error('Pass the isolated DebugFixture Excel PID.');
const api = await open({ pid });
const home = await api.projectHolding('Runner');
if (!home || !/DebugFixture/i.test(home.project)) throw new Error('Requires DebugFixture.xlsm.');
const project = await api.project(home.project);
const original = (await api.readModule('Runner', project.projectId)).text;
const lines = original.split(/\r?\n/);
const stopLine = lines.findIndex(line => line.includes('counter = counter + 1')) + 1;
const { check, done } = reporter();
const stopped = async () => (await api.breakpoints()).mode === 'break';

// Exercise the real input handler too: the API has its own historical modal rescue loop,
// so API-only checks cannot prove that a person typing an invalid expression can recover.
const targets = await (await fetch(`http://127.0.0.1:${api.devtoolsPort}/json`)).json();
const target = targets.find(target => target.title === 'xlide editor');
if (!target) throw new Error('The editor DevTools target is unavailable.');
const socket = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
let nextId = 0;
const pending = new Map();
socket.onmessage = event => {
  const message = JSON.parse(event.data);
  const reply = pending.get(message.id);
  if (!reply) return;
  pending.delete(message.id);
  if (message.error || message.result?.exceptionDetails) reply.reject(new Error(JSON.stringify(message)));
  else reply.resolve(message.result?.result?.value);
};
function page(expression) {
  return new Promise((resolve, reject) => {
    const id = ++nextId;
    pending.set(id, { resolve, reject });
    socket.send(JSON.stringify({ id, method: 'Runtime.evaluate',
      params: { expression, returnByValue: true, awaitPromise: true } }));
  });
}
const rows = () => page(`Array.from(document.querySelectorAll('#immediate-log > div'),
  node => ({text: node.textContent, kind: node.className}))`);
async function fromUi(expression, expected, failed = false) {
  await page("window.xlideBridge.runEditorCommand('xlide.panel.immediate')");
  const before = (await rows()).length;
  await page(`(() => {
    const input = document.querySelector('#immediate-input');
    input.value = ${JSON.stringify(expression)};
    input.focus();
    input.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', bubbles: true}));
  })()`);
  await waitFor('Immediate UI result', async () => (await rows()).length >= before + 2,
    { budgetMs: 10000 });
  await wait(400); // Allow the independent native-output poller to reveal any duplicate.
  const added = (await rows()).slice(before);
  check('UI produces one echo and exactly one result', added.length === 2, JSON.stringify(added));
  check('UI returns the expected result',
    failed ? added[1].kind.includes('failed') && added[1].text.includes(expected)
      : !added[1].kind.includes('failed') && added[1].text === expected, JSON.stringify(added));
  check('the Immediate prompt keeps keyboard focus', await page(`document.activeElement?.id === 'immediate-input'`));
  check('UI evaluation preserves the breakpoint', await stopped());
}

async function evaluate(expression, expected) {
  const reply = await api.immediate(expression);
  check(`${expression} returns ${JSON.stringify(expected)}`,
    reply.ran && !reply.failed && reply.text.trim() === expected, JSON.stringify(reply));
  check('the original debugger session remains paused', await stopped());
}

async function invalidExpression(expression) {
  const reply = await api.immediate(expression);
  check(`${expression} reports an error`, reply.ran && reply.failed && Boolean(reply.text), JSON.stringify(reply));
  check('no error dialog remains', (await api.dialogs()).dialogs.length === 0);
  check('an expression error preserves the paused session', await stopped());
}

try {
  await api.breakpoint('Runner', stopLine, { project: project.projectId, state: 'on' });
  await waitFor('installed breakpoint', async () => (await api.breakpoints()).breakpoints.some(
    point => point.module === 'Runner' && point.lines.includes(stopLine)), { budgetMs: 10000 });
  await api.caret(lines.findIndex(line => line.includes('Public Sub Walk')) + 1,
    { module: 'Runner', project: project.projectId });
  await wait(600); // Native pane activation settles independently of the API caret reply.
  await api.command('run');
  await waitFor('breakpoint', stopped, { budgetMs: 15000 });
  await evaluate('? counter + 1', '2');
  await evaluate('? UCase$(label) & CStr(counter)', 'START1');
  await evaluate('? ratio * 4', '2');
  await evaluate('? ratio * 4', '2');
  await evaluate('? ThisWorkbook.Name', 'DebugFixture.xlsm');
  await evaluate('? ""', '');
  await evaluate('? "first" & vbCrLf & "last"', 'first\r\nlast');
  await evaluate('Debug.Print label & ":" & CStr(counter)', 'start:1');
  const longOutput = await api.immediate('? String$(1500, "x")');
  check('a single result longer than 1024 characters is complete',
    !longOutput.failed && longOutput.text === 'x'.repeat(1023) + '\r\n' + 'x'.repeat(477),
    `length=${longOutput.text.length}, including native line wrapping`);
  // Reusing the native working buffer must not clear the visible Immediate history.
  const beforeHistory = (await rows()).length;
  for (let i = 0; i < 20; i++) await evaluate(`? "history-${i}"`, `history-${i}`);
  await waitFor('visible Immediate history', async () => (await rows()).length >= beforeHistory + 20,
    { budgetMs: 10000 });
  const history = (await rows()).slice(beforeHistory);
  check('all twenty results remain in the visible history exactly once', history.length === 20
    && history.every((row, i) => row.text === `history-${i}`));
  await fromUi('? "  padded  "', '  padded  ');
  await fromUi('Debug.Print "  printed  "', '  printed  ');
  await fromUi('Debug.Print ""', '');
  await fromUi('? 1/0', 'Division by zero', true);
  await fromUi('? ((', 'Expected: expression', true);
  await fromUi('? label & " again"', 'start again');
  await invalidExpression('? 1 / 0');
  await evaluate('? counter + 1', '2');
  await invalidExpression('? ((');
  await evaluate('? counter + 1', '2');
  await api.command('stepOver');
  await waitFor('next statement', async () => (await api.state()).caretLine > stopLine,
    { budgetMs: 10000 });
  await evaluate('? counter + 1', '3');
  check('the source is unchanged', (await api.readModule('Runner', project.projectId)).text === original);
  // Asked by name: the component list leaves the scratch module out on purpose, so reading it
  // here said "none inserted" whether or not one stood (see scratchModuleStands).
  check('no scratch module was inserted', !(await api.scratchModuleStands()));
  const beforeResume = (await rows()).length;
  await api.command('run');
  await waitFor('Walk finishes', async () => !(await stopped()), { budgetMs: 10000 });
  await waitFor('Debug.Print from running VBA', async () => (await rows()).length > beforeResume,
    { budgetMs: 10000 });
  await wait(400);
  const printed = (await rows()).slice(beforeResume);
  check('running VBA Debug.Print reaches the panel exactly once', printed.length === 1
    && printed[0].text.includes('2') && !printed[0].text.includes("'xlide:"), JSON.stringify(printed));

  // A PANEL LINE THAT STOPS AT THE BREAKPOINT. `Runner.Walk` typed in the panel runs through the
  // scratch procedure, so that line is still running, suspended beneath the stop, when the next
  // line arrives. The editor's own Immediate window evaluates it in the stopped scope; the panel
  // refused it as "already in progress" until the suspended line stopped holding it off.
  await page("window.xlideBridge.runEditorCommand('xlide.panel.immediate')");
  await page(`(() => {
    const input = document.querySelector('#immediate-input');
    input.value = 'Runner.Walk';
    input.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', bubbles: true}));
  })()`);
  await waitFor('the panel line stops at the breakpoint', stopped, { budgetMs: 15000 });
  check('the panel line stopped in the developer\'s own module',
    (await api.native()).activeModule === 'Runner');
  await fromUi('? counter', ' 1 ');
  await fromUi('? UCase$(label)', 'START');
  const beforeFinish = (await rows()).length;
  await api.command('run');
  await waitFor('the panel line finishes', async () => !(await stopped()), { budgetMs: 10000 });
  await waitFor('its output', async () => (await rows()).length > beforeFinish, { budgetMs: 10000 });
  await wait(400);
  const finished = (await rows()).slice(beforeFinish);
  check('Continue finishes the panel line, and its Debug.Print reaches the panel',
    finished.some(row => row.text.includes('value=2')), JSON.stringify(finished));
  const afterward = await api.immediate('? 6 * 7');
  check('the next line evaluates in design mode, so the finished line let go of the evaluator',
    afterward.ran && !afterward.failed && afterward.text.trim() === '42', JSON.stringify(afterward));
  check('and it left no scratch module standing', !(await api.scratchModuleStands()));
} finally {
  socket.close();
  await api.command('reset');
  await waitFor('design mode', async () => !(await stopped()), { budgetMs: 10000 });
  await api.breakpoint('Runner', stopLine, { project: project.projectId, state: 'off' });
}
process.exitCode = done();
