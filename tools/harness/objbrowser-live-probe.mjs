// Drives the LIVE editor and Object Browser pages semantically, through the dev build's
// two doors: the browser's DevTools protocol for real events on real elements, and the
// shim's debug API for the native truth (what module is shown, whether the palette is
// open). No pixel coordinates, no posted mouse messages - which is what finally lets this
// probe double-click a member and pin the navigate leg that synthetic input never could.
//
// Usage: node objbrowser-live-probe.mjs --api http://127.0.0.1:PORT/TOKEN [--cdp 9333]
// Prints a JSON verdict {pass, checks} on stdout; exits nonzero when any check fails.
// Invoked by Test-ObjectBrowser.ps1 against the Excel it launched.

const args = process.argv.slice(2);
const apiBase = args[args.indexOf("--api") + 1];
const cdpPort = args.includes("--cdp") ? Number(args[args.indexOf("--cdp") + 1]) : 9333;

const checks = [];
const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });
const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));

async function api(route) {
  const reply = await fetch(`${apiBase}/${route}`, { method: route.startsWith("command") ? "POST" : "GET" });
  return reply.json();
}

function connect(wsUrl) {
  const socket = new WebSocket(wsUrl);
  let nextId = 0;
  const waiting = new Map();

  socket.onmessage = (event) => {
    const message = JSON.parse(event.data);
    if (message.id && waiting.has(message.id)) {
      const { settle, reject } = waiting.get(message.id);
      waiting.delete(message.id);
      if (message.error) {
        reject(new Error(message.error.message));
      } else {
        settle(message.result);
      }
    }
  };

  const send = (method, params = {}, sessionId) =>
    new Promise((settle, reject) => {
      const id = ++nextId;
      waiting.set(id, { settle, reject });
      socket.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
    });

  return new Promise((settle, reject) => {
    socket.onopen = () => settle({ socket, send });
    socket.onerror = () => reject(new Error(`could not connect to ${wsUrl}`));
  });
}

async function attachToPage(send, titled) {
  const targets = await (await fetch(`http://127.0.0.1:${cdpPort}/json`)).json();
  const target = targets.find((one) => one.type === "page" && one.title === titled);
  if (!target) {
    throw new Error(`no page target titled ${titled}; saw ${targets.map((one) => one.title).join("|")}`);
  }

  const { sessionId } = await send("Target.attachToTarget", { targetId: target.id, flatten: true });
  return sessionId;
}

async function evaluate(send, sessionId, expression) {
  const reply = await send(
    "Runtime.evaluate",
    { expression, awaitPromise: true, returnByValue: true },
    sessionId,
  );
  if (reply.exceptionDetails) {
    throw new Error(reply.exceptionDetails.exception?.description ?? "evaluate failed");
  }
  return reply.result?.value;
}

const overall = setTimeout(() => {
  console.error("probe timed out after 90s");
  process.exit(2);
}, 90000);
overall.unref();

try {
  if (!apiBase) {
    throw new Error("--api http://127.0.0.1:PORT/TOKEN is required");
  }

  // The native truth before anything is driven: surface up, palette absent.
  let state = await api("state");
  check("the debug api answers with a ready surface", state.surfaceReady === true);
  check("no palette exists before the summons", state.paletteOpen === false);

  // One socket serves every target in the shared browser cluster.
  const versionReply = await (await fetch(`http://127.0.0.1:${cdpPort}/json/version`)).json();
  const { socket, send } = await connect(versionReply.webSocketDebuggerUrl);

  // Summon the Browser the way a person does: the toolbar button, clicked as an element.
  // Waited for, not assumed: the shim's surface-ready flag precedes the page finishing
  // its own boot, and a query fired into that gap found no toolbar at all.
  const editor = await attachToPage(send, "xlide editor");
  const clicked = await evaluate(send, editor, `(async () => {
    const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
    for (let waited = 0; waited < 15000; waited += 250) {
      const button = document.querySelector('[data-command="objectBrowser"]');
      if (button) { button.click(); return 'clicked'; }
      await sleep(250);
    }
    return 'no button';
  })()`);
  check("the toolbar button exists and takes a semantic click", clicked === "clicked", clicked);

  for (let waited = 0; waited < 10000; waited += 250) {
    state = await api("state");
    if (state.paletteOpen && state.paletteVisible) { break; }
    await sleep(250);
  }
  check("the click summons the palette", state.paletteOpen === true && state.paletteVisible === true);

  // Drive the palette itself: real libraries, real modules, real members from real code.
  await sleep(1000);
  const palette = await attachToPage(send, "Object Browser");
  const driven = await evaluate(send, palette, `(async () => {
    const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
    const rows = (pane) => [...document.querySelectorAll('#objbrowser-' + pane + ' .objbrowser-row')];
    const names = (pane) => rows(pane).map((row) => row.querySelector('.objbrowser-name').textContent);
    const picker = document.getElementById('objbrowser-library');

    for (let waited = 0; picker.options.length < 2 && waited < 10000; waited += 200) await sleep(200);
    const libraries = [...picker.options].map((option) => option.textContent);

    picker.selectedIndex = 0;
    picker.dispatchEvent(new Event('change', { bubbles: true }));
    for (let waited = 0; rows('modules').length < 2 && waited < 10000; waited += 200) await sleep(200);
    const modules = names('modules');

    const target = rows('modules').find((row) => row.querySelector('.objbrowser-name').textContent === 'CleanModule');
    if (!target) { return { libraries, modules, failed: 'no CleanModule row' }; }
    target.click();
    for (let waited = 0; rows('members').length < 1 && waited < 10000; waited += 200) await sleep(200);
    const members = names('members');

    const member = rows('members')[0];
    const picked = member.querySelector('.objbrowser-name').textContent;
    member.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true }));
    return { libraries, modules, members, picked };
  })()`);

  check("the picker lists the project first and real libraries after",
    Array.isArray(driven.libraries) && driven.libraries.length >= 2
      && driven.libraries[0].includes("(project)"),
    (driven.libraries ?? []).join("|"));
  check("the project scope lists the workbook's real modules",
    Array.isArray(driven.modules) && driven.modules.includes("CleanModule"),
    (driven.modules ?? []).join(","));
  check("a module's members come from its own code",
    Array.isArray(driven.members) && driven.members.length >= 1 && !driven.failed,
    driven.failed ?? (driven.members ?? []).join(","));

  // The pinned leg: the double-click reaches the editor and the shown module follows.
  let navigated = false;
  for (let waited = 0; waited < 10000; waited += 250) {
    state = await api("state");
    if (state.shownModule === "CleanModule") { navigated = true; break; }
    await sleep(250);
  }
  check("double-clicking a member navigates the editor to its module",
    navigated, `shownModule ${state.shownModule}`);

  // The native Browser stays retired: no type-2 window is visible.
  const windows = await api("windows");
  check("the native Object Browser window stays hidden",
    windows.windows.every((row) => row.type !== 2 || row.visible === false));

  socket.close();
  console.log(JSON.stringify({ pass: checks.every((one) => one.ok), checks }));
  process.exit(checks.every((one) => one.ok) ? 0 : 1);
} catch (failure) {
  checks.push({ name: "probe ran", ok: false, detail: String(failure.message ?? failure) });
  console.log(JSON.stringify({ pass: false, checks }));
  process.exit(1);
}
