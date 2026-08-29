/*
 * The agent card: the api's switch, and the address it hands out.
 *
 * WHAT ONLY THIS CAN ASK. The api used to be a dev-build-only door, so "is it off?" was answered
 * by the compiler and proved by grepping a Release binary. It ships now, which moves the question
 * to run time: the door has to be shut until someone opens it, it has to actually shut when they
 * close it, and the choice has to survive a restart. None of that is a compile-time fact any more.
 *
 * WHAT THIS CANNOT ASK, and where the gap is covered instead:
 *
 *   - That a RELEASE build leans shut. This suite runs against whatever build is open, and a dev
 *     build leans open. `verify.ps1` reads the published Release binary for the phrase the const
 *     folds to, which is the only place that fact is checkable.
 *   - That the door actually stops listening. Turning it off severs the connection this suite is
 *     driving over - which is the correct behaviour and makes it untestable FROM HERE. The check
 *     that the socket is gone is a PowerShell one; see `Test-ApiSwitch.ps1`.
 *
 * So this suite asks the half that can be asked without cutting its own line: that the card reads
 * the door honestly, that the text it hands over is THIS session's, and that it says what turning
 * the switch on costs.
 */

import { open, waitFor, comparingReporter } from "./xlide-api.mjs";
import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";

const api = await open();
const { check, done } = comparingReporter();

// The door as it stands, read from OUTSIDE the page: the discovery file is what a client finds.
const folder = join(process.env.LOCALAPPDATA, "xlide_vbide");
const discovered = readdirSync(folder)
  .filter((one) => one.startsWith("xlide-api-") && one.endsWith(".json"))
  .map((one) => JSON.parse(readFileSync(join(folder, one), "utf8")))
  .find((one) => one.pid === (api.pid ?? one.pid));

check("the session advertises itself in a discovery file", discovered !== undefined);

// ---- the card opens from its own toolbar button ------------------------------------------

const pressed = await api.act("toolbar", { command: "openAgent" });
check("the robot button opens the agent card", pressed.did, true);

await waitFor("the card", async () => (await api.ui()).agent !== null, { budgetMs: 15000 });
await waitFor("the door's state to arrive", async () =>
  (await api.ui()).agent?.busy === false, { budgetMs: 15000 });

const card = (await api.ui()).agent;
check("and it reads the door as open, which it must be for this suite to be talking to it",
  card.api, true);

// ---- the text is THIS session's ------------------------------------------------------------
//
// A generic instruction set would send an agent hunting for a discovery file and guessing which
// of several running editors it had found. Every one of these is a fact about this process.

const text = card.text;
check("the paste carries this session's port", text.includes(String(discovered.port)));
check("and this session's token", text.includes(discovered.token));
check("and this session's process", text.includes(String(discovered.pid)));
check("and this session's host application", text.includes(discovered.host));
check("and the address to start at", text.includes(discovered.agent));

// The in-process door is a second address onto the same capability, so it is named too.
check("and the in-process address", text.includes('GetObject(, "Xlide.Api")'));

// ---- and it says what it costs ---------------------------------------------------------------
//
// The switch opens a door into the developer's projects. A card that handed over an address
// without saying that would be the product being coy about the one thing that matters.

const said = await api.ask(`JSON.stringify({
  consequence: document.getElementById('agent-consequence')?.textContent ?? '',
  lit: document.getElementById('agent-lit')?.textContent ?? '',
  label: document.getElementById('agent-toggle-label')?.textContent ?? '',
  described: document.getElementById('agent-toggle')?.getAttribute('aria-describedby') ?? '',
  labelled: document.querySelector('label[for="agent-text"]') !== null,
  readOnly: document.getElementById('agent-text')?.readOnly ?? false,
})`);
const shown = JSON.parse(typeof said === "string" ? said : JSON.stringify(said));

check("the card says the door is on, in a word and not only a colour", shown.lit, "on");
check("and warns that a local program with the token can write and run code",
  /write and run code/.test(shown.consequence), true);
check("and that the choice outlives the session",
  /after a restart/.test(shown.consequence), true);
check("the switch is described by that warning, for a reader who cannot see the layout",
  shown.described, "agent-consequence");
check("the paste box has a real label, and is not editable by accident",
  { labelled: shown.labelled, readOnly: shown.readOnly }, { labelled: true, readOnly: true });

// ---- and it can be read without scrolling sideways -------------------------------------------
//
// The box exists to be selected and copied, so a horizontal scrollbar under it means reading what
// you are about to hand over takes two gestures. It wraps instead - and the wrap is DISPLAY only:
// the newlines are what the paste depends on, so they have to survive it.

const shape = await api.ask(`JSON.stringify((() => {
  const box = document.getElementById('agent-text');
  return {
    sideways: box.scrollWidth > box.clientWidth,
    wraps: getComputedStyle(box).whiteSpace,
  };
})())`);
const fits = JSON.parse(typeof shape === "string" ? shape : JSON.stringify(shape));

check("the paste does not scroll sideways", fits.sideways, false);
check("and wraps rather than clipping", fits.wraps, "pre-wrap");

// The longest line is the workbook's full path, which has no space to fold on - so this is the
// one that would have kept the scrollbar had the box only been told to wrap on words.
check("the text itself keeps its lines through the wrap", text.split("\n").length > 10, true);
check("and still carries an address to start at",
  /http:\/\/127\.0\.0\.1:\d+\/[0-9a-f]+/.test(text), true);

// ---- copying ---------------------------------------------------------------------------------

const copied = await api.act("agentCard", { press: "copy" });
check("the copy button answers", copied.did, true);
await waitFor("the copy to land", async () =>
  (await api.ui()).agent?.copied === true, { budgetMs: 8000 });
check("and reports that it reached the clipboard", (await api.ui()).agent.copied, true);

// ---- and it closes ---------------------------------------------------------------------------

const shut = await api.act("agentCard", { press: "close" });
check("the card closes", { did: shut.did, gone: (await api.ui()).agent }, { did: true, gone: null });

const missing = await api.act("agentCard", { press: "copy" });
check("and once closed it declines rather than pretending", missing.did, false);

process.exit(done());
