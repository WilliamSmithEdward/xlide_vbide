/*
 * What the editor says about a workbook Excel will not run. Driven by Test-MacrosDisabled.ps1,
 * which is what opens the workbook that way; on its own this measures an ordinary session.
 *
 * Excel puts itself into design mode for such a workbook and will not come out of it. Every
 * command aimed at the project greys, the Design Mode toggle stays pressed however often it is
 * pressed, and nothing the product offers clears it - the way out is closing the workbook and
 * opening it again with macros enabled. That is issue #9's shape, and the point of this probe is
 * that the product SAYS so rather than answering "currently disabled" and leaving it there.
 */

import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open({});
const { check, done } = reporter();

const state = await api.state();
console.log(`  pid ${api.pid}, showing ${state.shownProject}/${state.shownModule}, `
  + `published ${state.debugMode}`);

/* ---- the state itself ------------------------------------------------------------------------ */

const toggle = await api.bars("designMode");
check("Excel put ITSELF into design mode, with nobody pressing anything",
  toggle.places.every((one) => one.state === -1),
  JSON.stringify(toggle.places.map((one) => one.state)));

const reset = await api.bars("reset");
check("so every copy of reset is greyed",
  reset.enabledCount === 0, `${reset.enabledCount} of ${reset.places.length}`);

check("while the project itself still reports design mode",
  reset.mode === 2 && reset.modeError === null,
  `mode ${reset.mode}, publishedMode ${reset.publishedMode}, modeError ${reset.modeError}`);

// The three readings that made this state unreadable from outside: nothing is standing, so a
// caller looking for a cause finds none.
const dialogs = await api.dialogs().catch(() => ({ dialogs: [] }));
const forms = await api.userforms().catch(() => ({ forms: [] }));
check("with no dialog standing and no form showing, which is why it read as nothing at all",
  (dialogs.dialogs ?? []).length === 0 && (forms.forms ?? forms.userforms ?? []).length === 0,
  `dialogs ${(dialogs.dialogs ?? []).length}, forms ${(forms.forms ?? forms.userforms ?? []).length}`);

/* ---- and whether the product says so ---------------------------------------------------------- */

// The first press cannot know - the toggle is read back a tick later - so it is the second that
// has to be honest. A button answering "executed" while nothing moves is the false success this
// door has been cleaning out all week.
const first = await api.command("designMode");
await wait(1500);
const afterFirst = await api.bars("designMode");
check("pressing Design Mode does not move it",
  afterFirst.places.every((one) => one.state === -1),
  `${JSON.stringify(afterFirst.places.map((one) => one.state))} after ${JSON.stringify(first.detail)}`);

const second = await api.command("designMode");
check("and the next press says so rather than reporting success",
  second.ran === false && second.detail.includes("macros are disabled"),
  `ran=${second.ran} ${second.detail}`);

const refused = await api.command("reset");
check("a refused reset names the real cause, not the toggle",
  refused.ran === false
    && refused.detail.includes("will not leave it")
    && refused.detail.includes("macros are disabled"),
  refused.detail);

check("and it still names every control it found",
  refused.detail.includes("all 6 control(s)"), refused.detail.slice(0, 80));

process.exit(done());
