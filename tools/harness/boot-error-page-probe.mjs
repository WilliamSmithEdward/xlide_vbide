// Serves the BUILT page with a bundle that throws on load, and checks that the page's own black
// box caught it.
//
// This is the guard for the gap that cost the most on 2026-08-09. The console ring used to be
// installed by the host once the page reported itself ready; a bundle that throws while its
// modules initialise never reaches ready, so the ring never existed and the `console` route
// answered {"installed": false, "lines": []} at the one moment somebody was asking why the screen
// was blank. The cause was found by reading source instead.
//
// boot.js now installs ahead of the bundle and both records and PUSHES the error. Headless is the
// right place to pin it: provoking it against a real Excel means publishing a broken bundle, and a
// gate that has to break the developer's editor to run is a gate nobody runs.
//
// Prints a JSON verdict {pass, checks} on stdout and exits nonzero when any check fails.


import { runPageProbe } from "./page-probe.mjs";

/** The fault, in the shape it actually took: a const read during its own dead zone. */
const THROWING_BUNDLE =
  'throw new ReferenceError("Cannot access \'BUILTIN_OBJECTS\' before initialization");\n';

const DRIVE = `(() => {
  const checks = [];
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });

  const ring = window.__xlideConsole;

  check('the ring exists even though the bundle never ran', Array.isArray(ring),
    ring === undefined ? 'undefined: boot.js did not run, or ran after the bundle' : typeof ring);

  const uncaught = (ring || []).filter((one) => one.indexOf('UNCAUGHT') === 0);
  check('it caught the throw', uncaught.length > 0, (ring || []).join(' | ') || 'the ring is empty');

  check('and it names the error, not just that there was one',
    uncaught.some((one) => one.indexOf('ReferenceError') >= 0 && one.indexOf('BUILTIN_OBJECTS') >= 0),
    uncaught.join(' | '));

  check('with a file and a line to go to',
    uncaught.some((one) => /editor\\.js:\\d+:\\d+/.test(one)), uncaught.join(' | '));

  // The bundle is the thing that failed, so nothing it installs can be relied on here. This is
  // what makes the check meaningful: the surface genuinely is not there.
  check('the surface really did not come up', typeof window.xlideUi === 'undefined',
    typeof window.xlideUi);

  return { pass: checks.every((one) => one.ok), checks };
})()`;

await runPageProbe({
  label: "xlide-boot-error",
  needs: "boot.js",
  serve: (asked) => (asked === "/editor.js" ? THROWING_BUNDLE : null),
  drive: DRIVE,
});
