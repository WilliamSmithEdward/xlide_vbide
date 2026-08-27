/*
 * The analyzer rule catalog and the guarded severity overrides, through this engine.
 *
 * `analysis/rules` enumerates the analyzer actually bundled - the rules modal and the api both
 * render from it, so what they offer is exactly what this build can enforce. The overrides ride
 * `severityOverrides` on the diagnostics request, and the GUARD is upstream's: a warning or
 * information rule takes 'off'; an error rule takes at most 'warning', and only where marked
 * downgrade-safe. The engine IGNORES an illegal override rather than failing, which is why the
 * shim refuses them in words before they get here - and why this suite proves the ignoring, so
 * that refusal stays the only honest place.
 *
 *   node test/rule-overrides.mjs
 */

import { startEngine } from "./harness.mjs";

const { call, stop } = await startEngine("rule-overrides");
const CRLF = "\r\n";

let passed = 0;
const failures = [];
const check = (what, ok, detail) => {
  console.log(`  ${ok ? "ok  " : "FAIL"} ${what}${detail !== undefined ? `  -- ${String(detail).slice(0, 130)}` : ""}`);
  if (ok) { passed += 1; } else { failures.push(what); }
};

await call("initialize", {});

/* ---- the catalog ------------------------------------------------------------------------------ */

const { rules } = await call("analysis/rules", {});
check("the catalog is the analyzer's own, not a short list",
  rules.length >= 100, `${rules.length} rule(s)`);

const optionExplicit = rules.find((one) => one.code === "option-explicit-missing");
check("a style warning permits exactly 'off'",
  optionExplicit?.defaultSeverity === "warning"
    && optionExplicit?.allowed.length === 1 && optionExplicit?.allowed[0] === "off",
  JSON.stringify(optionExplicit));

const undeclared = rules.find((one) => one.code === "undeclared-variable");
check("a downgrade-safe error permits exactly 'warning'",
  undeclared?.defaultSeverity === "error"
    && undeclared?.allowed.length === 1 && undeclared?.allowed[0] === "warning",
  JSON.stringify(undeclared));

check("and the compile-equivalent rules permit nothing",
  rules.some((one) => one.defaultSeverity === "error" && one.allowed.length === 0),
  rules.filter((one) => one.allowed.length === 0).length + " fixed rule(s)");

/* ---- the overrides, against live diagnostics -------------------------------------------------- */

const diagnose = async (id, generation, source, severityOverrides) => {
  await call("project/open", {
    projectId: id, generation,
    modules: [{ moduleName: "M", source, type: "standard" }],
  });
  const answer = await call("textDocument/diagnostics", {
    documentKey: `${id}/M`, projectId: id, generation,
    source, moduleName: "M", moduleType: "standard",
    ...(severityOverrides === undefined ? {} : { severityOverrides }),
  });
  return answer.diagnostics ?? [];
};

// No Option Explicit, so the style rule fires; nothing else wrong.
const styleSource = ["Public Sub Go()", "    Dim n As Long", "    n = 1", "End Sub"].join(CRLF);

const plain = await diagnose("style", 1, styleSource);
check("the baseline finding is there to override",
  plain.some((d) => d.code === "option-explicit-missing"),
  plain.map((d) => d.code).join(",") || "(none)");

const silenced = await diagnose("style", 2, styleSource, { "option-explicit-missing": "off" });
check("'off' silences it machine-wide",
  !silenced.some((d) => d.code === "option-explicit-missing"),
  silenced.map((d) => d.code).join(",") || "(none)");

// WITH Option Explicit, so the undeclared assignment is a real error finding.
const errorSource = ["Option Explicit", "", "Public Sub Go()", "    missing1 = 1", "End Sub"].join(CRLF);

const asError = await diagnose("err", 1, errorSource);
const baseline = asError.find((d) => d.code === "undeclared-variable");
check("the error finding is there, as an error",
  baseline?.severity === "error", JSON.stringify(baseline?.severity));

const downgraded = await diagnose("err", 2, errorSource, { "undeclared-variable": "warning" });
check("the permitted downgrade re-severities it, without hiding it",
  downgraded.find((d) => d.code === "undeclared-variable")?.severity === "warning",
  JSON.stringify(downgraded.find((d) => d.code === "undeclared-variable")?.severity));

const illegal = await diagnose("err", 3, errorSource, { "undeclared-variable": "off" });
check("an ILLEGAL 'off' on an error rule is ignored - the finding stays",
  illegal.find((d) => d.code === "undeclared-variable")?.severity === "error",
  illegal.map((d) => `${d.code}:${d.severity}`).join(",") || "(none)");

await stop();
console.log(`\n${passed} passed, ${failures.length} failed`);
if (failures.length > 0) { failures.forEach((one) => console.log(`  ${one}`)); }
process.exit(failures.length > 0 ? 1 : 0);
