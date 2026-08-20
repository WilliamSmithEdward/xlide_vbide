/*
 * The VBA test runner, driven end to end through the door: install the support module, write
 * a module of tests covering every outcome the runner can answer, run them, and hold each
 * status to what the directives promised. The pane is fed by the same snapshot this route
 * answers, so what passes here is what the panel shows.
 *
 *   node tools\harness\test-runner.mjs
 */
import { open, reporter, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const { check, done } = reporter();
const name = `TestProbe${process.pid % 10000}`;

const MODULE = [
  "' @xlide-test",
  "Public Sub Adds()",
  "    XlideAssert.AreEqual 4, 2 + 2",
  "End Sub",
  "",
  "' A directive block can carry metadata, and an aside after -- is never metadata.",
  "' @xlide-test tags=\"smoke,math\" owner=william timeout=5s -- flaky on Tuesdays",
  "Public Sub FailsOnPurpose()",
  "    XlideAssert.AreEqual 5, 2 + 2, \"arithmetic drifted\"",
  "    XlideAssert.WriteLine \"still ran to the end\"",
  "End Sub",
  "",
  "' @xlide-test",
  "Public Sub Raises()",
  "    Err.Raise 13",
  "End Sub",
  "",
  "' @xlide-test-skip reason=\"not on this machine\"",
  "Public Sub SkipsAlways()",
  "    XlideAssert.Fail \"never runs\"",
  "End Sub",
  "",
  "' @xlide-test-xfail reason=\"known defect 12\"",
  "Public Sub KnownBroken()",
  "    XlideAssert.IsTrue False",
  "End Sub",
  "",
  "' @xlide-test-xfail reason=\"was fixed and nobody noticed\"",
  "Public Sub SecretlyFixed()",
  "    XlideAssert.IsTrue True",
  "End Sub",
  "",
  "' @xlide-test expected-error=13",
  "Public Sub WantsTypeMismatch()",
  "    Err.Raise 13",
  "End Sub",
  "",
  "' @xlide-test expected-error=13",
  "Public Sub WantsButGetsAnother()",
  "    Err.Raise 9",
  "End Sub",
  "",
  "' @xlide-test expected-error",
  "Public Sub AnyErrorWillDo()",
  "    Err.Raise 76",
  "End Sub",
  "",
  "' @xlide-test",
  "Public Sub ThrowsHelper()",
  "    XlideAssert.Throws 13, \"MustMismatch\"",
  "End Sub",
  "",
  "Public Sub MustMismatch()",
  "    Err.Raise 13",
  "End Sub",
  "",
  "' @xlide-test",
  "Private Sub NotDiscovered()",
  "End Sub",
  "",
  "' @xlide-tost",
  "Public Sub TypoedDirective()",
  "End Sub",
].join("\r\n");

try {
  // ---- the support module's install gate ----
  const installed = await api.tests({ action: "install" });
  check("installing XlideAssert answers installed", installed.support === "installed", installed.support);

  const refused = await api.tests({ action: "run" }).catch((error) => ({ error: String(error) }));
  check("a run with no tests still answers rather than failing",
    refused.detail !== undefined || refused.error !== undefined);

  // ---- discovery ----
  await api.component("add", { kind: "module", name, project: project.projectId });
  await api.writeModule(name, MODULE, project.projectId);
  const listed = await waitFor("discovery to see the module", async () => {
    const now = await api.tests();
    return now.rows.filter((row) => row.module === name).length >= 10 ? now : null;
  }, { budgetMs: 15000 });
  const mine = listed.rows.filter((row) => row.module === name);

  check("discovery finds every directive-marked public Sub", mine.length === 10,
    mine.map((row) => row.procedure).join(", "));
  check("a Private Sub under a directive is not offered - the runner could never call it",
    !mine.some((row) => row.procedure === "NotDiscovered"));
  check("a typoed directive discovers nothing", !mine.some((row) => row.procedure === "TypoedDirective"));
  check("an unmarked helper Sub is not a test", !mine.some((row) => row.procedure === "MustMismatch"));

  // AUTO-REDISCOVERY: the pane heard about this module from the analysis pass alone - the
  // write above provoked a pass, the pass handed its snapshot over, and the rows repainted
  // with no tests() call and no refresh press. Read off the pane's own DOM, which paints
  // whether or not the pane is showing.
  const autoHeard = await waitFor("the pane to hear the new tests on its own", async () =>
    (await api.ask(`document.querySelector("#tests-list")?.textContent.includes(${JSON.stringify(name)}) === true`)) === true,
  { budgetMs: 20000 }).then(() => true).catch(() => false);
  check("the pane rediscovers on its own when code changes - no refresh pressed", autoHeard);

  const tagged = mine.find((row) => row.procedure === "FailsOnPurpose");
  check("metadata rides discovery: tags, owner, and a timeout in seconds",
    tagged?.tags.join(",") === "smoke,math" && tagged?.owner === "william" && tagged?.timeoutMs === 5000,
    JSON.stringify({ tags: tagged?.tags, owner: tagged?.owner, timeoutMs: tagged?.timeoutMs }));
  check("a declaration line is where the row navigates to",
    (mine.find((row) => row.procedure === "Adds")?.line ?? 0) === 2,
    `Adds is at line ${mine.find((row) => row.procedure === "Adds")?.line}`);

  // ---- the run, every outcome at once ----
  const ran = await api.tests({ action: "run", module: name, timeoutMs: 180000 });
  const status = (procedure) => ran.rows.find((row) => row.module === name && row.procedure === procedure);

  check("a passing test passes", status("Adds")?.status === "passed", JSON.stringify(status("Adds")));
  check("a failed assertion fails with the latched message, user text first",
    status("FailsOnPurpose")?.status === "failed"
    && /arithmetic drifted Expected <5> but was <4>\./.test(status("FailsOnPurpose")?.message ?? ""),
    status("FailsOnPurpose")?.message);
  check("...and the WriteLine after the failed assertion still arrived - assertions are not fatal",
    (status("FailsOnPurpose")?.output ?? []).includes("still ran to the end"),
    JSON.stringify(status("FailsOnPurpose")?.output));
  check("a raised error fails with its number and source",
    status("Raises")?.status === "failed" && /VBA error 13/.test(status("Raises")?.message ?? ""),
    status("Raises")?.message);
  check("a skip directive skips without running - its Fail never latched",
    status("SkipsAlways")?.status === "skipped" && status("SkipsAlways")?.message === "not on this machine",
    JSON.stringify(status("SkipsAlways")));
  check("an expected failure reads xfail, not red",
    status("KnownBroken")?.status === "xfail", status("KnownBroken")?.status);
  check("an expected failure that PASSES reads xpass, which is a red worth seeing",
    status("SecretlyFixed")?.status === "xpass"
    && /Expected failure did not occur/.test(status("SecretlyFixed")?.message ?? ""),
    JSON.stringify(status("SecretlyFixed")));
  check("expected-error=13 passes when 13 is raised",
    status("WantsTypeMismatch")?.status === "passed", status("WantsTypeMismatch")?.status);
  check("expected-error=13 fails when 9 arrives, naming both",
    status("WantsButGetsAnother")?.status === "failed"
    && /Expected VBA error 13, but got VBA error 9/.test(status("WantsButGetsAnother")?.message ?? ""),
    status("WantsButGetsAnother")?.message);
  check("a bare expected-error takes any error at all",
    status("AnyErrorWillDo")?.status === "passed", status("AnyErrorWillDo")?.status);
  check("Assert.Throws reaches its target through the staged dispatcher",
    status("ThrowsHelper")?.status === "passed", JSON.stringify(status("ThrowsHelper")));

  // ---- one test alone, and the failed set ----
  const one = await api.tests({ action: "run", test: `${name}.Adds`, timeoutMs: 60000 });
  // The answer says how many ran and in which file, because a session holds more than one.
  check("run with test= runs exactly that test",
    one.rows.find((row) => row.id === `${name}.Adds`)?.status === "passed"
    && /^ran 1 in /.test(one.detail), one.detail);

  const failedAgain = await api.tests({ action: "runFailed", timeoutMs: 120000 });
  const rerunIds = [];
  for (const row of failedAgain.rows) {
    if (row.module === name && row.status !== "none" && row.durationMs > 0
      && ["failed", "error", "xpass"].includes(row.status)) {
      rerunIds.push(row.procedure);
    }
  }

  check("runFailed reruns the failing set and it fails the same way",
    failedAgain.rows.find((row) => row.id === `${name}.FailsOnPurpose`)?.status === "failed"
    && failedAgain.rows.find((row) => row.id === `${name}.SecretlyFixed`)?.status === "xpass");

  // ---- the injected modules leave no trace ----
  const components = (await api.tests()).rows;
  check("the generated runner and dispatcher are gone when the run ends",
    !components.some((row) => row.module.startsWith("XlideRun") || row.module === "XlideTestDispatch"));
  const support = await api.tests();
  check("XlideAssert still reads installed after runs - injection re-cases nothing",
    support.support === "installed", support.support);
} finally {
  await api.component("remove", { name, project: project.projectId }).catch(() => {});
  process.exitCode = done();
}
