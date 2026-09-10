/*
 * ACCESS, WHICH IS NOT EXCEL WITH A DIFFERENT ICON, against a real one.
 *
 * Every other live suite runs in Excel, and for a year that meant the three places this product
 * asks which host it is in were only ever answered one way. The test runner had therefore never
 * once run a test in Access and nobody knew (2026-09-06). This suite is the other answer, and
 * every check in it corresponds to something that was actually broken:
 *
 *   - the host could not be REACHED at all: the walk that finds a host's document window knew
 *     Excel's worksheet pane and Word's document pane and answered zero for Access, so the
 *     Immediate window and the test runner both said "the host application could not be reached";
 *   - `Application.Run` is spelled differently here - a bare procedure name, where Excel wants
 *     `'file'!Module.Proc` and Word wants `Module.Proc`, and Access refuses both;
 *   - Access writes `Option Compare Database` into every module it creates, which made the
 *     support module it had just installed compare unequal to the canonical source, so the pane
 *     called its own module outdated for ever;
 *   - and a generated module that declared `Option Compare Text` then carried TWO of them, which
 *     does not compile, and a VBA project that does not compile resolves NOTHING - so the run
 *     was told "cannot find the procedure" about every name in the database.
 *
 * Runs against AccessFixture.accdb alone:
 *
 *   tools\harness\Start-Access.ps1 -Database artifacts\fixtures\AccessFixture.accdb -Fresh
 *   node tools\harness\access.mjs
 *
 * It puts back everything it changes, so a rerun starts where the first run did.
 */
import { mkdirSync, realpathSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { open, wait, waitFor, reporter } from "./xlide-api.mjs";

const api = await open();
const { check, done } = reporter();

const BOOK = "AccessFixture.accdb";
const at = { project: BOOK };

const refusalOf = async (promise) => {
  try {
    await promise;
    return null;
  } catch (error) {
    return String(error.message ?? error);
  }
};

const scaffolding = (components) =>
  components.filter((one) => /^XlideRun|^XlideTestDispatch$/i.test(one.name)).map((one) => one.name);

let helpersText = null;

try {
  // ---- the project the tree reads ----
  const project = await api.project();
  const names = project.components.map((one) => one.name).sort().join(",");
  check("the database's modules are all there, and all standard - Access has no UserForms",
    names === "Helpers,Pricing,PricingTests,XlideAssert"
      && project.components.every((one) => one.kind === "module"),
    names);

  // ---- the host is reachable, which is what the Immediate window proves ----
  // Both of these answered "The host application could not be reached" until the window walk
  // learned Access's OMain frame, so they are the check that the host is reachable AT ALL.
  const arithmetic = await api.immediate("?1+1");
  check("the Immediate window evaluates in Access", arithmetic.ran && arithmetic.text === "2" && !arithmetic.failed,
    JSON.stringify(arithmetic));

  const intoTheProject = await api.immediate("?Pricing.DiscountRate(10)");
  check("and evaluates against the database's own code", intoTheProject.text === "0.1", JSON.stringify(intoTheProject));

  const greeting = await api.immediate("?Greeting()");
  check("a bare procedure name resolves, which is the spelling Access's Run takes",
    greeting.text === "hello from access", JSON.stringify(greeting));

  // ---- the support module, and the line Access adds to it ----
  const listed = await api.tests({ action: "refresh" });
  check("XlideAssert reads INSTALLED, not outdated: Access writes its own Option Compare into "
    + "every module it creates, and comparing that verbatim called the module an edited copy",
    listed.support === "installed", listed.support);
  check("all four tests are discovered", listed.rows.length === 4,
    listed.rows.map((one) => one.procedure).join(","));

  // ---- the run itself ----
  const ran = await api.tests({ action: "run" });
  const outcome = (procedure) => ran.rows.find((one) => one.procedure === procedure);
  check("the run reaches the host and reports per test", /^ran 4 in/.test(ran.detail), ran.detail);
  check("the three that should pass, pass",
    ["NoDiscountUnderFive", "FivePercentFromFive", "ExpressAddsItsFee"].every((one) => outcome(one)?.status === "passed"),
    ran.rows.map((one) => one.procedure + "=" + one.status).join(", "));
  check("and the one that fails on purpose fails, in the assertion's own words",
    outcome("FailsOnPurpose")?.status === "failed"
      && /Expected True but was False/i.test(outcome("FailsOnPurpose")?.message ?? ""),
    outcome("FailsOnPurpose")?.message);

  // THE SCAFFOLDING GOES AGAIN, and in Access that matters more than anywhere else: a generated
  // module saved into the database and then removed leaves Access holding a catalogue entry for
  // something that is gone, and the next open of the file raises "File not found" from VBA
  // before anything can run (lessons, finding 72).
  const after = await api.project();
  check("the generated run modules are gone afterwards", scaffolding(after.components).length === 0,
    scaffolding(after.components).join(",") || "none");

  const rerun = await api.tests({ action: "run" });
  check("and a second run is as good as the first, so nothing was left half-standing",
    rerun.rows.filter((one) => one.status === "passed").length === 3,
    rerun.rows.map((one) => one.procedure + "=" + one.status).join(", "));

  // ---- what Access does not have ----
  const onForm = await refusalOf(api.component("add", { kind: "form", name: "XlideNotAForm" }));
  check("a UserForm is refused in words rather than relayed as a COM error - Access VBA has none",
    onForm !== null && /userform|form/i.test(onForm), onForm);
  check("and nothing was added by the attempt",
    (await api.project()).components.every((one) => one.name !== "XlideNotAForm"));

  // ---- a module round trip through the door ----
  helpersText = (await api.readModule("Helpers", BOOK)).text ?? "";
  const edited = helpersText.replace("hello from access", "hello again");
  await api.writeModule("Helpers", edited, BOOK, { by: "access.mjs" });
  const readBack = (await api.readModule("Helpers", BOOK)).text ?? "";
  check("a module written through the door reads back as it was written",
    readBack.includes("hello again"), readBack.split(/\r?\n/).length + " lines");

  const afterWrite = await api.immediate("?Greeting()");
  check("and the host runs the new text, not the old",
    afterWrite.text === "hello again", JSON.stringify(afterWrite));

  await api.writeModule("Helpers", helpersText, BOOK, { by: "access.mjs" });
  helpersText = null;
  check("and it goes back", (await api.immediate("?Greeting()")).text === "hello from access");

  // ---- source control, whose save and dirty checks are the host's ----
  //
  // Excel answers them from Workbooks and Word from Documents; Access has neither, a database
  // being no document. The project's own Saved flag says whether its modules need saving and
  // the editor's own Save is what saves them. Before this, Access answered "dirty: false" for
  // ever, and a checkout imported over an unsaved edit that Excel refuses to touch (2026-09-09).
  const scratch = join(tmpdir(), `xlide-access-scm-${process.pid}`);
  rmSync(scratch, { recursive: true, force: true });
  mkdirSync(scratch, { recursive: true });
  const folder = realpathSync.native(scratch);
  const scm = (args = {}) => api.scm({ ...args, project: BOOK });
  try {
    await scm({ action: "settings", folder });
    await scm({ action: "init" });
    await scm({ action: "identity", name: "Access Suite", email: "access@example.com" });
    const first = await scm({ action: "commit", message: "the fixture", by: "access.mjs" });
    check("the database's modules commit, saving on the way", (first.committed ?? []).includes("Helpers") && first.status?.dirty === false,
      `${(first.committed ?? []).length} committed, dirty ${first.status?.dirty}`);

    helpersText = (await api.readModule("Helpers", BOOK)).text ?? "";
    await api.writeModule("Helpers", helpersText.replace("hello from access", "hello unsaved"), BOOK, { by: "access.mjs" });
    const dirty = await waitFor("the project to read as dirty", async () => (await scm()).dirty === true, { budgetMs: 10000 })
      .then(() => true).catch(() => false);
    check("an unsaved module edit makes the project dirty", dirty);
    const refused = await refusalOf(scm({ action: "checkout", ref: "HEAD" }));
    check("and a checkout over it is refused, as it is in Excel", refused !== null && /unsaved|save/i.test(refused), refused ?? "went through");
    check("with the edit still in the module", ((await api.readModule("Helpers", BOOK)).text ?? "").includes("hello unsaved"));

    await api.command("save");
    const clean = await waitFor("the save to clean the project", async () => (await scm()).dirty === false, { budgetMs: 10000 })
      .then(() => true).catch(() => false);
    check("the editor's own Save cleans it", clean);

    const second = await scm({ action: "commit", message: "the edit", modules: ["Helpers"], by: "access.mjs" });
    check("and the edit commits, the database clean",
      (second.committed ?? []).includes("Helpers") && second.status?.dirty === false, `dirty ${second.status?.dirty}`);

    // Put back through a commit too, which is the save-then-commit path over a dirty project.
    await api.writeModule("Helpers", helpersText, BOOK, { by: "access.mjs" });
    const third = await scm({ action: "commit", message: "put back", modules: ["Helpers"], by: "access.mjs" });
    check("a commit over an unsaved edit saves the database itself first",
      (third.committed ?? []).includes("Helpers") && third.status?.dirty === false, `dirty ${third.status?.dirty}`);
    helpersText = null;
  } finally {
    await scm({ action: "forget" }).catch(() => {});
    await wait(500);
    rmSync(scratch, { recursive: true, force: true });
  }

  const stats = await api.stats();
  check("no COM wrapper was leaked by any of it", stats.comWrappersLive < 100, `${stats.comWrappersLive} live`);
} finally {
  if (helpersText !== null) {
    await api.writeModule("Helpers", helpersText, BOOK, { by: "access.mjs" }).catch(() => {});
  }
}

done();
