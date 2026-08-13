/*
 * Builds the form fixture's contents in the open workbook: the form declared by
 * form-plan.mjs, its code-behind, saved and proven to compile.
 *
 * The same buildForm the designer suite verifies against, which is the point: the fixture on
 * disk and the suite's expectations are one declaration.
 */
import { open } from "./xlide-api.mjs";
import { FORM_CODE, FORM_MODULE, buildForm } from "./form-plan.mjs";

const api = await open({});

const health = await api.doctor();
if (!health.healthy) {
  console.error(`the door is not healthy: ${health.findings.join("; ")}`);
  process.exit(1);
}

const { projectId: project } = await api.project();

await buildForm(api, project);
await api.writeModule(FORM_MODULE, FORM_CODE, project);

const design = await api.designer(FORM_MODULE, project);
console.log(`  ${FORM_MODULE}: ${design.controls.length} control(s), "${design.form.caption}", ${design.form.width}x${design.form.height}`);

await api.command("save");

// The code-behind names real controls, so a compile is the proof the form and its code agree.
const compiled = await api.compile();
if (!compiled.started || !compiled.compiled) {
  console.error(`  the fixture does not compile: ${JSON.stringify(compiled)}`);
  process.exit(1);
}

console.log("  saved, and it compiles");
