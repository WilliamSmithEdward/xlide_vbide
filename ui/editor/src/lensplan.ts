/*
 * What stands above each procedure: the lenses, worked out from the module's own text.
 *
 * Two kinds. A test the runner knows about gets "Run test" and "Debug test"; every procedure gets
 * a count of where it is used. Nothing here asks anybody anything - the count is a question for
 * the engine, asked later and only for the lenses on screen - so this is a pure function of the
 * text, the procedures found in it, and the runner's list, and it is tested as one.
 *
 * WHERE A LENS SITS is the page's reading of the text as it stands, not the line the runner last
 * reported. The runner rediscovers after a write, so its line lags the typing; a lens placed by
 * it would hang one line off for exactly as long as the developer is looking at it.
 *
 * WHICH PROCEDURES ARE TESTS is the runner's reading and nobody else's. A Sub marked a moment ago
 * and not yet discovered has no Run lens, the same way it has no row in the Tests pane: a lens
 * that the runner would answer with "no test named that" is an offer nothing stands behind.
 */

import { headerNameColumn, type ProcedureRange } from "./procedureat.js";

export type LensKind = "runTest" | "debugTest" | "references";

export interface PlannedLens {
  kind: LensKind;
  /** The header line the lens sits above. */
  line: number;
  /** The column the procedure's name starts at, 1-based: where the references question is asked. */
  column: number;
  procedure: string;
  /** The runner's id for the test, on its two lenses; absent on a references lens. */
  test?: string;
}

/** A test as the runner lists it, for the module being planned. */
export interface KnownTest {
  id: string;
  procedure: string;
}

/**
 * The lenses for one module, in the order they are drawn: top to bottom, and on a test's line
 * Run, Debug, then the count.
 *
 * `headerLine` reads a line of the text the procedures were scanned from. A header the scan found
 * and this cannot read a name out of gets no lens, rather than one in the wrong column.
 */
export function planLenses(
  procedures: readonly ProcedureRange[],
  headerLine: (line: number) => string,
  tests: readonly KnownTest[],
): PlannedLens[] {
  // VBA names do not distinguish case, and neither does the runner's list.
  const testsByName = new Map(tests.map((test) => [test.procedure.toLowerCase(), test.id]));
  const planned: PlannedLens[] = [];

  for (const procedure of procedures) {
    const column = headerNameColumn(headerLine(procedure.header));
    if (column === null) {
      continue;
    }

    const at = { line: procedure.header, column, procedure: procedure.name };

    // Only a Sub is ever a test; a Property that shares a test's name is not one.
    const test = procedure.kind === "Sub" ? testsByName.get(procedure.name.toLowerCase()) : undefined;
    if (test !== undefined) {
      planned.push({ kind: "runTest", ...at, test }, { kind: "debugTest", ...at, test });
    }

    planned.push({ kind: "references", ...at });
  }

  return planned;
}

/** What a test's lenses say. */
export const RUN_TEST_TITLE = "Run test";
export const DEBUG_TEST_TITLE = "Debug test";

/**
 * What the count says. Null is a count nobody could give - no engine, or no answer in time - and
 * reads as the bare word, which still opens the list: a lens that vanished whenever the engine
 * was busy would be a control that comes and goes.
 */
export function referencesTitle(count: number | null): string {
  if (count === null) {
    return "references";
  }

  return count === 0 ? "no references" : count === 1 ? "1 reference" : `${count} references`;
}
