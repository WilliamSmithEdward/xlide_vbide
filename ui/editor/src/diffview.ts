/*
 * One side-by-side comparison, drawn one way.
 *
 * The import/export dialog and the Changes pane both show "this text became that text", and both
 * are handed the same rows by the same comparison in the host (ModuleSync.Diff). Drawing them
 * twice would be two chances for the two surfaces to disagree about what a removed line looks
 * like, so the drawing lives here and both call it.
 *
 * The rows come from the HOST rather than being worked out here on purpose: the comparison is
 * bounded there (a middle of more than 2,000 differing lines is answered as a block rather than
 * lined up, because the table for that is quadratic in memory), and a second implementation in
 * the page would not know that.
 */

/** One row of a side-by-side comparison, as the host reports it. */
export interface SyncDiffLine {
  leftNumber: number | null;
  rightNumber: number | null;
  left: string;
  right: string;
  kind: "equal" | "changed" | "added" | "removed" | "gap";
}

/**
 * Draws the rows into `into`, replacing whatever was there.
 *
 * `prefix` names the css classes, so a pane can dress the rows as its own without the structure
 * being copied: `sync` gives `sync-diff-row`, `sync-gutter`, `sync-code`, `sync-gap`.
 */
export function drawDiffRows(into: HTMLElement, lines: readonly SyncDiffLine[], prefix = "sync"): void {
  into.replaceChildren();

  for (const line of lines) {
    const row = document.createElement("div");
    row.className = `${prefix}-diff-row`;
    row.dataset.kind = line.kind;

    // A gap is a run of identical lines left out, and it spans the whole width: it belongs to
    // neither side, so drawing it in the left column would read as a line of the left file.
    if (line.kind === "gap") {
      const span = document.createElement("span");
      span.className = `${prefix}-gap`;
      span.textContent = line.left;
      row.appendChild(span);
      into.appendChild(row);
      continue;
    }

    // A line only one side has carries no number for the other, and the host spells that absence
    // as null rather than as a missing field. Tested for loosely on purpose: `=== undefined` let a
    // literal "null" through into the gutter.
    const leftNumber = document.createElement("span");
    leftNumber.className = `${prefix}-gutter`;
    leftNumber.textContent = line.leftNumber == null ? "" : String(line.leftNumber);

    const leftText = document.createElement("span");
    leftText.className = `${prefix}-code`;
    leftText.textContent = line.left;

    const rightNumber = document.createElement("span");
    rightNumber.className = `${prefix}-gutter`;
    rightNumber.textContent = line.rightNumber == null ? "" : String(line.rightNumber);

    const rightText = document.createElement("span");
    rightText.className = `${prefix}-code`;
    rightText.textContent = line.right;

    row.append(leftNumber, leftText, rightNumber, rightText);
    into.appendChild(row);
  }
}
