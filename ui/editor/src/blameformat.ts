/*
 * What a blame hint says, worked out without monaco so it can be tested in node.
 *
 * The host answers `scm?action=blame` with one row per committed line of the LIVE module - the
 * mapping from the committed file's lines to the editor's is the host's, through the same diff
 * that draws the rows - and a list of live lines that are not committed. This file turns those
 * into the words the editor paints at the end of each line and shows on hover. The layer that
 * paints them (blame.ts) owns the decorations; this owns the spelling, so a change to the
 * spelling fails here in milliseconds rather than in Excel.
 */

/** One committed line as the host reports it: the LIVE line number and who last touched it. */
export interface BlameRow {
  line: number;
  hash: string;
  short: string;
  author: string;
  /** ISO 8601 with the author's offset, as git blame's author-time and author-tz spell it. */
  when: string;
  summary: string;
}

/** The blame of a whole module, coerced from the wire. */
export interface BlameReading {
  head: string;
  lines: BlameRow[];
  uncommitted: number[];
}

/** The hint on a line that is not committed. Leading spaces keep it off the code's last word. */
export const UNCOMMITTED_HINT = "  not committed";

/** yyyy-mm-dd from an ISO timestamp, in the author's own day; "" when the stamp cannot be read. */
export function blameDate(when: string): string {
  // The DATE PART OF THE TEXT, not the local day of the instant: git prints author-time in the
  // author's offset, and a commit made at 23:30 in one zone is still that day's commit when
  // read in another. Ten characters of a well-formed stamp are the date; anything shorter or
  // stranger is parsed as a last resort.
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(when);
  if (match) {
    return `${match[1]}-${match[2]}-${match[3]}`;
  }

  const parsed = new Date(when);
  if (Number.isNaN(parsed.getTime())) {
    return "";
  }

  const pad = (part: number): string => String(part).padStart(2, "0");
  return `${parsed.getFullYear()}-${pad(parsed.getMonth() + 1)}-${pad(parsed.getDate())}`;
}

/** The seven characters a person quotes a commit by, from whichever form the row carries. */
export function shortHashOf(row: { short?: string; hash?: string }): string {
  const short = (row.short ?? "").trim();
  if (short.length > 0) {
    return short;
  }

  return (row.hash ?? "").trim().slice(0, 7);
}

/** `  author, yyyy-mm-dd, short` - the end-of-line hint. Empty parts are left out, not blanked. */
export function blameHint(row: BlameRow): string {
  const parts = [row.author.trim(), blameDate(row.when), shortHashOf(row)]
    .filter((part) => part.length > 0);
  return parts.length === 0 ? "" : `  ${parts.join(", ")}`;
}

/** The hover's markdown: the commit's summary under who and when. */
export function blameHover(row: BlameRow): string {
  const short = shortHashOf(row);
  const who = [row.author.trim(), row.when.trim()].filter((part) => part.length > 0).join(", ");
  const head = [short ? `**${short}**` : "", who].filter((part) => part.length > 0).join(" ");
  const summary = row.summary.trim();
  return [head, summary].filter((part) => part.length > 0).join("\n\n");
}

/**
 * The host's reply as a reading, tolerating what the wire may not carry.
 *
 * Every field is coerced rather than trusted: a row whose line is not a positive integer is
 * dropped, an uncommitted entry that is not a number is dropped, and a reply that carries an
 * error or no `lines` reads as nothing committed - which paints nothing, rather than painting
 * a hint on the wrong line.
 */
export function blameReadingOf(reply: Record<string, unknown> | null | undefined): BlameReading {
  const lines: BlameRow[] = [];
  const rawLines = Array.isArray(reply?.lines) ? reply.lines : [];
  for (const raw of rawLines) {
    if (typeof raw !== "object" || raw === null) {
      continue;
    }

    const row = raw as Record<string, unknown>;
    const line = Number(row.line);
    if (!Number.isInteger(line) || line < 1) {
      continue;
    }

    lines.push({
      line,
      hash: typeof row.hash === "string" ? row.hash : "",
      short: typeof row.short === "string" ? row.short : "",
      author: typeof row.author === "string" ? row.author : "",
      when: typeof row.when === "string" ? row.when : "",
      summary: typeof row.summary === "string" ? row.summary : "",
    });
  }

  lines.sort((a, b) => a.line - b.line);

  const rawUncommitted = Array.isArray(reply?.uncommitted) ? reply.uncommitted : [];
  const uncommitted = [...new Set(rawUncommitted
    .map((one) => Number(one))
    .filter((one) => Number.isInteger(one) && one >= 1))]
    .sort((a, b) => a - b);

  return {
    head: typeof reply?.head === "string" ? reply.head : "",
    lines,
    uncommitted,
  };
}
