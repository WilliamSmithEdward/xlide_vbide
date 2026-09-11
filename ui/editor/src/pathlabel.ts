/*
 * A long path made short enough for a button by giving way in the MIDDLE.
 *
 * The two ends are the parts that say something: the start names the drive or share the path
 * lives on, and the end names the folder itself - the one a developer chose or is being offered.
 * The run of folders between is what gives way, the way a file dialog shortens a path too. The
 * cuts land on folder separators wherever the path has them, so what is left reads as whole
 * folder names, and a path that already fits comes back exactly as it was. The whole path stays
 * the caller's to show somewhere it can still be read in full: a tooltip.
 */

/** The fewest characters a shortened path keeps, below which neither end would say anything. */
const LEAST = 12;

/** Three dots rather than one character, matching every other "..." the surface draws. */
const GAP = "...";

/** `path` at most `most` characters long, shortened from its middle when it is longer. */
export function shortenPath(path: string, most = 40): string {
  if (path.length <= most) {
    return path;
  }

  const room = Math.max(LEAST, most) - GAP.length;

  // The end names the folder, so it has the larger share of the room.
  const endRoom = Math.ceil(room * 0.6);
  const startRoom = room - endRoom;

  let start = path.slice(0, startRoom);
  const startCut = Math.max(start.lastIndexOf("\\"), start.lastIndexOf("/"));
  if (startCut > 0) {
    start = start.slice(0, startCut + 1);
  }

  let end = path.slice(path.length - endRoom);
  const endCuts = [end.indexOf("\\"), end.indexOf("/")].filter((at) => at >= 0);
  const endCut = endCuts.length > 0 ? Math.min(...endCuts) : -1;
  if (endCut >= 0 && endCut < end.length - 1) {
    end = end.slice(endCut);
  }

  return `${start}${GAP}${end}`;
}
