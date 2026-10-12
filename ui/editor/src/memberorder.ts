/**
 * The order an unfolded module lists its members in the explorer (#92).
 *
 * Three orders: "code" as written, "name" alphabetically, "kind" by kind and then by name. Under
 * every one, a property's Get, Let and Set sit together, in that order, where the first of them
 * would have sat: they are one property to the developer looking for it, and VBA lets the legs
 * stand anywhere in the module.
 */

export type MemberSort = "code" | "name" | "kind";

export interface Member {
  name: string;
  kind: string;
  line: number;
}

/** Where each kind falls under "kind": declarations first, then the procedures. */
const KIND_RANK: Record<string, number> = {
  Enum: 0,
  Type: 1,
  Sub: 2,
  Function: 3,
  "Property Get": 4,
  "Property Let": 4,
  "Property Set": 4,
};

/** A property's legs in the order they are read: Get, then Let, then Set. */
const LEG_RANK: Record<string, number> = { "Property Get": 0, "Property Let": 1, "Property Set": 2 };

const isProperty = (member: Member): boolean => member.kind.startsWith("Property ");

/** The members in the chosen order, a property's legs together. Never mutates its input. */
export function orderMembers<T extends Member>(members: readonly T[], sort: MemberSort): T[] {
  // One entry per property name, holding its legs; everything else is an entry of its own.
  const entries: { first: T; rows: T[] }[] = [];
  const properties = new Map<string, { first: T; rows: T[] }>();
  for (const member of members) {
    if (!isProperty(member)) {
      entries.push({ first: member, rows: [member] });
      continue;
    }

    const key = member.name.toLowerCase();
    const held = properties.get(key);
    if (held) {
      held.rows.push(member);
      continue;
    }

    const entry = { first: member, rows: [member] };
    properties.set(key, entry);
    entries.push(entry);
  }

  for (const entry of properties.values()) {
    entry.rows.sort((a, b) => (LEG_RANK[a.kind] ?? 9) - (LEG_RANK[b.kind] ?? 9) || a.line - b.line);
  }

  const byName = (a: T, b: T): number =>
    a.name.localeCompare(b.name, undefined, { sensitivity: "base" }) || a.line - b.line;

  if (sort === "name") {
    entries.sort((a, b) => byName(a.first, b.first));
  } else if (sort === "kind") {
    entries.sort((a, b) =>
      (KIND_RANK[a.first.kind] ?? 9) - (KIND_RANK[b.first.kind] ?? 9) || byName(a.first, b.first));
  }

  return entries.flatMap((entry) => entry.rows);
}
