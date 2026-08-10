/*
 * The split tree, as arithmetic - no DOM, no pointer, no Monaco.
 *
 * Both layouts are the same shape: a section of tool panes and the editor area are each a
 * tree of splits ending in groups, and every drag does one of four things to it - put a
 * thing in a group, split a group beside another, take a thing out, or dissolve what is
 * left empty. That last one is where the bugs live: removing a pane can empty a group, which
 * can leave a split with one child, which must collapse into its parent, and the sizes have
 * to stay a partition of one through all of it.
 *
 * It lives here, separately and purely, so it can be tested by calling it. The live probes
 * exercise it through real drags, which is the right test of the GESTURE and a slow, flaky
 * way to test the ALGEBRA: a tree that mis-collapses shows up as a pane that vanished, three
 * drags later, on one machine.
 */

/** A group of tabs. The active tab is one of them, or empty when the group is being built. */
export interface TreeGroup<T> {
  kind: "group";
  tabs: T[];
  active: T;
}

/** A row or column of children, with sizes that partition one. */
export interface TreeSplit<T> {
  kind: "split";
  direction: "row" | "column";
  children: TreeNode<T>[];
  sizes: number[];
}

export type TreeNode<T> = TreeGroup<T> | TreeSplit<T>;

/** Which side of a group a newcomer lands on. */
export type SplitSide = "left" | "right" | "top" | "bottom";

/** The first group in a tree, reading depth first, or null for an empty one. */
export function firstGroup<T>(node: TreeNode<T> | null): TreeGroup<T> | null {
  if (!node) {
    return null;
  }
  if (node.kind === "group") {
    return node;
  }
  for (const child of node.children) {
    const found = firstGroup(child);
    if (found) {
      return found;
    }
  }
  return null;
}

/** Every group in a tree, depth first. */
export function allGroups<T>(node: TreeNode<T> | null): TreeGroup<T>[] {
  if (!node) {
    return [];
  }
  if (node.kind === "group") {
    return [node];
  }
  return node.children.flatMap(allGroups);
}

/** The group holding a tab, or null. */
export function groupHolding<T>(node: TreeNode<T> | null, tab: T): TreeGroup<T> | null {
  return allGroups(node).find((group) => group.tabs.includes(tab)) ?? null;
}

/**
 * Drops empty groups and collapses what that leaves behind: a split with one child becomes
 * that child, a split with none becomes nothing, and the remaining sizes are renormalised so
 * they still partition one.
 *
 * Renormalising matters. A split whose sizes no longer sum to one lays out at the wrong
 * proportions - subtly, so it reads as a rendering quirk rather than as arithmetic.
 */
export function prune<T>(node: TreeNode<T> | null): TreeNode<T> | null {
  if (!node) {
    return null;
  }

  if (node.kind === "group") {
    return node.tabs.length > 0 ? node : null;
  }

  const children: TreeNode<T>[] = [];
  const sizes: number[] = [];
  node.children.forEach((child, index) => {
    const kept = prune(child);
    if (kept) {
      children.push(kept);
      sizes.push(node.sizes[index] ?? 1);
    }
  });

  if (children.length === 0) {
    return null;
  }
  if (children.length === 1) {
    return children[0]!;
  }

  const total = sizes.reduce((sum, size) => sum + size, 0) || 1;
  return { kind: "split", direction: node.direction, children, sizes: sizes.map((size) => size / total) };
}

/**
 * Puts a newcomer beside an existing group, splitting the tree where that group sits.
 *
 * A split in the SAME direction absorbs the newcomer as a sibling rather than nesting a
 * two-deep tree of one axis - which is what keeps three panes side by side looking like
 * three panes side by side, with two splitters, instead of a pane next to a pair.
 */
export function splitBeside<T>(
  root: TreeNode<T> | null,
  target: TreeGroup<T>,
  side: SplitSide,
  newcomer: TreeGroup<T>,
): TreeNode<T> | null {
  if (!root) {
    return newcomer;
  }

  const direction: "row" | "column" = side === "left" || side === "right" ? "row" : "column";
  const first = side === "left" || side === "top";

  const replace = (node: TreeNode<T>): TreeNode<T> => {
    if (node === target) {
      return {
        kind: "split",
        direction,
        children: first ? [newcomer, node] : [node, newcomer],
        sizes: [0.5, 0.5],
      };
    }

    if (node.kind !== "split") {
      return node;
    }

    const children = node.children.map(replace);

    // Absorb a same-direction split, carrying its children's sizes through their parent's
    // share so the proportions on screen do not jump.
    const flattened: TreeNode<T>[] = [];
    const sizes: number[] = [];
    children.forEach((child, index) => {
      if (child.kind === "split" && child.direction === node.direction) {
        child.children.forEach((inner, innerIndex) => {
          flattened.push(inner);
          sizes.push((node.sizes[index] ?? 1) * (child.sizes[innerIndex] ?? 1));
        });
      } else {
        flattened.push(child);
        sizes.push(node.sizes[index] ?? 1);
      }
    });

    const total = sizes.reduce((sum, size) => sum + size, 0) || 1;
    return { kind: "split", direction: node.direction, children: flattened, sizes: sizes.map((s) => s / total) };
  };

  return replace(root);
}

/**
 * Moves a divider inside a split: the pair it sits between trade space, and neither may go
 * below a minimum expressed as a fraction of the whole. Only the pair changes, so everything
 * else on screen stays exactly where it was.
 */
export function resizeAt(sizes: number[], index: number, delta: number, minimum: number): number[] {
  if (index <= 0 || index >= sizes.length) {
    return [...sizes];
  }

  const before = sizes[index - 1] ?? 0;
  const after = sizes[index] ?? 0;
  const pair = before + after;
  const floor = Math.min(minimum, pair / 2);

  const next = [...sizes];
  const grown = Math.max(floor, Math.min(pair - floor, before + delta));
  next[index - 1] = grown;
  next[index] = pair - grown;
  return next;
}

/**
 * Where in a strip a position means, given each tab's midpoint: before the first tab whose
 * midpoint has not been crossed, or at the end.
 */
export function indexAtMidpoints(midpoints: number[], position: number): number {
  for (let i = 0; i < midpoints.length; i++) {
    if (position < midpoints[i]!) {
      return i;
    }
  }
  return midpoints.length;
}
