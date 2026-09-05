/*
 * The folder view's shape: modules grouped by the '@Folder("Parent.Child") annotation each one
 * carries, the Rubberduck convention (#23). The host reads the annotations; this only arranges
 * what it said. Pure, so the arithmetic is tested by calling it rather than by clicking.
 *
 * A folder is a path of names with a dot between a parent and its child. Names are matched
 * without regard to case, because VBA is, and the first spelling seen is the one drawn - so
 * '@Folder("shared") and '@Folder("Shared") are one folder, called whatever the first module
 * called it. Folders come before modules at every level and sort by name; the modules keep the
 * order the flat tree gives them, which the caller supplies.
 *
 * A module with no annotation sits at the workbook's root, which is where the original puts
 * them too (its default folder is named after the project, and here the project IS the root).
 */

export interface FolderNode<T> {
  /** The folder's own name, spelled as the first module to name it spelled it. */
  name: string;
  /** The dotted path from the root, in the spellings drawn. Empty for the root itself. */
  path: string;
  folders: FolderNode<T>[];
  modules: T[];
}

/** The folder path a component carries, or nothing. */
export interface Foldered {
  folder?: string | null;
}

/** The segments of a folder path, trimmed, with the empty ones dropped. */
export function folderSegments(path: string | null | undefined): string[] {
  if (!path) {
    return [];
  }
  return path.split(".").map((segment) => segment.trim()).filter((segment) => segment.length > 0);
}

/** The key two spellings of one folder share. */
export function folderKey(path: string): string {
  return path.toLowerCase();
}

/**
 * Arranges components into folders. `order` sorts the modules within a folder; folders sort by
 * name among themselves. The root's `folders` are the top-level ones, its `modules` the
 * unannotated ones.
 */
export function buildFolderTree<T extends Foldered>(components: readonly T[], order: (a: T, b: T) => number): FolderNode<T> {
  const root: FolderNode<T> = { name: "", path: "", folders: [], modules: [] };

  for (const component of components) {
    let at = root;
    for (const segment of folderSegments(component.folder)) {
      const key = folderKey(segment);
      let child = at.folders.find((folder) => folderKey(folder.name) === key);
      if (!child) {
        child = {
          name: segment,
          path: at.path ? `${at.path}.${segment}` : segment,
          folders: [],
          modules: [],
        };
        at.folders.push(child);
      }
      at = child;
    }
    at.modules.push(component);
  }

  const sortNode = (node: FolderNode<T>): void => {
    node.folders.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
    node.modules.sort(order);
    for (const folder of node.folders) {
      sortNode(folder);
    }
  };
  sortNode(root);

  return root;
}

/** Every folder in the tree, parents before children, root left out. */
export function allFolders<T>(root: FolderNode<T>): FolderNode<T>[] {
  const out: FolderNode<T>[] = [];
  const walk = (node: FolderNode<T>): void => {
    for (const folder of node.folders) {
      out.push(folder);
      walk(folder);
    }
  };
  walk(root);
  return out;
}

/** The paths of a folder and every folder above it, root-most first: "A", "A.B", "A.B.C". */
export function ancestorPaths(path: string | null | undefined): string[] {
  const segments = folderSegments(path);
  const out: string[] = [];
  for (let depth = 1; depth <= segments.length; depth++) {
    out.push(segments.slice(0, depth).join("."));
  }
  return out;
}
