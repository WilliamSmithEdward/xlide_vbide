// The editor's internal contribution modules ship without type declarations. They are imported
// for their side effects only (registering a controller), so an untyped module is all they need
// to be.
declare module "monaco-editor/editor/contrib/*";

// Two internals are imported for what they HOLD rather than for what they do when loaded, so
// these say what is taken out of them and no more. The editor builds its right-click menu from
// this registry, and an entry it registers unconditionally can only be moved by reaching the
// entry itself.
declare module "monaco-editor/platform/actions/common/actions.js" {
  /** As much of a menu entry as this surface reads or writes. */
  export interface MenuEntry {
    command?: { id: string };
    submenu?: unknown;
    when?: unknown;
    group?: string;
    order?: number;
  }

  export const MenuId: {
    readonly EditorContext: unknown;
    readonly EditorContextPeek: unknown;
  };

  export const MenuRegistry: {
    getMenuItems(id: unknown): MenuEntry[];
    appendMenuItem(id: unknown, item: MenuEntry): { dispose(): void };
  };
}

declare module "monaco-editor/platform/contextkey/common/contextkey.js" {
  export const ContextKeyExpr: {
    /** Never matches, which is how an entry is taken off a menu without being unregistered. */
    false(): unknown;
  };
}
