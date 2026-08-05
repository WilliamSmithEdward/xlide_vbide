/*
 * The xlide Object Browser: the referenced type libraries and the project itself, browsable
 * and searchable, in place of the native browser the canvas retired (lesson 32 — it could
 * neither float nor be adopted).
 *
 * A library picker chooses the subject: each open workbook's project, or any type library
 * the projects reference — Excel, VBA, Office, stdole — served by the host's typelib reader.
 * Types on the left, members with VBA-spelled signatures on the right, the selected member's
 * full signature in the detail strip. Project members jump to their definition.
 */

import type { ExplorerProcedure, ExplorerProject } from "./explorer.js";
import type { ObLibrary, ObMember, ObType } from "./bridge.js";

export interface ObjectBrowserDeps {
  /** The workspace as the host last published it. */
  projects(): ExplorerProject[];
  /** A module's procedures, from the engine; null when no answer came. */
  outline(module: string, workbook?: string): Promise<ExplorerProcedure[] | null>;
  /** The referenced libraries; null when no answer came. */
  libraries(): Promise<ObLibrary[] | null>;
  /** A library's types; null when no answer came. */
  types(library: string): Promise<ObType[] | null>;
  /** A type's members; null when no answer came. */
  members(library: string, typeName: string): Promise<ObMember[] | null>;
  /** Jump to a member's definition, for the project's own members. */
  navigate(module: string, line: number, workbook?: string): void;
  /** The view closed; focus belongs to the editor again. */
  closed(): void;
}

const MODULE_ICONS: Record<number, string> = {
  1: "symbol-namespace",
  2: "symbol-class",
  3: "window",
  100: "file",
};

const TYPE_ICONS: Record<string, string> = {
  class: "symbol-class",
  enum: "symbol-enum",
  module: "symbol-namespace",
  type: "symbol-structure",
};

function memberIcon(kind: string): string {
  if (kind.startsWith("Property")) {
    return "symbol-property";
  }

  switch (kind) {
    case "Const":
      return "symbol-constant";
    case "Field":
      return "symbol-field";
    case "Event":
      return "symbol-event";
    default:
      return "symbol-method";
  }
}

/** What the picker points at: one of the workbook projects, or one host library. */
type Scope =
  | { project: ExplorerProject }
  | { library: ObLibrary };

/**
 * Opens the browser. One at a time: opening while open focuses the search box of the one
 * that exists, which is also what a second press of its key should mean.
 */
export function openObjectBrowser(deps: ObjectBrowserDeps): void {
  const existing = document.getElementById("objbrowser-search") as HTMLInputElement | null;
  if (existing) {
    existing.focus();
    existing.select();
    return;
  }

  const backdrop = document.createElement("div");
  backdrop.id = "objbrowser-backdrop";

  const card = document.createElement("div");
  card.id = "objbrowser-card";
  card.setAttribute("role", "dialog");
  card.setAttribute("aria-modal", "true");
  card.setAttribute("aria-label", "Object Browser");

  const head = document.createElement("div");
  head.id = "objbrowser-head";

  const title = document.createElement("span");
  title.id = "objbrowser-title";
  title.textContent = "Object Browser";

  const picker = document.createElement("select");
  picker.id = "objbrowser-library";
  picker.setAttribute("aria-label", "Library");

  const search = document.createElement("input");
  search.id = "objbrowser-search";
  search.type = "text";
  search.placeholder = "Search";
  search.setAttribute("aria-label", "Search types and members");

  const close = document.createElement("button");
  close.type = "button";
  close.id = "objbrowser-close";
  close.title = "Close (Esc)";
  close.setAttribute("aria-label", "Close the Object Browser");
  close.innerHTML = '<span class="codicon codicon-close" aria-hidden="true"></span>';

  head.append(title, picker, search, close);

  const body = document.createElement("div");
  body.id = "objbrowser-body";

  const typesPane = document.createElement("div");
  typesPane.id = "objbrowser-modules";
  typesPane.setAttribute("role", "listbox");
  typesPane.setAttribute("aria-label", "Types");

  const membersPane = document.createElement("div");
  membersPane.id = "objbrowser-members";
  membersPane.setAttribute("role", "listbox");
  membersPane.setAttribute("aria-label", "Members");

  body.append(typesPane, membersPane);

  const detail = document.createElement("div");
  detail.id = "objbrowser-detail";
  detail.textContent = "Loading…";

  card.append(head, body, detail);
  backdrop.appendChild(card);
  document.body.appendChild(backdrop);
  search.focus();

  // --- state ---------------------------------------------------------------------------

  const scopes: Scope[] = [];
  let scope: Scope | null = null;
  let selectedType: string | null = null;
  let disposed = false;

  /** Outlines by "workbook module"; library members by "library type". */
  const outlines = new Map<string, ExplorerProcedure[]>();
  const libraryTypes = new Map<string, ObType[]>();
  const libraryMembers = new Map<string, ObMember[]>();

  const dismiss = (): void => {
    disposed = true;
    document.removeEventListener("keydown", onKey, true);
    backdrop.remove();
    deps.closed();
  };

  const onKey = (event: KeyboardEvent): void => {
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      dismiss();
    }
  };

  document.addEventListener("keydown", onKey, true);
  close.addEventListener("click", dismiss);
  backdrop.addEventListener("click", (event) => {
    if (event.target === backdrop) {
      dismiss();
    }
  });

  const query = (): string => search.value.trim().toLowerCase();

  // --- members pane --------------------------------------------------------------------

  const showMember = (icon: string, name: string, context: string, description: string, jump?: () => void): void => {
    const item = document.createElement("div");
    item.className = "objbrowser-row";
    item.setAttribute("role", "option");
    item.tabIndex = 0;

    const glyph = document.createElement("span");
    glyph.className = `codicon codicon-${icon}`;
    glyph.setAttribute("aria-hidden", "true");

    const label = document.createElement("span");
    label.className = "objbrowser-name";
    label.textContent = name;

    const trailing = document.createElement("span");
    trailing.className = "objbrowser-context";
    trailing.textContent = context;

    item.append(glyph, label, trailing);
    item.addEventListener("click", () => {
      detail.textContent = description;
      for (const other of membersPane.querySelectorAll(".objbrowser-row.selected")) {
        other.classList.remove("selected");
      }
      item.classList.add("selected");
    });

    if (jump) {
      item.addEventListener("dblclick", jump);
      item.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          jump();
        }
      });
    }

    membersPane.appendChild(item);
  };

  const renderMembers = (): void => {
    membersPane.replaceChildren();
    const wanted = query();
    let shown = 0;

    if (scope && "project" in scope) {
      const project = scope.project;
      for (const component of project.components) {
        if (wanted.length === 0 && component.name !== selectedType) {
          continue;
        }

        const held = outlines.get(`${project.name} ${component.name}`) ?? [];
        for (const procedure of held) {
          if (wanted.length > 0 && !procedure.name.toLowerCase().includes(wanted)) {
            continue;
          }

          shown++;
          showMember(
            memberIcon(procedure.kind),
            procedure.name,
            wanted.length > 0 ? component.name : procedure.kind,
            `${procedure.kind} ${procedure.name} — ${project.name}.${component.name}, line ${procedure.line}`,
            () => {
              deps.navigate(component.name, procedure.line, project.name);
              dismiss();
            });
        }
      }
    } else if (scope && "library" in scope) {
      const library = scope.library.name;
      const types = libraryTypes.get(library) ?? [];
      for (const type of types) {
        if (wanted.length === 0 && type.name !== selectedType) {
          continue;
        }

        const held = libraryMembers.get(`${library} ${type.name}`);
        if (!held) {
          continue;
        }

        for (const member of held) {
          if (wanted.length > 0 && !member.name.toLowerCase().includes(wanted)) {
            continue;
          }

          shown++;
          showMember(
            memberIcon(member.kind),
            member.name,
            wanted.length > 0 ? type.name : member.kind,
            member.signature + ` — Member of ${library}.${type.name}`
              + (member.description.length > 0 ? ` — ${member.description}` : ""));
        }
      }
    }

    if (shown === 0) {
      const empty = document.createElement("div");
      empty.className = "objbrowser-empty";
      empty.textContent = wanted.length > 0
        ? "Nothing matches among the loaded members."
        : selectedType
          ? "No members here."
          : "Pick a type on the left, or search.";
      membersPane.appendChild(empty);
    }
  };

  // --- types pane ----------------------------------------------------------------------

  const pickType = (name: string): void => {
    selectedType = name;

    if (scope && "library" in scope) {
      const library = scope.library.name;
      if (!libraryMembers.has(`${library} ${name}`)) {
        void deps.members(library, name).then((rows) => {
          if (!disposed && rows) {
            libraryMembers.set(`${library} ${name}`, rows);
            renderMembers();
          }
        });
      }
    }

    renderTypes();
    renderMembers();
  };

  const showType = (icon: string, name: string, trailing: string): void => {
    const item = document.createElement("div");
    item.className = "objbrowser-row";
    item.setAttribute("role", "option");
    item.tabIndex = 0;

    if (name === selectedType) {
      item.classList.add("selected");
    }

    const glyph = document.createElement("span");
    glyph.className = `codicon codicon-${icon}`;
    glyph.setAttribute("aria-hidden", "true");

    const label = document.createElement("span");
    label.className = "objbrowser-name";
    label.textContent = name;

    const context = document.createElement("span");
    context.className = "objbrowser-context";
    context.textContent = trailing;

    item.append(glyph, label, context);
    item.addEventListener("click", () => pickType(name));
    item.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        pickType(name);
      }
    });

    typesPane.appendChild(item);
  };

  const renderTypes = (): void => {
    typesPane.replaceChildren();
    const wanted = query();

    if (scope && "project" in scope) {
      for (const component of scope.project.components) {
        if (wanted.length > 0 && !component.name.toLowerCase().includes(wanted)) {
          const held = outlines.get(`${scope.project.name} ${component.name}`) ?? [];
          if (!held.some((procedure) => procedure.name.toLowerCase().includes(wanted))) {
            continue;
          }
        }

        showType(MODULE_ICONS[component.kind] ?? "symbol-namespace", component.name, "");
      }
    } else if (scope && "library" in scope) {
      const types = libraryTypes.get(scope.library.name);
      if (!types) {
        const loading = document.createElement("div");
        loading.className = "objbrowser-empty";
        loading.textContent = "Loading types…";
        typesPane.appendChild(loading);
        return;
      }

      for (const type of types) {
        if (wanted.length > 0 && !type.name.toLowerCase().includes(wanted)) {
          const held = libraryMembers.get(`${scope.library.name} ${type.name}`);
          if (!held || !held.some((member) => member.name.toLowerCase().includes(wanted))) {
            continue;
          }
        }

        showType(TYPE_ICONS[type.kind] ?? "symbol-class", type.name, type.kind);
      }
    }
  };

  // --- scope switching -----------------------------------------------------------------

  const adoptScope = (next: Scope): void => {
    scope = next;
    selectedType = null;

    if ("project" in next) {
      detail.textContent = `${next.project.name} — this workbook's project`;
      for (const component of next.project.components) {
        const key = `${next.project.name} ${component.name}`;
        if (!outlines.has(key)) {
          void deps.outline(component.name, next.project.name).then((procedures) => {
            if (!disposed && procedures) {
              outlines.set(key, procedures);
              renderTypes();
              renderMembers();
            }
          });
        }
      }
    } else {
      detail.textContent = `${next.library.name} — ${next.library.description}`;
      if (!libraryTypes.has(next.library.name)) {
        void deps.types(next.library.name).then((rows) => {
          if (!disposed && rows) {
            libraryTypes.set(next.library.name, rows);
            renderTypes();
            renderMembers();
          }
        });
      }
    }

    renderTypes();
    renderMembers();
  };

  picker.addEventListener("change", () => {
    const picked = scopes[picker.selectedIndex];
    if (picked) {
      adoptScope(picked);
    }
  });

  search.addEventListener("input", () => {
    renderTypes();
    renderMembers();
  });

  // --- data ----------------------------------------------------------------------------

  for (const project of deps.projects()) {
    scopes.push({ project });
    const option = document.createElement("option");
    option.textContent = `${project.name} (project)`;
    picker.appendChild(option);
  }

  void deps.libraries().then((rows) => {
    if (disposed || !rows) {
      return;
    }

    for (const library of rows) {
      scopes.push({ library });
      const option = document.createElement("option");
      option.textContent = library.name;
      option.title = library.description;
      picker.appendChild(option);
    }

    if (scope === null && scopes.length > 0) {
      picker.selectedIndex = 0;
      adoptScope(scopes[0]);
    }
  });

  if (scopes.length > 0) {
    picker.selectedIndex = 0;
    adoptScope(scopes[0]);
  } else {
    detail.textContent = "Loading…";
  }
}
