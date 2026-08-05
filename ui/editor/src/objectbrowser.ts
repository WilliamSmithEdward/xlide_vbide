/*
 * The xlide Object Browser page: the whole document of the floating palette window the host
 * opens beside the editor (the developer's choice, 2026-08-05 — the native browser can
 * neither float nor be adopted, lesson 32, so it retired in favour of this).
 *
 * A library picker chooses the subject: each open workbook's project, or any type library
 * the projects reference — Excel, VBA, Office, stdole — served by the host's typelib reader.
 * Types on the left, members with VBA-spelled signatures on the right, the selected member's
 * full signature in the detail strip. Project members carry their line and jump to their
 * definition in the editor; Escape asks the host to hide the window, exactly like its
 * close box.
 *
 * The page speaks to its own host window directly — it is a second browser surface with its
 * own transport, not a view inside the editor page.
 */

import { webView2Transport } from "./bridge.js";
import type { HostTransport, ObLibrary, ObMember, ObType } from "./bridge.js";

const TYPE_ICONS: Record<string, string> = {
  class: "symbol-class",
  enum: "symbol-enum",
  module: "symbol-namespace",
  type: "symbol-structure",
  form: "window",
  document: "file",
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
    case "Enum":
      return "symbol-enum";
    case "Type":
      return "symbol-structure";
    default:
      return "symbol-method";
  }
}

const byName = <T extends { name: string }>(a: T, b: T): number =>
  a.name.localeCompare(b.name, undefined, { sensitivity: "base" });

/** The palette's questions to its host, id-matched with a timeout apiece. */
class PaletteHost {
  private nextId = 1;
  private readonly pending = new Map<number, { resolve(rows: unknown): void; timer: number }>();

  constructor(private readonly transport: HostTransport) {
    transport.subscribe((message) => {
      if (message.type === "obLibrariesResult" || message.type === "obTypesResult"
        || message.type === "obMembersResult") {
        const waiter = this.pending.get(message.id);
        if (waiter) {
          this.pending.delete(message.id);
          clearTimeout(waiter.timer);
          waiter.resolve(
            message.type === "obLibrariesResult" ? message.libraries ?? []
              : message.type === "obTypesResult" ? message.types ?? []
                : message.members ?? []);
        }
      }
    });
  }

  libraries(): Promise<ObLibrary[] | null> {
    return this.ask((id) => ({ type: "obLibraries", id }));
  }

  types(library: string): Promise<ObType[] | null> {
    return this.ask((id) => ({ type: "obTypes", id, library }));
  }

  members(library: string, typeName: string): Promise<ObMember[] | null> {
    return this.ask((id) => ({ type: "obMembers", id, library, typeName }));
  }

  navigate(module: string, line: number, project?: string): void {
    this.transport.post({ type: "navigate", module, line, column: 1, ...(project ? { project } : {}) });
  }

  close(): void {
    this.transport.post({ type: "close" });
  }

  private ask<T>(message: (id: number) => Parameters<HostTransport["post"]>[0]): Promise<T | null> {
    const id = this.nextId++;
    return new Promise((resolve) => {
      const timer = window.setTimeout(() => {
        this.pending.delete(id);
        resolve(null);
      }, 8000);
      this.pending.set(id, { resolve: resolve as (rows: unknown) => void, timer });
      this.transport.post(message(id));
    });
  }
}

/** Canned answers for opening the page in a plain browser, where there is no host. */
function demoPaletteTransport(): HostTransport {
  let deliver: ((message: never) => void) | null = null;
  const send = (message: unknown): void => {
    setTimeout(() => deliver?.(message as never), 30);
  };

  return {
    post(message) {
      if (message.type === "obLibraries") {
        send({
          type: "obLibrariesResult",
          id: message.id,
          libraries: [
            { name: "scratch.xlsm", description: "VBAProject", kind: "project" },
            { name: "Excel", description: "Microsoft Excel Object Library", kind: "library" },
            { name: "VBA", description: "Visual Basic For Applications", kind: "library" },
          ],
        });
      }
      if (message.type === "obTypes") {
        send({
          type: "obTypesResult",
          id: message.id,
          types: message.library === "scratch.xlsm"
            ? [{ name: "Module1", kind: "module" }, { name: "ThisWorkbook", kind: "document" }]
            : [
              { name: "Range", kind: "class" },
              { name: "Worksheet", kind: "class" },
              { name: "XlDirection", kind: "enum" },
            ],
        });
      }
      if (message.type === "obMembers") {
        send({
          type: "obMembersResult",
          id: message.id,
          members: message.library === "scratch.xlsm"
            ? [
              { name: "Greet", kind: "Sub", signature: "Public Sub Greet(name As String)", description: "", line: 3 },
              { name: "Total", kind: "Function", signature: "Public Function Total() As Double", description: "", line: 9 },
            ]
            : [
              { name: "Address", kind: "Property", signature: "Property Get Address([RowAbsolute As Variant]) As String", description: "Returns the address.", line: 0 },
              { name: "Select", kind: "Function", signature: "Function Select() As Variant", description: "", line: 0 },
            ],
        });
      }
    },
    subscribe(handler) {
      deliver = handler as (message: never) => void;
    },
  };
}

/** Boots the palette page: builds the whole document and starts asking. */
export function bootObjectBrowserPage(): void {
  document.title = "Object Browser";
  document.body.classList.add("objbrowser-page");

  // The document is the editor's index.html; its shell skeleton belongs to the other view.
  document.getElementById("shell")?.remove();

  const host = new PaletteHost(webView2Transport() ?? demoPaletteTransport());

  const root = document.createElement("div");
  root.id = "objbrowser-card";

  const head = document.createElement("div");
  head.id = "objbrowser-head";

  const picker = document.createElement("select");
  picker.id = "objbrowser-library";
  picker.setAttribute("aria-label", "Library");

  const search = document.createElement("input");
  search.id = "objbrowser-search";
  search.type = "text";
  search.placeholder = "Search";
  search.setAttribute("aria-label", "Search types and members");

  // What the search reads: the left pane's groups, the selected type's members, or both.
  const scopePick = document.createElement("select");
  scopePick.id = "objbrowser-scope";
  scopePick.setAttribute("aria-label", "Search scope");
  for (const [value, label] of [["group", "Group"], ["object", "Object"], ["all", "All"]] as const) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label;
    scopePick.appendChild(option);
  }

  head.append(picker, search, scopePick);

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

  // The details pane sits under both lists behind its own splitter, the way the native
  // browser stacked its member pane: signature in the code face, then where the member
  // lives, then whatever the library documents about it.
  const splitter = document.createElement("div");
  splitter.id = "objbrowser-splitter";
  splitter.setAttribute("role", "separator");
  splitter.setAttribute("aria-orientation", "horizontal");
  splitter.setAttribute("aria-label", "Resize the details pane");
  splitter.tabIndex = 0;

  const detail = document.createElement("div");
  detail.id = "objbrowser-detail";
  detail.setAttribute("aria-live", "polite");

  const detailSignature = document.createElement("div");
  detailSignature.id = "objbrowser-detail-signature";

  const detailContext = document.createElement("div");
  detailContext.id = "objbrowser-detail-context";

  const detailDescription = document.createElement("div");
  detailDescription.id = "objbrowser-detail-description";

  detail.append(detailSignature, detailContext, detailDescription);

  const setDetail = (signature: string, context = "", description = ""): void => {
    detailSignature.textContent = signature;
    detailContext.textContent = context;
    detailContext.hidden = context.length === 0;
    detailDescription.textContent = description;
    detailDescription.hidden = description.length === 0;
  };

  setDetail("Loading...");

  const sizeDetail = (height: number): void => {
    const clamped = Math.max(48, Math.min(Math.round(window.innerHeight * 0.6), height));
    detail.style.height = `${clamped}px`;
  };

  splitter.addEventListener("pointerdown", (event) => {
    event.preventDefault();
    splitter.setPointerCapture(event.pointerId);
    const startY = event.clientY;
    const startHeight = detail.getBoundingClientRect().height;
    const move = (moved: PointerEvent): void => sizeDetail(startHeight + (startY - moved.clientY));
    const stop = (): void => {
      splitter.removeEventListener("pointermove", move);
      splitter.removeEventListener("pointerup", stop);
      splitter.removeEventListener("pointercancel", stop);
    };
    splitter.addEventListener("pointermove", move);
    splitter.addEventListener("pointerup", stop);
    splitter.addEventListener("pointercancel", stop);
  });

  splitter.addEventListener("keydown", (event) => {
    if (event.key === "ArrowUp" || event.key === "ArrowDown") {
      event.preventDefault();
      const held = detail.getBoundingClientRect().height;
      sizeDetail(held + (event.key === "ArrowUp" ? 16 : -16));
    }
  });

  root.append(head, body, splitter, detail);
  document.body.appendChild(root);
  search.focus();

  // --- state ---------------------------------------------------------------------------

  let libraries: ObLibrary[] = [];
  let scope: ObLibrary | null = null;
  let selectedType: string | null = null;

  const typesOf = new Map<string, ObType[]>();
  const membersOf = new Map<string, ObMember[]>();

  const query = (): string => search.value.trim().toLowerCase();

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      event.preventDefault();
      host.close();
    }
  });

  // The window can be hidden and summoned again; coming back, the search box should be
  // ready to type into, exactly as it was on first open.
  window.addEventListener("focus", () => search.focus());

  // --- members pane --------------------------------------------------------------------

  const renderMembers = (): void => {
    membersPane.replaceChildren();
    const wanted = query();
    const mode = scopePick.value;

    // Group mode leaves this pane alone; Object mode filters the selected type's members;
    // All mode searches the members of every loaded type.
    const filterMembers = wanted.length > 0 && mode !== "group";
    const spanTypes = wanted.length > 0 && mode === "all";
    let shown = 0;

    if (scope) {
      const library = scope;
      const types = typesOf.get(library.name) ?? [];
      for (const type of types) {
        if (!spanTypes && type.name !== selectedType) {
          continue;
        }

        // In All mode a group whose own name matches brings its whole membership along,
        // loading it on first sight so the pull is not limited to types already visited.
        // Two characters before loading, or a lone letter fans out across the library.
        const pullWhole = spanTypes && wanted.length >= 2 && type.name.toLowerCase().includes(wanted);

        const held = membersOf.get(`${library.name} ${type.name}`);
        if (!held) {
          if (pullWhole) {
            loadMembers(type.name);
          }

          continue;
        }

        for (const member of [...held].sort(byName)) {
          if (filterMembers && !pullWhole && !member.name.toLowerCase().includes(wanted)) {
            continue;
          }

          shown++;
          const item = document.createElement("div");
          item.className = "objbrowser-row";
          item.setAttribute("role", "option");
          item.tabIndex = 0;

          const glyph = document.createElement("span");
          glyph.className = `codicon codicon-${memberIcon(member.kind)}`;
          glyph.setAttribute("aria-hidden", "true");

          const label = document.createElement("span");
          label.className = "objbrowser-name";
          label.textContent = member.name;

          const trailing = document.createElement("span");
          trailing.className = "objbrowser-context";
          trailing.textContent = spanTypes ? type.name : member.kind;

          item.append(glyph, label, trailing);

          const describe = (): void => {
            const where = library.kind === "project" && member.line > 0 ? `, line ${member.line}` : "";
            setDetail(member.signature, `Member of ${library.name}.${type.name}${where}`, member.description);
          };

          item.addEventListener("click", () => {
            describe();
            for (const other of membersPane.querySelectorAll(".objbrowser-row.selected")) {
              other.classList.remove("selected");
            }
            item.classList.add("selected");
          });

          if (library.kind === "project" && member.line > 0) {
            const jump = (): void => host.navigate(type.name, member.line, library.name);
            item.addEventListener("dblclick", jump);
            item.addEventListener("keydown", (event) => {
              if (event.key === "Enter") {
                event.preventDefault();
                jump();
              }
            });
          }

          membersPane.appendChild(item);
        }
      }
    }

    if (shown === 0) {
      const empty = document.createElement("div");
      empty.className = "objbrowser-empty";
      empty.textContent = !selectedType && !spanTypes
        ? "Pick a type on the left."
        : filterMembers
          ? "Nothing matches among the loaded members."
          : "No members here.";
      membersPane.appendChild(empty);
    }
  };

  // --- types pane ----------------------------------------------------------------------

  /** Requests in flight, so a re-render mid-flight does not ask twice. */
  const loadingMembers = new Set<string>();

  const loadMembers = (typeName: string): void => {
    if (!scope) {
      return;
    }

    const key = `${scope.name} ${typeName}`;
    if (membersOf.has(key) || loadingMembers.has(key)) {
      return;
    }

    loadingMembers.add(key);
    const library = scope;
    void host.members(library.name, typeName).then((rows) => {
      loadingMembers.delete(key);
      if (rows) {
        membersOf.set(key, rows);

        // Both panes: an arrival can add member matches to the types pane in All mode.
        renderTypes();
        renderMembers();
      }
    });
  };

  const pickType = (name: string): void => {
    selectedType = name;
    loadMembers(name);
    renderTypes();
    renderMembers();
  };

  const renderTypes = (): void => {
    typesPane.replaceChildren();
    const wanted = query();
    const mode = scopePick.value;

    if (!scope) {
      return;
    }

    const types = typesOf.get(scope.name);
    if (!types) {
      const loading = document.createElement("div");
      loading.className = "objbrowser-empty";
      loading.textContent = "Loading...";
      typesPane.appendChild(loading);
      return;
    }

    // Object mode leaves this pane alone. Group mode filters it by name; All mode also
    // keeps a type whose loaded members match, so the other pane has something to show.
    for (const type of [...types].sort(byName)) {
      if (wanted.length > 0 && mode !== "object" && !type.name.toLowerCase().includes(wanted)) {
        if (mode === "group") {
          continue;
        }

        const held = membersOf.get(`${scope.name} ${type.name}`);
        if (!held || !held.some((member) => member.name.toLowerCase().includes(wanted))) {
          continue;
        }
      }

      const item = document.createElement("div");
      item.className = "objbrowser-row";
      item.setAttribute("role", "option");
      item.tabIndex = 0;

      if (type.name === selectedType) {
        item.classList.add("selected");
      }

      const glyph = document.createElement("span");
      glyph.className = `codicon codicon-${TYPE_ICONS[type.kind] ?? "symbol-class"}`;
      glyph.setAttribute("aria-hidden", "true");

      const label = document.createElement("span");
      label.className = "objbrowser-name";
      label.textContent = type.name;

      const context = document.createElement("span");
      context.className = "objbrowser-context";
      context.textContent = type.kind;

      item.append(glyph, label, context);
      item.addEventListener("click", () => pickType(type.name));
      item.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          pickType(type.name);
        }
      });

      typesPane.appendChild(item);
    }
  };

  // --- scope switching -----------------------------------------------------------------

  const adoptScope = (next: ObLibrary): void => {
    scope = next;
    selectedType = null;
    setDetail(next.name, next.description);

    if (!typesOf.has(next.name)) {
      void host.types(next.name).then((rows) => {
        if (rows && scope === next) {
          typesOf.set(next.name, rows);
          renderTypes();
          renderMembers();
        }
      });
    }

    renderTypes();
    renderMembers();
  };

  picker.addEventListener("change", () => {
    const picked = libraries[picker.selectedIndex];
    if (picked) {
      adoptScope(picked);
    }
  });

  search.addEventListener("input", () => {
    renderTypes();
    renderMembers();
  });

  scopePick.addEventListener("change", () => {
    renderTypes();
    renderMembers();
  });

  // --- data ----------------------------------------------------------------------------

  void host.libraries().then((rows) => {
    const first = rows?.[0];
    if (!rows || !first) {
      setDetail("No libraries answered. Close this window and open the Object Browser again.");
      return;
    }

    libraries = rows;
    for (const library of rows) {
      const option = document.createElement("option");
      option.textContent = library.kind === "project" ? `${library.name} (project)` : library.name;
      option.title = library.description;
      picker.appendChild(option);
    }

    picker.selectedIndex = 0;
    adoptScope(first);
  });
}
