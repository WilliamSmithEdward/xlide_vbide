/**
 * The code window's Object and Procedure lists, above each code editor (#101).
 *
 * The native editor's two drop-downs: on the left (General) and the objects the module can
 * handle events of - a document module's own object, a class's Class, a form and its controls,
 * every WithEvents variable - and on the right, for (General), (Declarations) and the module's
 * own procedures, or for an object, the events it raises with the ones already handled in bold.
 * Picking a procedure or a handled event goes to it; picking an event not handled yet asks the
 * host to write its handler, as the native editor does. Picking an object goes to its default
 * event's handler the same way. The lists follow the caret, so they always say where it is.
 *
 * The procedures come from the text on the page (procedureat.ts, the same scan the status bar
 * reads); only the objects and their events need the host, because only the engine knows a
 * form's controls and a type library's events.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";
import type { HostCodeObject } from "./bridge.js";
import type { DocumentId } from "./documents.js";
import { procedureAt, scanProcedures, type ProcedureRange } from "./procedureat.js";

export interface CodeBarDeps {
  /** The objects a module can handle events of, given its text on screen; null when no answer came. */
  objects(id: DocumentId, text: string): Promise<HostCodeObject[] | null>;
  /** Go to, or write and go to, an object's handler for an event. */
  eventHandler(id: DocumentId, object: string, event: string): void;
}

const GENERAL = "(General)";
const DECLARATIONS = "(Declarations)";

export class CodeBar {
  readonly root: HTMLElement;
  private readonly objectSelect: HTMLSelectElement;
  private readonly procedureSelect: HTMLSelectElement;
  private id: DocumentId | null = null;
  private objects: HostCodeObject[] = [];
  private procedures: ProcedureRange[] = [];
  private objectsTicket = 0;
  private scanTimer: ReturnType<typeof setTimeout> | null = null;
  private objectsTimer: ReturnType<typeof setTimeout> | null = null;
  /** Set while the lists are being filled, so their change events are not read as picks. */
  private painting = false;

  constructor(private readonly editor: monaco.editor.IStandaloneCodeEditor, private readonly deps: CodeBarDeps) {
    this.root = document.createElement("div");
    this.root.className = "code-bar";
    this.root.hidden = true;

    this.objectSelect = document.createElement("select");
    this.objectSelect.className = "code-bar-object";
    this.objectSelect.setAttribute("aria-label", "Object");
    this.procedureSelect = document.createElement("select");
    this.procedureSelect.className = "code-bar-procedure";
    this.procedureSelect.setAttribute("aria-label", "Procedure");
    this.root.append(this.objectSelect, this.procedureSelect);

    this.objectSelect.addEventListener("change", () => this.objectPicked());
    this.procedureSelect.addEventListener("change", () => this.procedurePicked());

    editor.onDidChangeCursorPosition(() => this.followCaret());
    editor.onDidChangeModelContent(() => this.textChanged());
  }

  /** What the group shows now: a code document lists, anything else hides the bar. */
  show(id: DocumentId | null): void {
    const code = id !== null && !id.face;
    this.root.hidden = !code;
    if (!code) {
      this.id = null;
      return;
    }

    const changed = this.id === null || this.id.module !== id.module || this.id.project !== id.project;
    this.id = id;
    if (changed) {
      this.objects = [];
      this.scan();
      this.paint(GENERAL);
      this.followCaret();
    }
    void this.loadObjects();
  }

  /** The two lists as they stand, for the xlide api. */
  snapshot(): { object: string; procedure: string; objects: string[]; procedures: string[] } {
    return {
      object: this.objectSelect.value,
      procedure: this.procedureSelect.value,
      objects: [...this.objectSelect.options].map((one) => one.value),
      procedures: [...this.procedureSelect.options].map((one) => one.text),
    };
  }

  /** Picks an entry the way a click does, for the xlide api. */
  pick(object: string, procedure?: string): boolean {
    if (![...this.objectSelect.options].some((one) => one.value === object)) {
      return false;
    }
    this.objectSelect.value = object;
    this.objectPicked();
    if (procedure === undefined) {
      return true;
    }
    if (![...this.procedureSelect.options].some((one) => one.value === procedure)) {
      return false;
    }
    this.procedureSelect.value = procedure;
    this.procedurePicked();
    return true;
  }

  private scan(): void {
    this.procedures = scanProcedures(this.editor.getModel()?.getValue() ?? "");
  }

  private async loadObjects(): Promise<void> {
    const id = this.id;
    if (!id) {
      return;
    }
    const ticket = ++this.objectsTicket;
    const objects = await this.deps.objects(id, this.editor.getModel()?.getValue() ?? "");
    if (ticket !== this.objectsTicket || objects === null || this.id !== id) {
      return;
    }
    this.objects = objects;
    this.followCaret();
  }

  private textChanged(): void {
    if (this.root.hidden) {
      return;
    }
    // The procedures after a pause in typing; the objects after a longer one, because only a
    // WithEvents line or a control changes them and each answer is a host round trip.
    if (this.scanTimer) {
      clearTimeout(this.scanTimer);
    }
    this.scanTimer = setTimeout(() => {
      this.scan();
      this.followCaret();
    }, 300);
    if (this.objectsTimer) {
      clearTimeout(this.objectsTimer);
    }
    this.objectsTimer = setTimeout(() => void this.loadObjects(), 1500);
  }

  /** The object a procedure handles an event of, by the longest object name its prefix matches. */
  private handlerOf(name: string): { object: HostCodeObject; event: string } | null {
    let best: { object: HostCodeObject; event: string } | null = null;
    for (const object of this.objects) {
      const prefix = `${object.name}_`;
      if (!name.toLowerCase().startsWith(prefix.toLowerCase())) {
        continue;
      }
      const event = name.slice(prefix.length);
      const known = object.events.find((one) => one.name.toLowerCase() === event.toLowerCase());
      if (known && (!best || object.name.length > best.object.name.length)) {
        best = { object, event: known.name };
      }
    }
    return best;
  }

  /** The lists say where the caret is: the handler's object and event, or (General) and the procedure. */
  private followCaret(): void {
    if (this.root.hidden) {
      return;
    }
    const line = this.editor.getPosition()?.lineNumber ?? 1;
    const at = procedureAt(this.procedures, line);
    const handler = at ? this.handlerOf(at.name) : null;
    if (handler) {
      this.paint(handler.object.name, handler.event);
    } else {
      this.paint(GENERAL, at ? procedureValue(at) : DECLARATIONS);
    }
  }

  /** Fills both lists for an object, selecting one entry of the right one. */
  private paint(objectName: string, selected?: string): void {
    this.painting = true;
    try {
      const objectNames = [GENERAL, ...this.objects.map((one) => one.name)];
      if (!sameOptions(this.objectSelect, objectNames)) {
        this.objectSelect.replaceChildren(...objectNames.map((name) => option(name, name)));
      }
      this.objectSelect.value = objectNames.includes(objectName) ? objectName : GENERAL;

      const handled = new Set(this.procedures.map((one) => one.name.toLowerCase()));
      const object = this.objects.find((one) => one.name === this.objectSelect.value);
      const entries: HTMLOptionElement[] = object
        ? object.events.map((event) => {
          const entry = option(event.name, event.name);
          if (handled.has(`${object.name}_${event.name}`.toLowerCase())) {
            entry.style.fontWeight = "bold";
          }
          return entry;
        })
        : [
          option(DECLARATIONS, DECLARATIONS),
          ...this.procedures
            .filter((one) => !this.handlerOf(one.name))
            .map(procedureValue)
            // By name, as the native list orders them, a property's legs Get, Let, Set.
            .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }))
            .map((value) => option(value, value)),
        ];
      this.procedureSelect.replaceChildren(...entries);
      if (selected !== undefined && entries.some((one) => one.value === selected)) {
        this.procedureSelect.value = selected;
      } else if (!object) {
        this.procedureSelect.value = DECLARATIONS;
      } else {
        this.procedureSelect.selectedIndex = -1;
      }
    } finally {
      this.painting = false;
    }
  }

  private objectPicked(): void {
    if (this.painting) {
      return;
    }
    const name = this.objectSelect.value;
    if (name === GENERAL) {
      this.paint(GENERAL, DECLARATIONS);
      return;
    }
    // As the native list does: an object picked goes to its default event's handler - a sheet's
    // SelectionChange, a button's Click - and writes that handler when the module has none,
    // whatever other handlers of the object stand.
    const object = this.objects.find((one) => one.name === name);
    this.paint(name);
    const event = object?.defaultEvent;
    if (!object || !event || !this.id) {
      return;
    }
    this.procedureSelect.value = event;
    const handler = this.procedures.find((one) => one.name.toLowerCase() === `${object.name}_${event}`.toLowerCase());
    if (handler) {
      this.goTo(handler.header);
      return;
    }
    this.deps.eventHandler(this.id, object.name, event);
  }

  private procedurePicked(): void {
    if (this.painting || !this.id) {
      return;
    }
    const value = this.procedureSelect.value;
    const objectName = this.objectSelect.value;
    if (objectName === GENERAL) {
      if (value === DECLARATIONS) {
        this.goTo(1);
        return;
      }
      const procedure = this.procedures.find((one) => procedureValue(one) === value);
      if (procedure) {
        this.goTo(procedure.header);
      }
      return;
    }

    const handler = this.procedures.find((one) => one.name.toLowerCase() === `${objectName}_${value}`.toLowerCase());
    if (handler) {
      this.goTo(handler.header);
      return;
    }
    this.deps.eventHandler(this.id, objectName, value);
  }

  /** The caret onto a procedure's body, as the native lists leave it. */
  private goTo(header: number): void {
    const model = this.editor.getModel();
    if (!model) {
      return;
    }
    const line = Math.min(header + 1, model.getLineCount());
    this.editor.setPosition({ lineNumber: line, column: model.getLineMaxColumn(line) });
    this.editor.revealLineInCenterIfOutsideViewport(header);
    this.editor.focus();
  }
}

/**
 * A procedure in the right-hand list, as the native list writes it: the name, and a property's
 * leg after it, `Speed [PropertyGet]`, because the legs share a name.
 */
function procedureValue(procedure: ProcedureRange): string {
  return procedure.kind.startsWith("Property ")
    ? `${procedure.name} [${procedure.kind.replace("Property ", "Property")}]`
    : procedure.name;
}

function option(value: string, label: string): HTMLOptionElement {
  const entry = document.createElement("option");
  entry.value = value;
  entry.textContent = label;
  return entry;
}

function sameOptions(select: HTMLSelectElement, values: string[]): boolean {
  return select.options.length === values.length
    && [...select.options].every((one, at) => one.value === values[at]);
}
