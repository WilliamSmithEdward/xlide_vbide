/*
 * The lenses, drawn: "Run test" and "Debug test" above a test, and a count of where it is used
 * above every procedure.
 *
 * What goes where is lensplan.ts, a pure function of the text and the runner's list. This file
 * is the part that needs an editor: it hands the plan to the editor's own CodeLens feature, asks
 * the host for a count when a lens comes on screen, and carries a click to the thing it names.
 *
 * A COUNT IS ASKED FOR ONLY WHEN ITS LENS IS SHOWN. The editor resolves a lens when it scrolls
 * into view and not before, so a 65,000-line module costs what its visible screen costs. Each
 * answer is kept for as long as it can still be right: until this module's text moves, or until
 * the analysis says something moved anywhere in the project - a call added in another module
 * changes this one's counts without touching a character of it.
 *
 * A COUNT IS FOR THE MODULE IT WAS ASKED OF, BY NAME. A split shows two modules at once and the
 * editor draws lenses in both; asking about "the module the host is showing" would answer for
 * whichever group last had the focus, and put one module's counts over another's procedures.
 *
 * A CLICK OPENS XLIDE'S OWN REFERENCES LIST, not the editor's peek: the list draws a use in a
 * module nobody has opened, which the peek cannot, and that use is the one worth being shown.
 *
 * Nothing is drawn over a past version or a form's text view: the first is not the module as it
 * stands, and neither is a place a test is run from.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";
import { docKeyOf, type DocumentId } from "./documents.js";
import {
  DEBUG_TEST_TITLE, RUN_TEST_TITLE, planLenses, referencesTitle, type PlannedLens,
} from "./lensplan.js";
import type { ProcedureRange } from "./procedureat.js";

/** A test as the runner lists it: enough to find its procedure and to ask for it to be run. */
export interface LensTest {
  id: string;
  file: string;
  module: string;
  procedure: string;
}

export interface LensDeps {
  /** The developer's switch. Off draws nothing and asks nothing. */
  enabled(): boolean;
  idOf(model: monaco.editor.ITextModel): DocumentId | null;
  /** The procedures of the model's text as it stands; the caller keeps them per text version. */
  procedures(model: monaco.editor.ITextModel): readonly ProcedureRange[];
  tests(): readonly LensTest[];
  /** How many places use the symbol at an offset, its declaration aside. Null when nobody could say. */
  countReferences(model: monaco.editor.ITextModel, offset: number): Promise<number | null>;
  runTest(id: string, file: string): void;
  debugTest(id: string, file: string): void;
  showReferences(model: monaco.editor.ITextModel, position: monaco.Position): void;
}

const RUN_TEST_COMMAND = "xlide.lens.runTest";
const DEBUG_TEST_COMMAND = "xlide.lens.debugTest";
const SHOW_REFERENCES_COMMAND = "xlide.lens.showReferences";

/** A lens as the editor holds it, with the plan it came from so its count can be asked for. */
interface PlannedCodeLens extends monaco.languages.CodeLens {
  planned: PlannedLens;
}

/** One lens as a driver reads it: where it is, what it is for, and what it says once resolved. */
export interface LensReading {
  line: number;
  kind: string;
  procedure: string;
  title: string | null;
}

/** What the editor actually drew above one line. */
export interface LensDrawn {
  line: number | null;
  titles: string[];
}

export class Lenses {
  private readonly refresh = new monaco.Emitter<monaco.languages.CodeLensProvider>();
  private refreshTimer: ReturnType<typeof setTimeout> | undefined;
  private provider: monaco.languages.CodeLensProvider | null = null;

  /** Moves whenever a count could have changed without this module's text changing. */
  private epoch = 0;
  private readonly counts = new Map<string, number | null>();
  private readonly planned = new WeakMap<monaco.editor.ITextModel, PlannedCodeLens[]>();

  /** How many counts were asked of the host, for the cost to be read rather than guessed. */
  private asked = 0;

  constructor(private readonly deps: LensDeps) {}

  install(languageId: string): void {
    monaco.editor.registerCommand(RUN_TEST_COMMAND, (_accessor, id: string, file: string) => {
      this.deps.runTest(id, file);
    });
    monaco.editor.registerCommand(DEBUG_TEST_COMMAND, (_accessor, id: string, file: string) => {
      this.deps.debugTest(id, file);
    });
    monaco.editor.registerCommand(SHOW_REFERENCES_COMMAND, (_accessor, uri: string, line: number, column: number) => {
      const model = monaco.editor.getModel(monaco.Uri.parse(uri));
      if (model) {
        this.deps.showReferences(model, new monaco.Position(line, column));
      }
    });

    this.provider = {
      onDidChange: this.refresh.event,
      provideCodeLenses: (model) => ({ lenses: this.plan(model), dispose: () => undefined }),
      resolveCodeLens: (model, lens) => this.resolve(model, lens as PlannedCodeLens),
    };
    monaco.languages.registerCodeLensProvider(languageId, this.provider);
  }

  /**
   * The analysis moved: a use was added or taken away somewhere in the project, so any count may
   * be stale. Counts already held are dropped and the editor is told to ask again.
   */
  analysisMoved(): void {
    this.epoch++;
    this.counts.clear();
    this.askAgain();
  }

  /**
   * The runner sent its list. It sends the whole list on every change, a landing result
   * included, and only WHICH procedures are tests matters here - so the editor is told to plan
   * again only when that has changed, and the counts are left alone either way: a test passing
   * moves no count, and re-asking every count on screen after each result of a long run would be
   * this feature costing the most exactly while the runner is busiest.
   */
  testsMoved(): void {
    const listed = this.deps.tests()
      .map((test) => `${test.file}\0${test.module}\0${test.procedure}\0${test.id}`)
      .join("\n")
      .toLowerCase();
    if (listed !== this.testsListed) {
      this.testsListed = listed;
      this.askAgain();
    }
  }

  /** The developer's switch moved. */
  switched(): void {
    this.askAgain();
  }

  private testsListed = "";

  /** Once, after the burst: an analysis pass announces itself several times. */
  private askAgain(): void {
    clearTimeout(this.refreshTimer);
    this.refreshTimer = setTimeout(() => {
      if (this.provider) {
        this.refresh.fire(this.provider);
      }
    }, 250);
  }

  private plan(model: monaco.editor.ITextModel): PlannedCodeLens[] {
    const id = this.deps.enabled() ? this.deps.idOf(model) : null;
    if (!id || id.face) {
      this.planned.delete(model);
      return [];
    }

    const tests = this.deps.tests()
      .filter((test) => test.module.toLowerCase() === id.module.toLowerCase()
        && (id.project === null || test.file.toLowerCase() === id.project.toLowerCase()));
    const fileOf = new Map(tests.map((test) => [test.id, test.file]));

    const lenses = planLenses(this.deps.procedures(model), (line) => model.getLineContent(line), tests)
      .map((planned): PlannedCodeLens => {
        const range = new monaco.Range(planned.line, 1, planned.line, 1);
        if (planned.kind === "references") {
          // No command yet: the count is the command's title, and it is asked for on screen.
          return { range, planned };
        }

        const run = planned.kind === "runTest";
        return {
          range,
          planned,
          command: {
            id: run ? RUN_TEST_COMMAND : DEBUG_TEST_COMMAND,
            title: run ? RUN_TEST_TITLE : DEBUG_TEST_TITLE,
            tooltip: run
              ? `Run ${planned.procedure} and show the outcome in the Tests pane`
              : `Run ${planned.procedure} under the debugger, stopping at its first line`,
            arguments: [planned.test, fileOf.get(planned.test ?? "") ?? ""],
          },
        };
      });

    this.planned.set(model, lenses);
    return lenses;
  }

  private async resolve(model: monaco.editor.ITextModel, lens: PlannedCodeLens): Promise<PlannedCodeLens> {
    const planned = lens.planned;
    if (!planned || planned.kind !== "references" || lens.command) {
      return lens;
    }

    // The plan is of the text as it stood. If the name is no longer where the plan put it, the
    // text has moved and a fresh plan is already on its way; an answer for this one would be a
    // count for whatever now stands at that offset.
    if (planned.line > model.getLineCount()
      || !model.getLineContent(planned.line).startsWith(planned.procedure, planned.column - 1)) {
      return lens;
    }

    const id = this.deps.idOf(model);
    const askedAt = model.getVersionId();
    const key = `${id ? docKeyOf(id.module, id.project) : model.uri.toString()}@${askedAt}#${this.epoch}`
      + `:${planned.line}:${planned.procedure.toLowerCase()}`;

    let count = this.counts.get(key);
    if (count === undefined) {
      this.asked++;
      count = await this.deps.countReferences(
        model, model.getOffsetAt({ lineNumber: planned.line, column: planned.column }));

      if (model.isDisposed() || model.getVersionId() !== askedAt) {
        return lens;
      }

      // A count nobody could give is not kept: the next time the lens is shown, it is asked again.
      if (count !== null) {
        this.counts.set(key, count);
      }
    }

    lens.command = {
      id: SHOW_REFERENCES_COMMAND,
      title: referencesTitle(count),
      tooltip: `Show where ${planned.procedure} is used`,
      arguments: [model.uri.toString(), planned.line, planned.column],
    };
    return lens;
  }

  // ---- what a driver reads and presses --------------------------------------------------------

  /** The plan for an editor's module, each lens with what it says once it has been resolved. */
  readingOf(editor: monaco.editor.ICodeEditor): LensReading[] {
    // With the switch off the editor stops asking, so the last plan is never replaced by an
    // empty one. It is not a plan any more, and is not read as one.
    const model = this.deps.enabled() ? editor.getModel() : null;
    return (model ? this.planned.get(model) ?? [] : []).map((lens) => ({
      line: lens.planned.line,
      kind: lens.planned.kind,
      procedure: lens.planned.procedure,
      title: lens.command?.title ?? null,
    }));
  }

  /**
   * What the editor drew, read off the page: one entry per line of lenses on screen, top to
   * bottom, with the words on it. The line is worked out from where the lenses sit - a lens is
   * drawn in the gap directly above the line it belongs to - and is null when no planned line
   * sits under one, which is itself worth seeing.
   */
  drawnIn(editor: monaco.editor.ICodeEditor): LensDrawn[] {
    return this.decorationsIn(editor).map(({ line, node }) => ({
      line,
      titles: [...node.querySelectorAll("a")].map((anchor) => anchor.textContent ?? ""),
    }));
  }

  /**
   * Presses a drawn lens the way a pointer does: on the anchor itself, through the editor's own
   * mouse handling, so what runs is whatever the editor wired to that anchor. Answers what it
   * pressed, or why it could not.
   */
  press(editor: monaco.editor.ICodeEditor, line: number, title: string): { did: boolean; detail: string } {
    const row = this.decorationsIn(editor).find((one) => one.line === line);
    if (!row) {
      return { did: false, detail: `no lenses are drawn above line ${line}` };
    }

    const anchors = [...row.node.querySelectorAll<HTMLElement>("a")];
    const anchor = anchors.find((one) => (one.textContent ?? "") === title)
      ?? anchors.find((one) => (one.textContent ?? "").toLowerCase().includes(title.toLowerCase()));
    if (!anchor) {
      return { did: false, detail: `line ${line} shows ${anchors.map((one) => `"${one.textContent}"`).join(", ")}, not "${title}"` };
    }

    const box = anchor.getBoundingClientRect();
    const at = { clientX: box.left + box.width / 2, clientY: box.top + box.height / 2 };
    const base = { bubbles: true, cancelable: true, view: window, button: 0, ...at };
    anchor.dispatchEvent(new PointerEvent("pointerdown", { ...base, buttons: 1, pointerId: 1, isPrimary: true, pointerType: "mouse" }));
    anchor.dispatchEvent(new MouseEvent("mousedown", { ...base, buttons: 1, detail: 1 }));
    anchor.dispatchEvent(new PointerEvent("pointerup", { ...base, buttons: 0, pointerId: 1, isPrimary: true, pointerType: "mouse" }));
    anchor.dispatchEvent(new MouseEvent("mouseup", { ...base, buttons: 0, detail: 1 }));
    return { did: true, detail: `pressed "${anchor.textContent}" above line ${line}` };
  }

  /** How many counts have been asked of the host since the page loaded. */
  countsAsked(): number {
    return this.asked;
  }

  private decorationsIn(editor: monaco.editor.ICodeEditor): { line: number | null; node: HTMLElement }[] {
    const host = editor.getDomNode();
    const model = editor.getModel();
    if (!host || !model) {
      return [];
    }

    const origin = host.getBoundingClientRect();
    const lineHeight = editor.getOption(monaco.editor.EditorOption.lineHeight);
    const lines = [...new Set((this.planned.get(model) ?? []).map((lens) => lens.planned.line))]
      .map((line) => ({ line, top: editor.getScrolledVisiblePosition({ lineNumber: line, column: 1 })?.top ?? null }))
      .filter((one): one is { line: number; top: number } => one.top !== null);

    return [...host.querySelectorAll<HTMLElement>(".codelens-decoration")]
      .filter((node) => node.getBoundingClientRect().height > 0 && node.querySelector("a") !== null)
      .map((node) => ({ node, top: node.getBoundingClientRect().top - origin.top }))
      .sort((a, b) => a.top - b.top)
      .map(({ node, top }) => ({
        node,
        // The line whose top edge is the nearest one below this lens, within two lines of it.
        line: lines
          .filter((one) => one.top > top && one.top - top <= lineHeight * 2)
          .sort((a, b) => a.top - b.top)[0]?.line ?? null,
      }));
  }
}
