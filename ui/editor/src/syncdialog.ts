/*
 * Import and export, in one window.
 *
 * The companion editor spells this as three commands and a preview panel. Here it is one surface
 * with a direction on it, because the two directions ask the same questions, namely which folder, which
 * modules, and what happens to the ones only one side has, and answering them in two places is
 * how the two ends of a round trip drift apart.
 *
 * NOTHING IS DECIDED HERE. The host works out the plan and the host carries it out; this draws the
 * plan and sends back the rows that are ticked. The xlide api's `sync` route reaches the very same
 * call in the host, so a plan read by a harness is the plan drawn here, and Apply leaves the
 * project exactly where the api would have left it. That is the whole reason the dialog is thin.
 */

import { drawDiffRows, type SyncDiffLine } from "./diffview.js";
import { openModal } from "./modal.js";

interface SyncItem {
  id: string;
  module: string;
  kind: string;
  file: string;
  status: string;
  checked: boolean;
  detail: string;
  warning?: string;
  inProject: boolean;
  inFolder: boolean;
  cannotCreate: boolean;
  leftTitle: string;
  rightTitle: string;
  diff: SyncDiffLine[];
  diffWithHeaders: SyncDiffLine[];
}

interface SyncPlan {
  direction: "export" | "import";
  project: string;
  projectId: string;
  folder: string;
  mode: string;
  items: SyncItem[];
  warnings: string[];
  error?: string;
}

/** How the host is reached. One function, because there is only one kind of request. */
export type SyncRequest = (args: Record<string, string>, body?: string) => Promise<Record<string, unknown>>;

/** The projects this session has open, and the one the developer is looking at. */
export interface OpenProjects {
  names: string[];
  current: string | null;
}

/**
 * The OPEN dialog, readable and drivable for the dev surface. Everything a driver does goes
 * through the dialog's own controls - the checkbox's change, the button's click, the input's
 * change event - so a driven Apply is the developer's Apply, which is the property module-sync
 * exists to prove and used to reach by document.querySelector from eight files away.
 */
export interface SyncDialogProbe {
  state(): {
    direction: string;
    folder: string;
    mode: string;
    busy: boolean;
    status: string;
    /** The project this dialog is acting on, and the ones it could be pointed at. */
    project: string;
    projects: string[];
    rows: { file: string; status: string; detail: string; ticked: boolean; actionable: boolean }[];
  };
  /** Presses a named control: apply, close, all, none, export, import. False when unknown. */
  press(control: string): boolean;
  /** Ticks or unticks a row by file name, through its checkbox. False when no such row. */
  tick(file: string, on: boolean): boolean;
  /** Types a folder path, through the input's own change event. */
  setFolder(path: string): void;
  /** Points the dialog at another open project, through the select's own change event. */
  chooseProject(name: string): boolean;
}

let liveDialog: SyncDialogProbe | null = null;

/** The open dialog's probe, or null when none is open. */
export const syncDialogProbe = (): SyncDialogProbe | null => liveDialog;

/** The word each status is drawn with, and the tone it is drawn in. */
const STATUS_TONE: Record<string, { label: string; tone: string }> = {
  "will-create": { label: "new", tone: "create" },
  "will-write": { label: "overwrite", tone: "write" },
  "will-update": { label: "update", tone: "write" },
  "will-remove": { label: "delete", tone: "remove" },
  "skipping-import": { label: "skipped", tone: "muted" },
  "read-error": { label: "unreadable", tone: "error" },
  unchanged: { label: "same", tone: "muted" },
};

/**
 * Opens the dialog. One at a time: opening while open focuses the one that exists, the way the
 * settings card behaves.
 */
export function openSyncDialog(
  request: SyncRequest,
  closed: () => void,
  session: OpenProjects = { names: [], current: null },
): void {
  if (document.getElementById("sync-backdrop")) {
    document.querySelector<HTMLElement>("#sync-folder")?.focus();
    return;
  }

  let direction: "export" | "import" = "export";
  let folder = "";
  /*
   * WHICH PROJECT, SENT WITH EVERY REQUEST.
   *
   * The dialog used to name none, and the host filled the gap with two different fallbacks: the
   * plan's identity from the SHOWN project and its modules from the editor's ACTIVE one. With two
   * workbooks open those are routinely different - measured 2026-08-21 with nothing contrived,
   * seconds after opening two fixtures side by side, a plan titled DebugFixture.xlsm whose rows
   * were TwinFixture's six modules. Applying it would have exported one workbook's modules into
   * the other's folder, or imported that folder over the other workbook's code.
   *
   * Naming it here closes the other half of that too: the plan on screen and the Apply that
   * carries it out cannot drift apart while the dialog is open, whatever the editor does behind
   * it. An empty string means the session offered nothing to choose from, and the host falls back
   * to the shown project exactly as before.
   */
  let project = session.current ?? session.names[0] ?? "";
  let plan: SyncPlan | null = null;
  let selectedId: string | null = null;
  let showHeaders = false;
  let busy = false;
  /** Something asked for a refresh while one was already running. */
  let wanted = false;
  const ticked = new Set<string>();

  const { card, dismiss } = openModal({
    backdropId: "sync-backdrop",
    cardId: "sync-card",
    label: "Import and export modules",
    closed: () => {
      liveDialog = null;
      closed();
    },
  });

  // ---- head -------------------------------------------------------------------------------
  const head = document.createElement("div");
  head.id = "sync-head";

  const title = document.createElement("span");
  title.id = "sync-title";
  title.textContent = "Import and export";

  const close = document.createElement("button");
  close.type = "button";
  close.id = "sync-close";
  close.title = "Close (Esc)";
  close.setAttribute("aria-label", "Close");
  close.innerHTML = '<span class="codicon codicon-close" aria-hidden="true"></span>';

  head.append(title, close);

  // ---- direction --------------------------------------------------------------------------
  const directions = document.createElement("div");
  directions.className = "sync-directions";
  directions.setAttribute("role", "radiogroup");
  directions.setAttribute("aria-label", "Direction");

  const makeDirection = (value: "export" | "import", icon: string, label: string, hint: string): HTMLButtonElement => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "sync-direction";
    button.dataset.direction = value;
    button.setAttribute("role", "radio");

    const glyph = document.createElement("span");
    glyph.className = `codicon codicon-${icon}`;
    glyph.setAttribute("aria-hidden", "true");

    const text = document.createElement("span");
    text.className = "sync-direction-text";
    const strong = document.createElement("span");
    strong.className = "sync-direction-label";
    strong.textContent = label;
    const small = document.createElement("span");
    small.className = "sync-direction-hint";
    small.textContent = hint;
    text.append(strong, small);

    button.append(glyph, text);
    button.addEventListener("click", () => {
      if (direction !== value) {
        direction = value;
        selectedId = null;
        void refresh();
      }
    });
    return button;
  };

  const exportButton = makeDirection("export", "export", "Export", "the workbook writes the folder");
  const importButton = makeDirection("import", "cloud-download", "Import", "the folder writes the workbook");
  directions.append(exportButton, importButton);

  // ---- folder -----------------------------------------------------------------------------
  // ---- project ----------------------------------------------------------------------------
  //
  // "Project", not "File", though a file is this product's word for a workbook everywhere else
  // (see scopeselect.ts). Here the rows below ARE files - the .bas and .cls on disk - and one
  // dialog cannot use the word for both without the reader having to work out which is meant.
  // Project is what VBA calls it, what the plan's own field is called, and what the tree holds.
  //
  // Hidden with one project open, the same rule the list panes' file select follows: a choice
  // between one thing and nothing is not a choice, and the title names it anyway.
  const projectRow = document.createElement("div");
  projectRow.className = "sync-row";
  projectRow.hidden = session.names.length < 2;

  const projectLabel = document.createElement("label");
  projectLabel.className = "sync-row-label";
  projectLabel.textContent = "Project";
  projectLabel.htmlFor = "sync-project";

  const projectSelect = document.createElement("select");
  projectSelect.id = "sync-project";
  projectSelect.className = "sync-mode";
  for (const name of session.names) {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    projectSelect.appendChild(option);
  }
  if (project) {
    projectSelect.value = project;
  }

  // READ BACK, so the control and the value sent cannot disagree. Assigning a value the list does
  // not hold leaves a select showing its FIRST option and reports nothing - and then the dialog
  // would say one project on screen and name another in every request, which is the exact failure
  // this control was added to end.
  project = projectSelect.value;

  const projectHint = document.createElement("span");
  projectHint.className = "sync-row-hint";
  projectHint.textContent = "The one these modules are imported into, or exported from.";

  projectSelect.addEventListener("change", () => {
    project = projectSelect.value;

    // FORGOTTEN ON PURPOSE. Every project remembers its own folder, and carrying this one's over
    // would offer the new project the old one's folder as though it had chosen it - which is the
    // shape of the mistake this whole control exists to prevent. Cleared, the host answers with
    // what the newly chosen project remembers, and the plan puts it back in the box.
    folder = "";
    folderInput.value = "";
    void refresh();
  });

  projectRow.append(projectLabel, projectSelect, projectHint);

  const folderRow = document.createElement("div");
  folderRow.className = "sync-row";

  const folderLabel = document.createElement("label");
  folderLabel.className = "sync-row-label";
  folderLabel.textContent = "Folder";
  folderLabel.htmlFor = "sync-folder";

  const folderInput = document.createElement("input");
  folderInput.id = "sync-folder";
  folderInput.type = "text";
  folderInput.className = "sync-folder";
  folderInput.spellcheck = false;
  folderInput.placeholder = "Choose a folder to keep the .bas and .cls files in";
  folderInput.addEventListener("change", () => {
    folder = folderInput.value.trim();
    void refresh();
  });

  const browse = document.createElement("button");
  browse.type = "button";
  browse.className = "sync-button";
  browse.textContent = "Browse...";
  browse.addEventListener("click", () => void chooseFolder());

  folderRow.append(folderLabel, folderInput, browse);

  // ---- mode -------------------------------------------------------------------------------
  const modeRow = document.createElement("div");
  modeRow.className = "sync-row";

  const modeLabel = document.createElement("label");
  modeLabel.className = "sync-row-label";
  modeLabel.textContent = "Missing";
  modeLabel.htmlFor = "sync-mode";

  const modeSelect = document.createElement("select");
  modeSelect.id = "sync-mode";
  modeSelect.className = "sync-mode";

  /*
   * THE CHOSEN MODE, REMEMBERED, BECAUSE THE CONTROL CANNOT REMEMBER IT ITSELF.
   *
   * Each direction has its own two options, so drawDirection rebuilds the list every time it runs
   * - and it runs at the top of every refresh. replaceChildren threw the selection away, and the
   * request four lines later read `modeSelect.value` back off the rebuilt list: the first option.
   *
   * So choosing "Delete them" did nothing at all. The plan was asked for with the keeping mode
   * every time, the dropdown snapped back on its own, and the row that would have offered the
   * delete never appeared. Measured 2026-08-09: set to trueUp, dispatch change, read back
   * exportAll, synchronously. The api was never affected, which is why nothing caught it: a caller
   * passes the mode as an argument and never touches this control.
   *
   * Kept per direction rather than as one value, because trueUp means nothing to an import.
   */
  const modeFor: Record<"export" | "import", string> = {
    export: "exportAll",
    import: "updateOnly",
  };

  modeSelect.addEventListener("change", () => {
    modeFor[direction] = modeSelect.value;

    // The sentence beside it follows the choice immediately, rather than waiting for the plan to
    // come back: a hint that lags the control it explains is the confusion this replaced.
    drawModeHint();
    void refresh();
  });

  const modeHint = document.createElement("span");
  modeHint.className = "sync-row-hint";

  modeRow.append(modeLabel, modeSelect, modeHint);

  // ---- list and diff ----------------------------------------------------------------------
  const body = document.createElement("div");
  body.id = "sync-body";

  const listSide = document.createElement("div");
  listSide.id = "sync-list-side";

  const listHead = document.createElement("div");
  listHead.className = "sync-list-head";

  const selectAll = document.createElement("button");
  selectAll.type = "button";
  selectAll.className = "sync-link";
  selectAll.textContent = "All";
  selectAll.addEventListener("click", () => {
    for (const item of plan?.items ?? []) {
      if (actionable(item)) {
        ticked.add(item.id);
      }
    }

    drawList();
  });

  const selectNone = document.createElement("button");
  selectNone.type = "button";
  selectNone.className = "sync-link";
  selectNone.textContent = "None";
  selectNone.addEventListener("click", () => {
    ticked.clear();
    drawList();
  });

  const counts = document.createElement("span");
  counts.className = "sync-counts";

  listHead.append(counts, selectAll, selectNone);

  const list = document.createElement("div");
  list.id = "sync-list";
  list.setAttribute("role", "listbox");
  list.setAttribute("aria-label", "Modules");

  // A LISTBOX HAS TO BEHAVE LIKE ONE.
  //
  // The rows were drawn with role="option" inside role="listbox" and none of the behaviour that
  // role promises: nothing here could be reached, moved through or chosen without a mouse, and
  // anything reading the page aloud was told it was a listbox regardless. Declaring the role and
  // not the behaviour is worse than declaring neither, because it announces something untrue
  // (2026-08-09).
  //
  // Focus stays on the LIST and the current row is named by aria-activedescendant, which is the
  // pattern for a list whose rows also carry a checkbox: the tick is a Tab stop in its own right,
  // so the rows must not compete for the same stops.
  list.tabIndex = 0;
  list.addEventListener("keydown", (event) => {
    const items = plan?.items ?? [];
    if (items.length === 0) {
      return;
    }

    const at = Math.max(0, items.findIndex((item) => item.id === selectedId));

    if (event.key === " ") {
      // Space ticks the current row, the way it ticks a focused checkbox.
      const item = items[at];
      if (item && actionable(item)) {
        event.preventDefault();
        if (ticked.has(item.id)) {
          ticked.delete(item.id);
        } else {
          ticked.add(item.id);
        }

        drawList();
      }

      return;
    }

    const next = event.key === "ArrowDown" ? Math.min(items.length - 1, at + 1)
      : event.key === "ArrowUp" ? Math.max(0, at - 1)
      : event.key === "Home" ? 0
      : event.key === "End" ? items.length - 1
      : -1;

    if (next < 0) {
      return;
    }

    event.preventDefault();
    if (items[next] && items[next].id !== selectedId) {
      selectedId = items[next].id;
      drawList();
      drawDiff();
    }
  });

  listSide.append(listHead, list);

  const diffSide = document.createElement("div");
  diffSide.id = "sync-diff-side";

  const diffHead = document.createElement("div");
  diffHead.className = "sync-diff-head";

  const leftTitle = document.createElement("span");
  leftTitle.className = "sync-diff-title";
  const rightTitle = document.createElement("span");
  rightTitle.className = "sync-diff-title";

  const headerToggle = document.createElement("label");
  headerToggle.className = "sync-header-toggle";
  const headerCheck = document.createElement("input");
  headerCheck.type = "checkbox";
  headerCheck.addEventListener("change", () => {
    showHeaders = headerCheck.checked;
    drawDiff();
  });
  const headerText = document.createElement("span");
  headerText.textContent = "Attributes";
  headerToggle.append(headerCheck, headerText);
  headerToggle.title =
    "The VERSION and Attribute lines a file carries. They are what makes an exported file come back "
    + "as the same kind of module, and nobody edits them, so they are out of the way by default.";

  diffHead.append(leftTitle, rightTitle, headerToggle);

  const diff = document.createElement("div");
  diff.id = "sync-diff";

  diffSide.append(diffHead, diff);
  body.append(listSide, diffSide);

  // ---- foot -------------------------------------------------------------------------------
  const foot = document.createElement("div");
  foot.id = "sync-foot";

  const status = document.createElement("div");
  status.id = "sync-status";
  status.setAttribute("role", "status");

  const cancel = document.createElement("button");
  cancel.type = "button";
  cancel.className = "sync-button";
  cancel.textContent = "Close";
  cancel.addEventListener("click", () => dismiss());

  const apply = document.createElement("button");
  apply.type = "button";
  apply.className = "sync-button sync-primary";
  apply.textContent = "Apply";
  apply.addEventListener("click", () => void applyPlan());

  foot.append(status, cancel, apply);

  card.append(head, directions, projectRow, folderRow, modeRow, body, foot);

  // ---- behaviour --------------------------------------------------------------------------
  const actionable = (item: SyncItem): boolean =>
    item.status !== "unchanged" && item.status !== "skipping-import" && item.status !== "read-error";

  const say = (text: string, tone: "" | "error" | "good" = ""): void => {
    status.textContent = text;
    status.dataset.tone = tone;
  };

  function drawDirection(): void {
    for (const button of [exportButton, importButton]) {
      const active = button.dataset.direction === direction;
      button.classList.toggle("active", active);
      button.setAttribute("aria-checked", String(active));
    }

    modeSelect.replaceChildren();
    const choices = direction === "export"
      ? [
        { value: "exportAll", label: "Keep them" },
        { value: "trueUp", label: "Delete them" },
      ]
      : [
        { value: "updateOnly", label: "Keep them" },
        { value: "trueUpStandardClass", label: "Delete them" },
      ];

    for (const choice of choices) {
      const option = document.createElement("option");
      option.value = choice.value;
      option.textContent = choice.label;
      modeSelect.appendChild(option);
    }

    // Put the choice back on the list that was just rebuilt under it.
    modeSelect.value = modeFor[direction];

    // PLURAL, because this is a rule about a class of things and not about one file. "A file with
    // no module" reads as though the dialog has one in mind and is asking what to do with it.
    modeLabel.textContent = direction === "export"
      ? "Files with no module"
      : "Modules with no file";

    drawModeHint();
  }

  /**
   * What the CHOSEN option does, in the present tense.
   *
   * This said "Deleting makes the folder match the project exactly" whatever was selected, so a
   * reader who had chosen to keep their files was being told about deleting, next to a control
   * that said Keep. The one sentence beside a choice has to be about the choice that is made, or
   * it reads as a warning about what is happening rather than a note about what is not.
   */
  function drawModeHint(): void {
    const deleting = modeSelect.value === "trueUp" || modeSelect.value === "trueUpStandardClass";

    modeHint.textContent = direction === "export"
      ? deleting
        ? "Files in the folder that no longer match a module are deleted, so the folder ends up "
          + "matching the project exactly."
        : "Files in the folder that no longer match a module are left where they are."
      : deleting
        ? "Modules with no file are removed from the workbook. Only standard and class modules are "
          + "ever removed; sheets, the workbook and forms are always left alone."
        : "Modules with no file are left in the workbook.";
  }

  function drawList(): void {
    list.replaceChildren();

    const items = plan?.items ?? [];
    if (items.length === 0) {
      const empty = document.createElement("div");
      empty.className = "sync-empty";
      empty.textContent = direction === "export"
        ? "This project has no modules to export."
        : "There are no .bas or .cls files in this folder.";
      list.appendChild(empty);
    }

    items.forEach((item, index) => {
      const row = document.createElement("div");
      row.className = "sync-item";
      row.dataset.id = item.id;
      row.id = `sync-row-${index}`;
      row.setAttribute("role", "option");
      row.setAttribute("aria-selected", String(item.id === selectedId));
      row.classList.toggle("selected", item.id === selectedId);

      const tick = document.createElement("input");
      tick.type = "checkbox";
      tick.className = "sync-tick";
      tick.checked = ticked.has(item.id);
      tick.disabled = !actionable(item);
      tick.setAttribute("aria-label", `${item.detail}: ${item.file}`);
      tick.addEventListener("click", (event) => event.stopPropagation());
      tick.addEventListener("change", () => {
        if (tick.checked) {
          ticked.add(item.id);
        } else {
          ticked.delete(item.id);
        }

        drawCounts();
      });

      const text = document.createElement("div");
      text.className = "sync-item-text";

      const name = document.createElement("span");
      name.className = "sync-item-name";
      name.textContent = item.file;

      const detail = document.createElement("span");
      detail.className = "sync-item-detail";
      detail.textContent = item.detail;

      text.append(name, detail);

      const chip = document.createElement("span");
      const tone = STATUS_TONE[item.status] ?? { label: item.status, tone: "muted" };
      chip.className = "sync-chip";
      chip.dataset.tone = tone.tone;
      chip.textContent = tone.label;

      row.append(tick, text, chip);

      if (item.warning) {
        const warn = document.createElement("span");
        warn.className = "codicon codicon-warning sync-item-warning";
        warn.title = item.warning;
        row.appendChild(warn);
      }

      row.addEventListener("click", () => {
        selectedId = item.id;
        drawList();
        drawDiff();
      });

      list.appendChild(row);
    });

    // The row the keyboard is on, named for anything reading the list aloud, and kept in view.
    const current = items.findIndex((item) => item.id === selectedId);
    if (current >= 0) {
      list.setAttribute("aria-activedescendant", `sync-row-${current}`);
      list.children[current]?.scrollIntoView({ block: "nearest" });
    } else {
      list.removeAttribute("aria-activedescendant");
    }

    drawCounts();
  }

  function drawCounts(): void {
    const items = plan?.items ?? [];
    const willAct = items.filter((item) => ticked.has(item.id)).length;
    counts.textContent = `${items.length} module${items.length === 1 ? "" : "s"}, ${willAct} selected`;
    apply.disabled = busy || willAct === 0;
  }

  function drawDiff(): void {
    diff.replaceChildren();

    const item = (plan?.items ?? []).find((candidate) => candidate.id === selectedId);
    if (!item) {
      leftTitle.textContent = "";
      rightTitle.textContent = "";
      const hint = document.createElement("div");
      hint.className = "sync-empty";
      hint.textContent = "Choose a module to see what would change.";
      diff.appendChild(hint);
      return;
    }

    leftTitle.textContent = item.leftTitle;
    rightTitle.textContent = item.rightTitle;

    const lines = showHeaders ? item.diffWithHeaders : item.diff;
    if (lines.length === 0) {
      const hint = document.createElement("div");
      hint.className = "sync-empty";
      hint.textContent = "Nothing on either side.";
      diff.appendChild(hint);
      return;
    }

    // Drawn by the shared renderer, which the Changes pane uses too: one idea of what a removed
    // line looks like, in one place.
    drawDiffRows(diff, lines, "sync");
  }

  function drawWarnings(): void {
    const warnings = plan?.warnings ?? [];
    if (warnings.length > 0) {
      say(warnings.join("  "), "error");
    }
  }

  async function refresh(): Promise<void> {
    drawDirection();

    // A REFRESH ASKED FOR WHILE ONE IS IN FLIGHT IS REMEMBERED, NOT DROPPED.
    //
    // This used to return here, which reads as harmless and is not: pressing Import while the
    // export plan is still being worked out changed the direction, threw the refresh away, and
    // left the export's rows on screen under the Import heading. The developer is then looking at
    // a list that does not match the direction they chose, and Apply would do what the list says.
    //
    // Only ever one pending: what matters is that the LAST thing asked for is what ends up drawn,
    // not that every intermediate state is rendered on the way.
    if (busy) {
      wanted = true;
      return;
    }

    busy = true;
    apply.disabled = true;
    say("Working out what would change...");

    const answer = await request({
      direction,
      ...(project ? { project } : {}),
      ...(folder ? { folder } : {}),
      mode: modeSelect.value,
    });

    busy = false;

    if (typeof answer.error === "string") {
      plan = null;
      drawList();
      drawDiff();
      say(answer.error, "error");
      if (wanted) {
        wanted = false;
        await refresh();
      }

      return;
    }

    plan = answer as unknown as SyncPlan;
    folder = plan.folder;
    folderInput.value = plan.folder;
    title.textContent = `Import and export: ${plan.project}`;

    // The plan's own ticks are the starting point, and they are re-read on every refresh: a row
    // that has become "unchanged" since the last look should not stay ticked from before.
    ticked.clear();
    for (const item of plan.items) {
      if (item.checked) {
        ticked.add(item.id);
      }
    }

    if (!plan.items.some((item) => item.id === selectedId)) {
      selectedId = plan.items[0]?.id ?? null;
    }

    drawList();
    drawDiff();

    const willAct = plan.items.filter((item) => ticked.has(item.id)).length;
    say(willAct === 0 ? "Everything already matches." : `${willAct} of ${plan.items.length} would change.`);
    drawWarnings();

    if (wanted) {
      wanted = false;
      await refresh();
    }
  }

  async function chooseFolder(): Promise<void> {
    say("Waiting for a folder...");
    const answer = await request({ action: "browse", direction, ...(folder ? { folder } : {}) });
    if (typeof answer.folder === "string" && answer.folder.length > 0) {
      folder = answer.folder;
      folderInput.value = folder;
    }

    await refresh();
  }

  async function applyPlan(): Promise<void> {
    if (busy || ticked.size === 0) {
      return;
    }

    busy = true;
    apply.disabled = true;
    say(direction === "export" ? "Writing the folder..." : "Writing the workbook...");

    const answer = await request(
      // The SAME project the plan above was read for. Without it the apply resolved itself
      // all over again, so a plan drawn for one workbook could be carried out against another.
      { action: "apply", direction, ...(project ? { project } : {}), folder, mode: modeSelect.value },
      [...ticked].join("\n"),
    );

    busy = false;

    if (typeof answer.error === "string") {
      say(answer.error, "error");
      apply.disabled = false;
      return;
    }

    const failed = (answer.failed as string[] | undefined) ?? [];
    const summary = String(answer.summary ?? "");

    // Straight back to a fresh plan, which is the honest confirmation: everything that was applied
    // now reads as "same", and anything that did not is still sitting there saying so.
    await refresh();
    say(failed.length > 0 ? `${summary}: ${failed.join("; ")}` : summary, failed.length > 0 ? "error" : "good");
  }

  close.addEventListener("click", () => dismiss());

  // Registered for the dev surface the moment the controls exist, cleared by dismiss. Every
  // probe action lands on a control's own handler, so what a driver does IS what a click does.
  liveDialog = {
    state: () => ({
      direction,
      folder: folderInput.value,
      mode: modeSelect.value,
      busy,
      status: status.textContent ?? "",
      project: projectSelect.value,
      projects: [...projectSelect.options].map((one) => one.value),
      rows: (plan?.items ?? []).map((item) => ({
        file: item.file,
        status: item.status,
        detail: item.detail,
        ticked: ticked.has(item.id),
        actionable: actionable(item),
      })),
    }),
    press: (control) => {
      const button = control === "apply" ? apply
        : control === "close" ? close
        : control === "all" ? selectAll
        : control === "none" ? selectNone
        : control === "export" ? exportButton
        : control === "import" ? importButton
        : null;
      button?.click();
      return button !== null;
    },
    tick: (file, on) => {
      const item = (plan?.items ?? []).find((one) => one.file === file);
      if (!item) {
        return false;
      }

      const box = list.querySelector<HTMLInputElement>(`[data-id="${CSS.escape(item.id)}"] .sync-tick`);
      if (!box) {
        return false;
      }

      if (box.checked !== on) {
        box.click();
      }

      return true;
    },
    chooseProject: (name) => {
      const option = [...projectSelect.options].find((one) => one.value === name);
      if (!option) {
        return false;
      }
      projectSelect.value = name;
      projectSelect.dispatchEvent(new Event("change"));
      return true;
    },
    setFolder: (path) => {
      folderInput.value = path;
      folderInput.dispatchEvent(new Event("change"));
    },
  };

  drawDirection();
  drawList();
  drawDiff();
  void refresh();
  folderInput.focus();
}
