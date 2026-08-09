/*
 * Import and export, in one window.
 *
 * The companion editor spells this as three commands and a preview panel. Here it is one surface
 * with a direction on it, because the two directions ask the same questions — which folder, which
 * modules, and what happens to the ones only one side has — and answering them in two places is
 * how the two ends of a round trip drift apart.
 *
 * NOTHING IS DECIDED HERE. The host works out the plan and the host carries it out; this draws the
 * plan and sends back the rows that are ticked. The debug api's `sync` route reaches the very same
 * call in the host, so a plan read by a harness is the plan drawn here, and Apply leaves the
 * project exactly where the api would have left it. That is the whole reason the dialog is thin.
 */

interface SyncDiffLine {
  leftNumber: number | null;
  rightNumber: number | null;
  left: string;
  right: string;
  kind: "equal" | "changed" | "added" | "removed";
}

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
export function openSyncDialog(request: SyncRequest, closed: () => void): void {
  if (document.getElementById("sync-backdrop")) {
    document.querySelector<HTMLElement>("#sync-folder")?.focus();
    return;
  }

  let direction: "export" | "import" = "export";
  let folder = "";
  let plan: SyncPlan | null = null;
  let selectedId: string | null = null;
  let showHeaders = false;
  let busy = false;
  const ticked = new Set<string>();

  const backdrop = document.createElement("div");
  backdrop.id = "sync-backdrop";

  const card = document.createElement("div");
  card.id = "sync-card";
  card.setAttribute("role", "dialog");
  card.setAttribute("aria-modal", "true");
  card.setAttribute("aria-label", "Import and export modules");

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
  modeSelect.addEventListener("change", () => void refresh());

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

  card.append(head, directions, folderRow, modeRow, body, foot);
  backdrop.appendChild(card);

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
        { value: "exportAll", label: "Leave the file alone" },
        { value: "trueUp", label: "Delete the file" },
      ]
      : [
        { value: "updateOnly", label: "Leave the module alone" },
        { value: "trueUpStandardClass", label: "Delete the module" },
      ];

    for (const choice of choices) {
      const option = document.createElement("option");
      option.value = choice.value;
      option.textContent = choice.label;
      modeSelect.appendChild(option);
    }

    modeLabel.textContent = direction === "export"
      ? "A file with no module"
      : "A module with no file";
    modeHint.textContent = direction === "export"
      ? "Deleting makes the folder match the project exactly."
      : "Only standard and class modules are ever deleted.";
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

    for (const item of items) {
      const row = document.createElement("div");
      row.className = "sync-item";
      row.dataset.id = item.id;
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

    for (const line of lines) {
      const row = document.createElement("div");
      row.className = "sync-diff-row";
      row.dataset.kind = line.kind;

      // A line only one side has carries no number for the other, and the host spells that
      // absence as null rather than as a missing field. Tested for loosely on purpose: `=== undefined`
      // let a literal "null" through into the gutter.
      const leftNumber = document.createElement("span");
      leftNumber.className = "sync-gutter";
      leftNumber.textContent = line.leftNumber == null ? "" : String(line.leftNumber);

      const leftText = document.createElement("span");
      leftText.className = "sync-code";
      leftText.textContent = line.left;

      const rightNumber = document.createElement("span");
      rightNumber.className = "sync-gutter";
      rightNumber.textContent = line.rightNumber == null ? "" : String(line.rightNumber);

      const rightText = document.createElement("span");
      rightText.className = "sync-code";
      rightText.textContent = line.right;

      row.append(leftNumber, leftText, rightNumber, rightText);
      diff.appendChild(row);
    }
  }

  function drawWarnings(): void {
    const warnings = plan?.warnings ?? [];
    if (warnings.length > 0) {
      say(warnings.join("  "), "error");
    }
  }

  async function refresh(): Promise<void> {
    drawDirection();
    if (busy) {
      return;
    }

    busy = true;
    apply.disabled = true;
    say("Working out what would change...");

    const answer = await request({
      direction,
      ...(folder ? { folder } : {}),
      mode: modeSelect.value,
    });

    busy = false;

    if (typeof answer.error === "string") {
      plan = null;
      drawList();
      drawDiff();
      say(answer.error, "error");
      return;
    }

    plan = answer as unknown as SyncPlan;
    folder = plan.folder;
    folderInput.value = plan.folder;
    title.textContent = `Import and export — ${plan.project}`;

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
      { action: "apply", direction, folder, mode: modeSelect.value },
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
    say(failed.length > 0 ? `${summary} — ${failed.join("; ")}` : summary, failed.length > 0 ? "error" : "good");
  }

  const onKey = (event: KeyboardEvent): void => {
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      dismiss();
    }
  };

  function dismiss(): void {
    backdrop.remove();
    document.removeEventListener("keydown", onKey, true);
    closed();
  }

  close.addEventListener("click", () => dismiss());
  backdrop.addEventListener("pointerdown", (event) => {
    if (event.target === backdrop) {
      dismiss();
    }
  });

  document.addEventListener("keydown", onKey, true);
  document.body.appendChild(backdrop);

  drawDirection();
  drawList();
  drawDiff();
  void refresh();
  folderInput.focus();
}
