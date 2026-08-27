/*
 * The analyzer rules dialog: every rule the bundled analyzer ships, with the severity moves it
 * permits, saved FOR THIS MACHINE the moment a tick changes.
 *
 * The list is fetched from the host, which reads it from the engine, which enumerates the
 * analyzer actually running - never a list written down here, so a rule added upstream appears
 * the day the engine is rebuilt. The ticks are guarded by the analyzer's own rules: a warning or
 * information rule can be turned off; a rule whose default is error can at most be reported as a
 * warning, and only where the analyzer marks that safe - most error rules mirror a VBE compile
 * failure and render fixed, with the reason on the row.
 *
 * This is machine policy, persisted in user-space settings.json. One finding at one line is the
 * inline directives' job - those travel with the code, and the problems pane and lightbulb offer
 * them beside the machine-wide switch this dialog owns.
 */

import { openModal } from "./modal.js";
import type { HostAnalysisRule, HostAnalysisRules } from "./bridge.js";

interface RulesAccess {
  fetch(): Promise<HostAnalysisRules | null>;
  /** off | warning | error | information | default (clears the override). */
  set(code: string, severity: string): void;
}

/**
 * Opens the dialog. One at a time: opening while open focuses the one that exists. `focusCode`
 * scrolls to and highlights one rule, for the problems pane's "Analyzer Rules..." on a finding.
 */
export function openAnalysisRulesDialog(
  access: RulesAccess,
  closed: () => void,
  focusCode?: string,
): void {
  if (document.getElementById("analysis-rules-backdrop")) {
    (document.querySelector("#analysis-rules-card input") as HTMLElement | null)?.focus();
    return;
  }

  const { card, dismiss } = openModal({
    backdropId: "analysis-rules-backdrop",
    cardId: "analysis-rules-card",
    label: "Analyzer rules",
    closed,
  });

  const head = document.createElement("div");
  head.id = "settings-head";

  const title = document.createElement("span");
  title.textContent = "Analyzer rules";

  const close = document.createElement("button");
  close.type = "button";
  close.id = "settings-close";
  close.title = "Close (Esc)";
  close.setAttribute("aria-label", "Close analyzer rules");
  close.innerHTML = '<span class="codicon codicon-close" aria-hidden="true"></span>';

  head.append(title, close);
  card.appendChild(head);

  // Where these choices live, said plainly: the one fact that separates this dialog from the
  // inline directives sitting two menu items away.
  const scope = document.createElement("div");
  scope.className = "analysis-rules-scope";
  scope.textContent = "For this machine: saved in your user settings and applied to every "
    + "workbook you open. To silence one finding at one line instead, use the inline "
    + "suppression comment - it travels with the code.";
  card.appendChild(scope);

  const search = document.createElement("input");
  search.type = "search";
  search.id = "analysis-rules-search";
  search.placeholder = "Filter rules by code, title, or category";
  search.setAttribute("aria-label", "Filter rules");
  card.appendChild(search);

  const body = document.createElement("div");
  body.id = "analysis-rules-body";
  body.textContent = "Reading the rule catalog…";
  card.appendChild(body);

  close.addEventListener("click", dismiss);

  let catalog: HostAnalysisRule[] = [];
  // The standing overrides, kept current OPTIMISTICALLY as ticks change: the host validates and
  // persists, and its notice reports the outcome; the next open re-reads the truth.
  const overrides = new Map<string, string>();

  const rowFor = (rule: HostAnalysisRule): HTMLElement => {
    const row = document.createElement("div");
    row.className = "settings-row analysis-rule-row";
    row.dataset.code = rule.code;

    const text = document.createElement("div");
    text.className = "settings-text";

    const label = document.createElement("label");
    label.className = "settings-label";
    label.textContent = rule.title;
    label.htmlFor = `analysis-rule-${rule.code}`;

    const description = document.createElement("div");
    description.className = "settings-description";
    description.textContent = `${rule.code} · ${rule.category} · default ${rule.defaultSeverity}`;

    text.append(label, description);
    row.appendChild(text);

    const canTurnOff = rule.allowed.includes("off");
    const canDowngrade = rule.defaultSeverity === "error" && rule.allowed.includes("warning");

    /*
     * ONE right-hand container per row, whatever it holds. The row is flex space-between, so
     * three loose children - text, a label, a checkbox - spread into three ragged columns, which
     * is exactly what the first build looked like. Two children pin the text left and the
     * control right, and the checkboxes land on one shared edge whether or not a label sits
     * beside them.
     */
    const control = document.createElement("div");
    control.className = "analysis-rule-control";
    row.appendChild(control);

    if (canTurnOff) {
      const tick = document.createElement("input");
      tick.type = "checkbox";
      tick.id = `analysis-rule-${rule.code}`;
      tick.checked = overrides.get(rule.code) !== "off";
      tick.title = "Unticked: this rule reports nothing on this machine";
      tick.addEventListener("change", () => {
        if (tick.checked) { overrides.delete(rule.code); } else { overrides.set(rule.code, "off"); }
        access.set(rule.code, tick.checked ? "default" : "off");
      });
      control.appendChild(tick);
    } else if (canDowngrade) {
      const tick = document.createElement("input");
      tick.type = "checkbox";
      tick.id = `analysis-rule-${rule.code}`;
      tick.checked = overrides.get(rule.code) === "warning";
      tick.title = "Ticked: reported as a warning instead of an error, on this machine";
      tick.addEventListener("change", () => {
        if (tick.checked) { overrides.set(rule.code, "warning"); } else { overrides.delete(rule.code); }
        access.set(rule.code, tick.checked ? "warning" : "default");
      });
      const asWarning = document.createElement("label");
      asWarning.className = "analysis-rule-hint";
      asWarning.textContent = "as warning";
      asWarning.htmlFor = tick.id;
      control.append(asWarning, tick);
    }
    // A rule that permits nothing gets NO control at all: those live together in the
    // "Always on" section at the bottom, whose heading says why once instead of 99 times.

    return row;
  };

  const draw = (): void => {
    const wanted = search.value.trim().toLowerCase();
    body.replaceChildren();

    const shown = catalog.filter((rule) => wanted.length === 0
      || rule.code.includes(wanted)
      || rule.title.toLowerCase().includes(wanted)
      || rule.category.toLowerCase().includes(wanted));

    if (shown.length === 0) {
      const empty = document.createElement("div");
      empty.className = "settings-description";
      empty.textContent = catalog.length === 0
        ? "The rule catalog could not be read - the analysis engine is not up."
        : `Nothing matches "${search.value.trim()}".`;
      body.appendChild(empty);
      return;
    }

    /*
     * THE ADJUSTABLE RULES FIRST, grouped by category; everything the analyzer permits no
     * override on gathers into one "Always on" section at the bottom (the owner, 2026-08-27).
     * Ninety-nine of the 122 rules are fixed - interleaved, they buried the twenty-three a
     * person can actually change, and every one of them repeated the same badge.
     */
    const adjustable = shown.filter((rule) => rule.allowed.length > 0);
    const fixed = shown.filter((rule) => rule.allowed.length === 0);

    let category = "";
    for (const rule of adjustable) {
      if (rule.category !== category) {
        category = rule.category;
        const heading = document.createElement("div");
        heading.className = "analysis-rules-category";
        heading.textContent = category;
        body.appendChild(heading);
      }

      body.appendChild(rowFor(rule));
    }

    if (fixed.length > 0) {
      const heading = document.createElement("div");
      heading.className = "analysis-rules-category analysis-rules-always-on";
      heading.textContent = "Always on";
      body.appendChild(heading);

      const why = document.createElement("div");
      why.className = "settings-description";
      why.textContent = "These mirror VBE compile failures or deterministic runtime errors, so "
        + "the analyzer allows no override. An inline suppression comment still works on a "
        + "single finding.";
      body.appendChild(why);

      for (const rule of fixed) {
        body.appendChild(rowFor(rule));
      }
    }

    if (focusCode !== undefined && wanted.length === 0) {
      const target = body.querySelector<HTMLElement>(`[data-code="${CSS.escape(focusCode)}"]`);
      if (target) {
        target.classList.add("analysis-rule-focused");
        target.scrollIntoView({ block: "center" });
      }
    }
  };

  search.addEventListener("input", draw);

  void access.fetch().then((answer) => {
    if (!document.getElementById("analysis-rules-body")) {
      return;
    }

    catalog = answer?.rules ?? [];
    overrides.clear();
    for (const [code, severity] of Object.entries(answer?.overrides ?? {})) {
      overrides.set(code, severity);
    }

    draw();
    search.focus();
  });
}
