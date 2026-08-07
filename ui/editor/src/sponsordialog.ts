/*
 * Where to support the work, if it has been useful.
 *
 * Three addresses, opened by the HOST rather than followed here: the surface runs under a policy
 * that forbids navigating anywhere, so an anchor on this page would look like a link and do
 * nothing when pressed (the same reason the About dialog prints its repository as text). The host
 * opens them, and it will only open these three.
 *
 * Each row also copies, because a browser that is slow to come up, or a machine where the default
 * is something unexpected, should not leave anyone with no way to reach the page.
 */

interface SponsorLink {
  label: string;
  detail: string;
  url: string;
  /** A codicon name, or an emoji when the icon set has nothing for it. */
  icon: string;
}

const LINKS: SponsorLink[] = [
  {
    label: "GitHub Sponsors",
    detail: "Recurring or one-off, through GitHub",
    icon: "github",
    url: "https://github.com/sponsors/WilliamSmithEdward",
  },
  {
    label: "PayPal",
    detail: "One-off, no account needed",
    icon: "credit-card",
    url: "https://www.paypal.com/donate/?business=ML855BRLNR838&no_recurring=0&item_name=VBA+has+always+treated+me+well.+It+was+how+I+first+grew+professional+as+a+programmer%2C+I%27m+happy+to+show+it+some+love+%E2%9D%A4%EF%B8%8F&currency_code=USD",
  },
  {
    // An emoji, because the icon set has no note and no dollar: this row was drawing a codicon
    // that does not exist, which is not a broken glyph but nothing at all — the label simply sat
    // further left than the two above it (the developer, 2026-08-07).
    label: "Cash App",
    detail: "$williamesmithjcil",
    icon: "💵",
    url: "https://cash.app/$williamesmithjcil",
  },
];

/** True for an icon that is a character to print rather than a name to look up. */
function isEmoji(icon: string): boolean {
  return !/^[a-z0-9-]+$/.test(icon);
}

/** What the dialog needs of the world: somewhere outside the page to open an address. */
export interface SponsorHandlers {
  openExternal(url: string): void;
}

/**
 * Opens the dialog. One at a time, like the others: opening while open focuses what is there
 * rather than stacking a second card nobody asked for.
 */
export function openSponsorDialog(handlers: SponsorHandlers, closed?: () => void): void {
  const existing = document.getElementById("sponsor-card");
  if (existing) {
    existing.querySelector<HTMLElement>("#sponsor-close")?.focus();
    return;
  }

  const backdrop = document.createElement("div");
  backdrop.id = "sponsor-backdrop";

  const card = document.createElement("div");
  card.id = "sponsor-card";
  card.setAttribute("role", "dialog");
  card.setAttribute("aria-modal", "true");
  card.setAttribute("aria-label", "Support xlide");

  const head = document.createElement("div");
  head.id = "sponsor-head";

  const title = document.createElement("div");
  title.id = "sponsor-title";
  title.textContent = "Support xlide";

  const close = document.createElement("button");
  close.type = "button";
  close.id = "sponsor-close";
  close.title = "Close (Esc)";
  close.setAttribute("aria-label", "Close");
  close.innerHTML = '<span class="codicon codicon-close" aria-hidden="true"></span>';

  head.append(title, close);

  const blurb = document.createElement("p");
  blurb.className = "sponsor-blurb";
  blurb.textContent =
    "VBA has always treated me well. It is how I first grew professional as a programmer, and "
    + "xlide is what I wish it had come with. If it has been useful, here is where to say so.";

  const list = document.createElement("div");
  list.id = "sponsor-list";

  for (const link of LINKS) {
    const row = document.createElement("div");
    row.className = "sponsor-row";

    const open = document.createElement("button");
    open.type = "button";
    open.className = "sponsor-open";
    open.title = link.url;

    // Same box either way, so the three labels start on one line whichever kind of mark is above
    // them. An emoji is set in the box rather than beside it for exactly that reason.
    const icon = document.createElement("span");
    icon.setAttribute("aria-hidden", "true");
    if (isEmoji(link.icon)) {
      icon.className = "sponsor-emoji";
      icon.textContent = link.icon;
    } else {
      icon.className = `codicon codicon-${link.icon}`;
    }

    const words = document.createElement("span");
    words.className = "sponsor-words";

    const label = document.createElement("span");
    label.className = "sponsor-label";
    label.textContent = link.label;

    const detail = document.createElement("span");
    detail.className = "sponsor-detail";
    detail.textContent = link.detail;

    words.append(label, detail);

    const away = document.createElement("span");
    away.className = "codicon codicon-link-external sponsor-away";
    away.setAttribute("aria-hidden", "true");

    open.append(icon, words, away);
    open.addEventListener("click", () => handlers.openExternal(link.url));

    const copy = document.createElement("button");
    copy.type = "button";
    copy.className = "sponsor-copy";
    copy.textContent = "Copy";
    copy.title = "Copy the address";
    copy.addEventListener("click", () => {
      void navigator.clipboard?.writeText(link.url).then(
        () => {
          copy.textContent = "Copied";
          window.setTimeout(() => (copy.textContent = "Copy"), 1200);
        },
        () => {
          copy.textContent = "Press Ctrl+C";
        },
      );
    });

    row.append(open, copy);
    list.appendChild(row);
  }

  const thanks = document.createElement("p");
  thanks.className = "sponsor-thanks";
  thanks.textContent = "Nothing here is ever required. Thank you for using it either way.";

  card.append(head, blurb, list, thanks);
  backdrop.appendChild(card);
  document.body.appendChild(backdrop);

  const dismiss = (): void => {
    document.removeEventListener("keydown", onKey, true);
    backdrop.remove();
    closed?.();
  };

  function onKey(event: KeyboardEvent): void {
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      dismiss();
    }
  }

  close.addEventListener("click", () => dismiss());
  backdrop.addEventListener("mousedown", (event) => {
    if (event.target === backdrop) {
      dismiss();
    }
  });

  document.addEventListener("keydown", onKey, true);
  card.querySelector<HTMLElement>(".sponsor-open")?.focus();
}
