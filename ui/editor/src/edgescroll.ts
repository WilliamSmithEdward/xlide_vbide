/*
 * A strip that scrolls sideways, with an edge at each end that has more past it.
 *
 * Built for the command strip and then wanted by the tab strip, which is the moment to share it
 * rather than write it twice: two strips that scroll almost the same way is how one of them ends
 * up with the wheel behaviour and the other without.
 *
 * The edges are deliberately not buttons in appearance. They run nothing - they reveal what is
 * already there - and something that looks like the commands beside it is something a person
 * presses expecting a command.
 */

/** How far one press travels: enough to feel like progress, short enough that nothing sails past. */
const STEP = 4 * 28;

/** While held, it keeps going. Crossing a strip a press at a time is a tax on using its far end. */
const REPEAT_MILLISECONDS = 140;

export interface EdgeScroll {
  /** Re-reads the strip's width and shows or hides each edge. Safe to call at any time. */
  update(): void;
  dispose(): void;
}

/**
 * Gives a scrolling strip its two edges. The buttons are inserted before and after the strip in
 * its own parent, so the parent decides where they sit; everything else is wired here.
 */
export function installEdgeScroll(strip: HTMLElement, className = "toolbar-edge"): EdgeScroll {
  const parent = strip.parentElement;
  if (!parent) {
    throw new Error("a strip must be in the document before its edges can be placed");
  }

  const back = edgeButton("chevron-left", "Scroll left", "start", className);
  const forward = edgeButton("chevron-right", "Scroll right", "end", className);
  parent.insertBefore(back, strip);
  parent.insertBefore(forward, strip.nextSibling);

  const update = (): void => {
    const furthest = strip.scrollWidth - strip.clientWidth;
    // A pixel of slack: fractional widths mean scrollLeft rarely lands exactly on the end.
    back.hidden = strip.scrollLeft <= 1;
    forward.hidden = strip.scrollLeft >= furthest - 1;
  };

  const slide = (direction: -1 | 1): void => {
    strip.scrollBy({ left: direction * STEP, behavior: "smooth" });
  };

  const held: Array<() => void> = [];

  for (const [button, direction] of [[back, -1], [forward, 1]] as const) {
    button.addEventListener("click", () => slide(direction));
    button.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) return;
      const timer = window.setInterval(() => slide(direction), REPEAT_MILLISECONDS);
      const stop = (): void => {
        window.clearInterval(timer);
        window.removeEventListener("pointerup", stop);
        window.removeEventListener("pointercancel", stop);
      };
      window.addEventListener("pointerup", stop);
      window.addEventListener("pointercancel", stop);
      held.push(stop);
    });
  }

  // A vertical wheel over a horizontal strip should move it: that is the wheel most mice have,
  // and the strip is the only thing under the pointer.
  const onWheel = (event: WheelEvent): void => {
    if (event.deltaX !== 0 || event.deltaY === 0) return;
    if (strip.scrollWidth <= strip.clientWidth) return;
    event.preventDefault();
    strip.scrollLeft += event.deltaY;
  };

  strip.addEventListener("wheel", onWheel, { passive: false });
  strip.addEventListener("scroll", update, { passive: true });

  // A pane is resized by dragging a splitter, not only by resizing the window, so the element
  // itself is what has to be watched.
  const observer = new ResizeObserver(() => update());
  observer.observe(strip);

  update();

  return {
    update,
    dispose(): void {
      observer.disconnect();
      strip.removeEventListener("wheel", onWheel);
      strip.removeEventListener("scroll", update);
      for (const stop of held) stop();
      back.remove();
      forward.remove();
    },
  };
}

function edgeButton(
  icon: string,
  label: string,
  edge: "start" | "end",
  className: string,
): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `${className} ${className}-${edge}`;
  button.title = label;
  button.setAttribute("aria-label", label);
  // Nothing here is reachable only this way: every tab and command is also on a menu or a key,
  // and a keyboard user tabs through the strip rather than paging it.
  button.tabIndex = -1;
  button.hidden = true;
  button.innerHTML = `<span class="codicon codicon-${icon}" aria-hidden="true"></span>`;
  return button;
}
