// The quiet a live splitter drag keeps, shared by every splitter on the page.
//
// Each splitter used to call layoutChanged - a forced synchronous Monaco layout of every open
// editor - once per pointermove, on top of automaticLayout's ResizeObserver already tracking
// the same resize. That is the exact doubling the window-resize path diagnosed and fixed on
// 2026-08-05 ("running it per event doubled every layout of a drag, which read as latency and
// churn"), and the fix was applied only there: no splitter drag got it, so the most common
// layout gesture on the page still paid N editors x several pointermoves per frame (the
// audit's C10, 2026-08-12).
//
// So a drag does what the window resize does: the observer tracks the live geometry, the
// live-resize class keeps the minimap from flickering a frame behind, and ONE settling layout
// runs at the end for the final frame the observer may have missed. Keyboard steps are
// discrete and keep their immediate layout - one step, one layout is right there.

/** A live drag began: fade the minimap's churn out, exactly as the window resize does. */
export function beginLiveDrag(): void {
    document.body.classList.add("live-resize");
}

/** The drag ended, however it ended: quiet off, and the one settling layout. */
export function endLiveDrag(settle: () => void): void {
    document.body.classList.remove("live-resize");
    settle();
}

/**
 * A splitter's whole pointer lifecycle: capture on the press, apply the delta on each move,
 * and let go on release OR cancel - releasing the capture and taking all three listeners off
 * again. The bookkeeping is the point. Three splitters spelled it out separately, and every
 * one of them had to remember the same three removeEventListener calls; a drag that forgets
 * one keeps handling moves after the pointer is gone.
 *
 * What actually differs between splitters is which axis they read and whether they persist,
 * so those are the arguments: `positionOf` answers the coordinate that matters for this
 * splitter, and `persist` is omitted by the one whose size is not remembered.
 */
export function installSplitterDrag(splitter: HTMLElement, drag: {
    positionOf: (event: PointerEvent) => number;
    apply: (delta: number) => void;
    settle: () => void;
    persist?: () => void;
}): void {
    splitter.addEventListener("pointerdown", (event) => {
        if (event.button !== 0) {
            return;
        }

        event.preventDefault();
        splitter.setPointerCapture(event.pointerId);
        beginLiveDrag();
        let last = drag.positionOf(event);

        const move = (moved: PointerEvent): void => {
            const position = drag.positionOf(moved);
            drag.apply(position - last);
            last = position;
        };

        const end = (ended: PointerEvent): void => {
            splitter.releasePointerCapture(ended.pointerId);
            splitter.removeEventListener("pointermove", move);
            splitter.removeEventListener("pointerup", end);
            splitter.removeEventListener("pointercancel", end);
            endLiveDrag(drag.settle);
            drag.persist?.();
        };

        splitter.addEventListener("pointermove", move);
        splitter.addEventListener("pointerup", end);
        splitter.addEventListener("pointercancel", end);
    });
}
