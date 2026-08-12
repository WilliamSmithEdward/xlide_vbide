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
