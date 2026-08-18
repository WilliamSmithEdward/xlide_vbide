// WHICH OFFICE APPLICATION THIS ENGINE IS ANSWERING FOR.
//
// The VBE is shared - the add-in loads in Word, PowerPoint, Access and Outlook as readily as in
// Excel - and until 2026-08-18 every one of them was told it was Excel: Word's ThisDocument was
// typed as an Excel.Worksheet and offered a worksheet's members (the owner, having run it there).
//
// PROCESS STATE RATHER THAN A THREADED PARAMETER, and deliberately. The engine is a child of ONE
// Office process and the host cannot change under it, so this is a fact about the run rather than
// about a request. Threading it through assembleContext would put it in three call signatures -
// completion, hover, onType - to say the same unchanging thing at each of them.
//
// It arrives with `project/open`, which is the one message guaranteed to be sent before anything
// is asked, and defaults to Excel so an older shim against a newer engine behaves as it always
// did.

let current = 'excel';

/** Records the host, from `project/open`. An absent or empty value leaves the default alone. */
export function setHostApp(host: string | undefined): void {
    if (typeof host === 'string' && host.length > 0) {
        current = host.toLowerCase();
    }
}

export function hostApp(): string {
    return current;
}

/**
 * Whether the host's own object model is one this engine can speak about.
 *
 * Only Excel today. The analyzer carries a hand-authored Excel object model and none for the
 * other hosts, so in Word the honest answer about `ThisDocument` is NOTHING - a document module
 * that asserts no type loses the bogus worksheet members rather than gaining Word's own, which is
 * strictly better than being confidently wrong and is the state until a Word model exists.
 */
export function hostModelIsKnown(): boolean {
    return current === 'excel';
}
