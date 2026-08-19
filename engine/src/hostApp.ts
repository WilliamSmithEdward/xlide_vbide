// WHICH OFFICE APPLICATION THIS ENGINE IS ANSWERING FOR, and the object model that answers
// with it.
//
// The VBE is shared - the add-in loads in Word, PowerPoint, Access and Outlook as readily as in
// Excel - and until 2026-08-18 every one of them was told it was Excel: Word's ThisDocument was
// typed as an Excel.Worksheet and offered a worksheet's members (the owner, having run it
// there). The host token arrived first; the models arrived with the analyzer's host registry
// (xlide_vscode 4.0.0, issues #24/#25), and this file is where the two meet.
//
// PROCESS STATE RATHER THAN A THREADED PARAMETER, and deliberately. The engine is a child of ONE
// Office process and the host cannot change under it, so this is a fact about the run rather than
// about a request. Threading it through assembleContext would put it in three call signatures -
// completion, hover, onType - to say the same unchanging thing at each of them.
//
// It arrives with `project/open`, which is the one message guaranteed to be sent before anything
// is asked, and defaults to Excel so an older shim against a newer engine behaves as it always
// did.

import {
    EMPTY_HOST_MODEL,
    hostObjectModelForToken,
} from '../../../xlide_vscode/src/analyzer/host/hostRegistry';
import { getExcelObjectModel, type HostObjectModel } from '../../../xlide_vscode/src/analyzer';

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
 * The registry's model for the current host, or undefined for Excel - the registry answers
 * undefined there ON PURPOSE, so every downstream `?? getExcelObjectModel()` default keeps the
 * behavior Excel always had. Callers that thread a model into the analyzer's resolvers want
 * exactly this shape: nothing to thread in Excel, the host's own model elsewhere, and the
 * empty model (every lookup misses, nothing asserted) in a host the registry does not know.
 */
export function currentHostModelOverride(): HostObjectModel | undefined {
    return hostObjectModelForToken(current);
}

/** The current host's model, Excel's included: for callers that read it rather than thread it. */
export function currentHostModel(): HostObjectModel {
    return currentHostModelOverride() ?? getExcelObjectModel();
}

/**
 * Whether the host's own object model is one this engine can speak about. Excel, Word,
 * PowerPoint and Access today - whatever the registry holds, rather than a list kept here.
 * A host it answers with the EMPTY model for (Outlook, Visio, unknowns) is honestly unknown:
 * document modules assert no type and the knowledge routes say so, which beats being
 * confidently wrong in someone else's application.
 */
export function hostModelIsKnown(): boolean {
    return currentHostModel() !== EMPTY_HOST_MODEL;
}

/**
 * Whether the current host's VBE carries the MSForms designer at all. Access's does not -
 * Access VBA has its own Forms and no UserForms, so its VBE offers no Insert > UserForm and
 * `VBComponents.Import` of a .frm can only fail there (the owner, 2026-08-19). Everything that
 * would CREATE a userform consults this; reading one is not gated, because a form component
 * can never exist in such a host to begin with.
 */
export function hostCarriesMsForms(): boolean {
    return current !== 'access';
}
