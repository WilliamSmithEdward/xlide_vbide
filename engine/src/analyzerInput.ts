// One translation from what the host sends to what the analyzer takes.
//
// The host's ModulePayload and the analyzer's module input describe the same module, and every
// index the engine builds needs the conversion between them. It used to be written out by hand at
// each site, and the sites drifted: `implicitMembers` reached the worker seed and the project
// index but not the cross-module fingerprint, so a designer change produced the same fingerprint
// as before it and the memo replayed findings about a control that was gone (the 2026-08-19 ghost
// hunt). The field was then added to the fingerprint by hand, which fixed that field and left the
// next one to repeat the whole story.
//
// So the conversion lives here once. A new host-supplied fact - and there is a queue of them, the
// facts a CodeModule's text cannot carry - arrives at every index by being added in one place,
// rather than at whichever subset of four sites somebody remembered.

import type { VbaProjectModuleInput } from '../../../xlide_vscode/src/vbaProjectAnalysis';
import type { WorkerSeedModule } from '../../../xlide_vscode/src/analysisWorkerProtocol';
import type { EventHandlerDocumentType } from '../../../xlide_vscode/src/analyzer';
import type { ModulePayload } from './protocol';

/**
 * One module, in the shape both the project index and the worker seed accept.
 *
 * The two types are the same fields under two names, so one object satisfies both and the sites
 * that build an index and the site that seeds the worker cannot describe a module differently.
 */
export type AnalyzerModuleInput = VbaProjectModuleInput & WorkerSeedModule;

/**
 * What the analyzer should be told about one module.
 *
 * OMISSION IS A CLAIM HERE, and the opposite of the one a default would make. `implicitMembers`
 * and `predeclaredId` are three-state: a value, or absent meaning nobody read the designer or the
 * attribute header. Spreading the payload's own undefined through is what keeps absent absent -
 * so do not "tidy" these into `?? false` or `?? []`, which is the collapse that has now produced
 * three separate defects upstream (xlide_vscode#24, #47, #48), each one correct code reported as
 * broken in front of the developer who could see it was fine.
 *
 * NULL IS NOT ABSENT ON THE WIRE, which is why it is normalised here - but honestly about what
 * that is worth. JSON can spell an unknown two ways and the analyzer reads only one: a null
 * passes its `!== undefined` test and is stored as a supplied answer, and its analysis cache
 * folds null to the same key FALSE gets. Neither shows through today, measured both ways: a seed
 * carrying null stays silent, and reseeding it to false still reports, because the generation and
 * the cross-module fingerprint invalidate ahead of that cache. So this is a DEFENCE at the wire
 * boundary, not a fix for a defect anybody has seen. It is here because the shim's serialiser
 * happens to omit nulls and a wire protocol should not rest on every host happening to agree.
 */
export function analyzerInputFor(module: ModulePayload): AnalyzerModuleInput {
    return {
        moduleName: module.moduleName,
        source: module.source,
        type: module.type ?? 'standard',
        documentType: module.documentType as EventHandlerDocumentType | undefined,
        implicitMembers: module.implicitMembers ?? undefined,
        predeclaredId: module.predeclaredId ?? undefined,
    };
}
