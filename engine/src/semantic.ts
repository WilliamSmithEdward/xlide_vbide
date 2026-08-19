// Serves semantic tokens by reusing the editor extension's type-reference collectors.
//
// The surface's own tokenizer already colours the project's words: `project/open` hands it every
// name that denotes a type and every name that denotes a procedure, and its grammar paints them.
// That gets `ROneCOne.Create(...)` reading as a type and a call. What a grammar cannot do is tell
// a class from an enum from a user-defined type, or tell a host global from a local that shadows
// its name - those need the analysis, and this is where it comes from.
//
// Three collectors, the same set the extension uses: type references resolved against the project's
// own types, host globals that no declaration in the module has shadowed, and resolved method calls
// on a form's designer-declared controls (xlide_vscode#20) - `RegionPick.AddItem` paints the way
// `Len` does, while a property read or an unresolved member stays untouched.

import {
    collectHostGlobalTokens,
    collectImplicitMemberMethodTokens,
    resolveTypeSemanticTokens,
} from '../../../xlide_vscode/src/analyzer';
import { assembleContext } from './moduleContext';
import type { ModulePayload, SemanticTokenPayload, SemanticTokensParams } from './protocol';

export function semanticTokensFor(
    seeded: readonly ModulePayload[],
    params: SemanticTokensParams & { source: string },
): SemanticTokenPayload[] {
    const ctx = assembleContext(seeded, params);

    const tokens = [
        ...resolveTypeSemanticTokens(params.source, { projectTypes: ctx.projectTypes, model: ctx.hostModel }),
        // The host's own globals when the host is not Excel: in Word, ActiveDocument colours
        // and ActiveSheet does not.
        ...collectHostGlobalTokens(params.source, ctx.hostModel),
        // The collector guards meType itself - only an MSForms.* type engages it - so a
        // worksheet's Excel.Worksheet passes through as an early return, not a special case here.
        ...collectImplicitMemberMethodTokens(params.source, {
            implicitMembers: ctx.implicitMembers,
            meType: ctx.meType,
        }),
    ];

    // Sorted by position, because the surface encodes them as deltas from the token before and
    // an out-of-order pair would place every token after it wrongly. The two collectors each
    // scan the module independently, so their outputs interleave rather than concatenate.
    return tokens
        .map((token) => ({
            start: token.span.start,
            end: token.span.end,
            type: token.tokenType,
            ...(token.modifiers?.length ? { modifiers: [...token.modifiers] } : {}),
        }))
        .sort((left, right) => left.start - right.start || left.end - right.end);
}
