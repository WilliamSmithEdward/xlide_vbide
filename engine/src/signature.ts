// Serves call tips by reusing the editor extension's signature resolver.
//
// Same shape as hover: the resolver is pure, the project facts come from the shared assembly,
// and the answer is the callee's signature line with the active parameter marked. Signatures are
// never invented — a callee with no known signature yields no tip.

import { resolveSignatureHelp } from '../../../xlide_vscode/src/analyzer';
import { assembleContext } from './moduleContext';
import type { ModulePayload, SignatureHelpParams, SignatureInfoPayload } from './protocol';

export function signatureHelpFor(
    seeded: readonly ModulePayload[],
    params: SignatureHelpParams & { source: string },
): SignatureInfoPayload | null {
    const ctx = assembleContext(seeded, params);

    const info = resolveSignatureHelp(params.source, params.offset, {
        codeNames: ctx.codeNames,
        meType: ctx.meType,
        meProjectType: ctx.meProjectType,
        projectClassMembers: ctx.projectClassMembers,
        moduleName: ctx.current.name,
        projectProcedures: ctx.projectProcedures,
    });

    if (!info) {
        return null;
    }

    return {
        label: info.label,
        parameters: info.parameters.map((parameter) => ({
            label: parameter.label,
            documentation: parameter.documentation,
        })),
        activeParameter: info.activeParameter,
        documentation: info.documentation,
        details: info.details ? [...info.details] : undefined,
    };
}
