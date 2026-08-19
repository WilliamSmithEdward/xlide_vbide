// Serves hovers by reusing the editor extension's hover resolver.
//
// Same shape as completion: the resolver is pure, the project facts come from the shared
// assembly, and the answer describes the identifier under the cursor - its declaration line, its
// origin, and its documentation when any is known.

import { resolveHover } from '../../../xlide_vscode/src/analyzer';
import { assembleContext } from './moduleContext';
import type { HoverParams, HoverPayload, ModulePayload } from './protocol';

export function hoverFor(
    seeded: readonly ModulePayload[],
    params: HoverParams & { source: string },
): HoverPayload | null {
    const ctx = assembleContext(seeded, params);

    const info = resolveHover(params.source, params.offset, {
        codeNames: ctx.codeNames,
        meType: ctx.meType,
        meProjectType: ctx.meProjectType,
        moduleName: ctx.current.name,
        moduleKind: ctx.moduleKind,
        // A form's controls, so hovering RegionPick answers its control-hood and type, and
        // hovering a member answers the MSForms signature (xlide_vscode#19).
        implicitMembers: ctx.implicitMembers,
        projectClassMembers: ctx.projectClassMembers,
        projectTypes: ctx.projectTypes,
        projectProcedures: ctx.projectProcedures,
        // The host this engine runs in, when it is not Excel (the resolvers' default).
        model: ctx.hostModel,
    });

    if (!info) {
        return null;
    }

    return {
        signature: info.signature,
        details: [...info.details],
        documentation: info.documentation,
        span: { start: info.span.start, end: info.span.end },
    };
}
