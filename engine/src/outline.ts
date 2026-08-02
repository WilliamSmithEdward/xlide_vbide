// Serves a module's outline: its procedures, in declaration order, the way the extension's
// tree lists them under a module node. The parser the analyzer already runs is the source of
// truth, so the tree and the diagnostics describe the same procedures.

import { parseModule } from '../../../xlide_vscode/src/analyzer';
import type { ModulePayload, OutlineParams, OutlineProcedure } from './protocol';

// The extension's tree spells the kind before the name: "Sub Main", "Property Get Name".
const PROC_KIND_LABELS: Record<string, string> = {
    Sub: 'Sub',
    Function: 'Function',
    PropertyGet: 'Property Get',
    PropertyLet: 'Property Let',
    PropertySet: 'Property Set',
};

export function outlineFor(
    seeded: readonly ModulePayload[],
    params: OutlineParams,
): OutlineProcedure[] {
    // The live source when the request is about the module being edited; the seeded copy —
    // current as of the last write-back — for every other module.
    const source = params.source
        ?? seeded.find((module) => module.moduleName.toLowerCase() === params.moduleName.toLowerCase())?.source;
    if (source === undefined) {
        return [];
    }

    const procedures: OutlineProcedure[] = [];
    for (const member of parseModule(source).members) {
        if (member.kind !== 'Procedure') {
            continue;
        }

        procedures.push({
            name: member.name,
            kind: PROC_KIND_LABELS[member.procKind] ?? member.procKind,
            line: lineAt(source, member.span.start),
        });
    }

    return procedures;
}

/** 1-based physical line containing an offset. */
function lineAt(source: string, offset: number): number {
    let line = 1;
    const end = Math.min(Math.max(offset, 0), source.length);
    for (let i = 0; i < end; i++) {
        if (source.charCodeAt(i) === 10) {
            line++;
        }
    }
    return line;
}

/**
 * The project's own words: every name that denotes a type (class and document modules, Type and
 * Enum declarations) and every procedure name, across the whole project. The surface's tokenizer
 * takes these as word lists, which is how a name reads as what it is wherever it appears.
 */
export function projectWordsFor(
    modules: readonly ModulePayload[],
): { types: string[]; procedures: string[] } {
    const types = new Set<string>();
    const procedures = new Set<string>();

    for (const module of modules) {
        if (module.type === 'class' || module.type === 'document') {
            types.add(module.moduleName);
        }

        try {
            for (const member of parseModule(module.source).members) {
                if (member.kind === 'Procedure') {
                    procedures.add(member.name);
                } else if (member.kind === 'Type' || member.kind === 'Enum') {
                    const name = 'name' in member && typeof member.name === 'string' ? member.name : undefined;
                    if (name) {
                        types.add(name);
                    }
                }
            }
        } catch {
            // A module that will not parse lends no words; the rest still do.
        }
    }

    return { types: [...types].sort(), procedures: [...procedures].sort() };
}
