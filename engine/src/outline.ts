// Serves a module's outline: its procedures, in declaration order, the way the extension's
// tree lists them under a module node. The parser the analyzer already runs is the source of
// truth, so the tree and the diagnostics describe the same procedures.

import { parseModule } from '../../../xlide_vscode/src/analyzer';
import { lineStarts, toLineColumn } from './navigation';
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

    // ONE table for the module, then a binary search per procedure.
    //
    // This counted newlines from the start of the file for EVERY procedure, which is a scan whose
    // length is that procedure's offset. Summed over a module it is quadratic in the source, and
    // the module where it hurts is exactly the module a tree is worth having in: the 11,000-line
    // fixture has some 1,600 procedures with the last near offset 256,000, so the tree cost about
    // 200 million character reads and 238ms, per outline, and the surface asks for one after
    // every push it notices (2026-08-08).
    //
    // The diagnostics path had already learned this and says so in its own comment. This is the
    // same table and the same search.
    const starts = lineStarts(source);

    const procedures: OutlineProcedure[] = [];
    for (const member of parseModule(source).members) {
        if (member.kind !== 'Procedure') {
            continue;
        }

        procedures.push({
            name: member.name,
            kind: PROC_KIND_LABELS[member.procKind] ?? member.procKind,
            line: toLineColumn(starts, member.span.start).line,
        });
    }

    return procedures;
}

/**
 * The project's own words: every name that denotes a type (class and document modules, Type and
 * Enum declarations) and every procedure name, across the whole project. The surface's tokenizer
 * takes these as word lists, which is how a name reads as what it is wherever it appears.
 */
export function projectWordsFor(
    projectId: string,
    modules: readonly ModulePayload[],
): { types: string[]; procedures: string[] } {
    const types = new Set<string>();
    const procedures = new Set<string>();

    // This project's entries, rebuilt: a module that is gone leaves with them, and the map holds
    // one entry per open module rather than one per version of one.
    const held = wordsByModule.get(projectId) ?? new Map();
    const fresh = new Map<string, WordsEntry>();

    for (const module of modules) {
        // Name-derived, so it cannot be read from a cache of the TEXT: two empty document
        // modules have the same source and different names.
        if (module.type === 'class' || module.type === 'document') {
            types.add(module.moduleName);
        }

        const key = module.moduleName.toLowerCase();
        const was = held.get(key);
        const entry = was?.source === module.source ? was : { source: module.source, words: wordsOf(module.source) };
        fresh.set(key, entry);

        for (const name of entry.words.procedures) { procedures.add(name); }
        for (const name of entry.words.types) { types.add(name); }
    }

    wordsByModule.set(projectId, fresh);
    return { types: [...types].sort(), procedures: [...procedures].sort() };
}

/** Drops a project's cached words. Called when the project closes. */
export function forgetProjectWords(projectId: string): void {
    wordsByModule.delete(projectId);
}

interface WordsEntry {
    source: string;
    words: { types: string[]; procedures: string[] };
}

/**
 * What each module's text contributes, kept against the exact text it came from.
 *
 * Every seed asks this of every module, and a seed happens on every pass that found anything
 * changed - so a one-line edit re-parsed the WHOLE project to rediscover words that only one
 * module's had moved. Per project and per module, so two open workbooks do not evict each other
 * and a module that is deleted does not linger.
 */
const wordsByModule = new Map<string, Map<string, WordsEntry>>();

function wordsOf(source: string): { types: string[]; procedures: string[] } {
    const types: string[] = [];
    const procedures: string[] = [];

    try {
        for (const member of parseModule(source).members) {
            if (member.kind === 'Procedure') {
                procedures.push(member.name);
            } else if (member.kind === 'Type' || member.kind === 'Enum') {
                const name = 'name' in member && typeof member.name === 'string' ? member.name : undefined;
                if (name) {
                    types.push(name);
                }
            }
        }
    } catch {
        // A module that will not parse lends no words; the rest still do. Held as none, so one
        // that cannot parse is not re-parsed on every seed to fail the same way each time.
    }

    return { types, procedures };
}
