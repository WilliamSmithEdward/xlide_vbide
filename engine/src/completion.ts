// Serves completions by reusing the editor extension's analyzer resolvers.
//
// The resolvers are pure: source text in, plain completion data out, with project facts passed as
// data. The facts come from the shared assembly in moduleContext.ts — the same assembly hover
// uses — so every feature describes the same project. Reusing the resolvers keeps one completion
// behaviour shared between the two products instead of two that drift.

import {
    resolveIdentifierCompletions,
    resolveKeywordCompletions,
    resolveMemberCompletions,
    resolveTypeCompletions,
    type IdentifierCompletion,
    type KeywordCompletion,
    type MemberCompletion,
    type TypeCompletion,
} from '../../../xlide_vscode/src/analyzer';
import { assembleContext } from './moduleContext';
import type { CompletionItemPayload, CompletionParams, ModulePayload } from './protocol';

export function completionsFor(
    seeded: readonly ModulePayload[],
    params: CompletionParams & { source: string },
): CompletionItemPayload[] {
    const ctx = assembleContext(seeded, params);

    // In a declaration's type position — after As, most visibly — type names are the whole
    // answer: built-ins, the host model's objects, and the project's own classes and types.
    // First, the way the extension orders it; this resolver was simply never called here, and
    // "Dim i As Int" offered nothing (2026-08-04).
    const types = resolveTypeCompletions(params.source, params.offset, {
        projectTypes: ctx.projectTypes,
    });

    if (types.length > 0) {
        return types.map(typeItem);
    }

    // After a dot, the members of the receiver are the only sensible answer.
    const members = resolveMemberCompletions(params.source, params.offset, {
        codeNames: ctx.codeNames,
        meType: ctx.meType,
        meProjectType: ctx.meProjectType,
        projectClassMembers: ctx.projectClassMembers,
    });

    if (members.length > 0) {
        return members.map(memberItem);
    }

    // A grammar position that admits only keywords shows only keywords; anywhere else, the
    // identifiers in scope and the keywords share the list, the way the extension offers them.
    const keywords = resolveKeywordCompletions(params.source, params.offset);
    if (keywords.exclusive) {
        return keywords.items.map(keywordItem);
    }

    const identifiers = resolveIdentifierCompletions(params.source, params.offset, {
        codeNames: ctx.codeNameList,
        moduleName: ctx.current.name,
        moduleKind: ctx.moduleKind,
        projectMemberSurfaces: ctx.projectClassMembers,
        projectProcedures: ctx.projectProcedures,
        projectSymbols: ctx.projectSymbols,
    });

    return [...identifiers.map(identifierItem), ...keywords.items.map(keywordItem)];
}

function typeItem(type: TypeCompletion): CompletionItemPayload {
    return {
        label: type.name,
        kind: type.kind,
        detail: type.detail,
        documentation: type.documentation,
    };
}

function memberItem(member: MemberCompletion): CompletionItemPayload {
    return {
        label: member.name,
        kind: member.kind,
        detail: member.signature ?? (member.returns ? `${member.owner}.${member.name} As ${member.returns}` : member.owner),
        documentation: member.documentation,
    };
}

function identifierItem(identifier: IdentifierCompletion): CompletionItemPayload {
    return {
        label: identifier.name,
        kind: identifier.kind,
        detail: identifier.detail,
        documentation: identifier.documentation,
    };
}

function keywordItem(keyword: KeywordCompletion): CompletionItemPayload {
    return {
        label: keyword.label,
        kind: 'keyword',
        detail: keyword.detail,
        documentation: keyword.documentation,
        insertText: keyword.insertText,
        filterText: keyword.filterText,
        sortText: keyword.sortText,
    };
}
