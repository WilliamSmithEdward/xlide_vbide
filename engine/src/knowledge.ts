// What the language service KNOWS, served as data.
//
// Two catalogues, both already living in the analyzer and until now readable only by being a
// language feature: the host object model (the types and members completions and hover draw on)
// and the diagnostic rule set (what the analyzer checks, at which severity, on whose authority).
// An agent working through the api needs both to reason about what the service will and will not
// tell it - "is this member unknown or is the whole type unknown?" is unanswerable without the
// model, and "will this construct be flagged?" is unanswerable without the rules.
//
// EVERYTHING HERE IS A READ OF THE ANALYZER'S OWN TABLES. Nothing is restated: the reply is
// projected from the same objects the resolvers use, so it cannot drift from what the features
// actually know. The one editorial act is SHRINKING - members carry their doc's one-line summary
// rather than the whole doc, because the reply is an inventory, not the reference.
//
// The model answers for the host this engine was told it is running in (project/open's token),
// through the analyzer's host registry: Excel, Word, PowerPoint and Access carry models as of
// xlide_vscode 4.0.0. A host the registry does not know answers known:false with a note, which
// is the same honesty the features themselves practice: silence rather than another
// application's members.

import { currentHostModel, hostApp, hostModelIsKnown } from './hostApp.js';
import {
    DIAGNOSTIC_RULES,
    STRUCTURAL_DIAGNOSTIC_RULES,
    type HostMember,
    type HostObjectModel,
    type HostType,
} from '../../../xlide_vscode/src/analyzer';

/**
 * The rule fields this inventory projects, spelled structurally rather than imported: the two
 * catalogues (incremental and structural) are distinct literal types that both satisfy this,
 * and naming only what is read keeps the projection honest about what it depends on.
 */
interface RuleShape {
    code: string;
    title: string;
    defaultSeverity: string;
    category: string;
    diagnosticKind: string;
    vbeCompileEquivalent: boolean;
    confidence: string;
    specReference?: string;
    requiresWholeProject?: boolean;
}
import type {
    KnowledgeAnalyzerResult,
    KnowledgeAnalyzerRuleRow,
    KnowledgeModelMemberRow,
    KnowledgeModelParams,
    KnowledgeModelResult,
    KnowledgeModelTypeRow,
} from './protocol';

/**
 * The host object model as an inventory: every type with its member count, or - when `type`
 * names one - that type's members in full.
 */
export function objectModelKnowledge(params: KnowledgeModelParams): KnowledgeModelResult {
    const host = hostApp();
    if (!hostModelIsKnown()) {
        return {
            host,
            known: false,
            note:
                `No object model is wired for '${host}', so the language service asserts `
                + 'nothing about this host\'s own types - document modules get no host members '
                + 'rather than another host\'s.',
        };
    }

    const model = currentHostModel();

    if (params.type !== undefined && params.type.length > 0) {
        const resolved = resolveType(model, params.type);
        if (resolved === undefined) {
            return {
                host,
                known: true,
                note:
                    `No type named '${params.type}'. Names resolve as written in code, `
                    + 'qualified, or through the model\'s aliases; the bare inventory lists '
                    + 'every qualified name.',
            };
        }

        const [qualified, held] = resolved;
        return {
            host,
            known: true,
            source: model.source,
            type: {
                name: qualified,
                displayName: held.displayName,
                exhaustive: held.exhaustive === true,
                memberCount: held.members.length,
                members: held.members.map(memberRow),
            },
        };
    }

    const names = Object.keys(model.types).sort();
    const types: KnowledgeModelTypeRow[] = names.map((name) => {
        const held = model.types[name]!;
        return {
            name,
            displayName: held.displayName,
            members: held.members.length,
            exhaustive: held.exhaustive === true,
        };
    });

    return {
        host,
        known: true,
        source: model.source,
        typeCount: types.length,
        types,
        globals: { ...model.globals },
        aliasCount: Object.keys(model.aliases).length,
        constantCount: model.constants === undefined ? 0 : Object.keys(model.constants).length,
    };
}

/** The diagnostic rule catalogue: every rule, incremental and structural both. */
export function analyzerKnowledge(): KnowledgeAnalyzerResult {
    const rules: KnowledgeAnalyzerRuleRow[] = [];

    for (const [key, rule] of Object.entries(DIAGNOSTIC_RULES)) {
        rules.push(ruleRow(key, rule));
    }

    for (const [key, rule] of Object.entries(STRUCTURAL_DIAGNOSTIC_RULES)) {
        rules.push(ruleRow(key, rule));
    }

    rules.sort((left, right) => left.code.localeCompare(right.code));

    const categories = [...new Set(rules.map((rule) => rule.category))].sort();

    return { ruleCount: rules.length, categories, rules };
}

/**
 * A named type, resolved the way code reaches it: as written (`Worksheet`, through the alias
 * table), by exact qualified key, or - host-neutrally - by its bare display name, so
 * `Document` finds `Word.Document` in Word without this file knowing any host's prefix.
 */
function resolveType(model: HostObjectModel, name: string): [string, HostType] | undefined {
    const aliased = model.aliases[name.toLowerCase()];
    for (const candidate of [aliased, name]) {
        if (candidate !== undefined && candidate in model.types) {
            return [candidate, model.types[candidate]!];
        }
    }

    const wanted = name.toLowerCase();
    for (const [qualified, held] of Object.entries(model.types)) {
        if (held.displayName.toLowerCase() === wanted) {
            return [qualified, held];
        }
    }

    return undefined;
}

function memberRow(member: HostMember): KnowledgeModelMemberRow {
    const row: KnowledgeModelMemberRow = { name: member.name };
    if (member.kind !== undefined) {
        row.kind = member.kind;
    }
    if (member.returns !== undefined) {
        row.returns = member.returns;
    }
    if (member.signature !== undefined) {
        row.signature = member.signature;
    }
    if (member.doc?.summary !== undefined) {
        row.doc = member.doc.summary;
    }
    if (member.hidden === true) {
        // The type library's own hidden/restricted attribute (xlide_vscode#56): completion
        // upstream already declines these, and a consumer deciding what to OFFER - an Object
        // Browser's "show hidden members" switch - filters on this rather than on the naming
        // convention that happens to hold for Excel's library.
        row.hidden = true;
    }
    return row;
}

function ruleRow(key: string, rule: RuleShape): KnowledgeAnalyzerRuleRow {
    const row: KnowledgeAnalyzerRuleRow = {
        key,
        code: rule.code,
        title: rule.title,
        severity: rule.defaultSeverity,
        category: rule.category,
        kind: rule.diagnosticKind,
        compileEquivalent: rule.vbeCompileEquivalent,
        confidence: rule.confidence,
    };
    if (rule.specReference !== undefined) {
        row.spec = rule.specReference;
    }
    if (rule.requiresWholeProject === true) {
        row.wholeProject = true;
    }
    return row;
}
