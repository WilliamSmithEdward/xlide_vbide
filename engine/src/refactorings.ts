/*
 * The lightbulb's question, asked of the planners themselves: would this refactoring go through
 * here, as its entry would ask for it?
 *
 * Nothing is kept. Each planner runs exactly as choosing its entry would run it, and only the
 * refusal comes back - so the lightbulb and the refactoring cannot disagree about what applies,
 * because there is one implementation of "applies" and it is the refactoring's own. A lighter
 * check written for the bulb would drift from its planner the first time a refusal was added to
 * one and not the other.
 *
 * Cheap enough for every caret move. The analyzer memoises its parse on the text itself, and the
 * workbook's symbols are built only when a planner reaches for them: every edit invalidates them,
 * and building them for a 63,000-line module cost 170ms, which the first caret settle after each
 * keystroke paid even on a keyword (measured 2026-09-10). Inline and Make-a-parameter refuse a
 * keyword or any other non-local before they need symbols, so they are handed a way to build
 * them rather than the symbols themselves.
 */

import type { ProjectSymbols } from './navigation';
import type { ModulePayload, RefactoringVerdict } from './protocol';
import { encapsulateFieldFor } from './encapsulateField';
import { extractMethodFor } from './extractMethod';
import { extractVariableFor } from './extractVariable';
import { implementInterfaceFor } from './implementInterface';
import { inlineVariableFor } from './inlineVariable';
import { introduceParameterFor } from './introduceParameter';
import { moveToModuleOffered } from './moveToModule';

/** The workbook's symbols, or the way to build them when a planner first needs them. */
export type SymbolsWhenNeeded = ProjectSymbols | (() => ProjectSymbols);

/** The most one caret raises; more is a malformed request rather than a busy line. */
const MOST_CANDIDATES = 16;

/** Each candidate's verdict, at the candidate's own position in the request. */
export function refactoringVerdicts(
    symbols: SymbolsWhenNeeded,
    seeded: readonly ModulePayload[],
    projectId: string,
    moduleName: string,
    source: string,
    candidates: readonly unknown[],
): RefactoringVerdict[] {
    // Built at most once for the whole request, however many planners reach for them.
    let built: ProjectSymbols | null = typeof symbols === 'function' ? null : symbols;
    const symbolsNow = (): ProjectSymbols => {
        built ??= (symbols as () => ProjectSymbols)();
        return built;
    };

    return candidates.slice(0, MOST_CANDIDATES).map((candidate) => {
        const kind = kindOf(candidate);
        try {
            return { kind, refused: refusalOf(symbolsNow, seeded, projectId, moduleName, source, candidate) };
        } catch (error) {
            // A planner that throws here would throw when its entry was chosen, too, and an entry
            // that fails on arrival is exactly what this exists to withhold.
            const why = error instanceof Error ? error.message : String(error);
            return { kind, refused: `The ${kind || 'refactoring'} check failed (${why}), so it is not offered.` };
        }
    });
}

/** Every candidate refused for one reason: a module the engine does not hold, for one. */
export function refusingEvery(candidates: readonly unknown[], why: string): RefactoringVerdict[] {
    return candidates.slice(0, MOST_CANDIDATES).map((candidate) => ({ kind: kindOf(candidate), refused: why }));
}

function refusalOf(
    symbolsNow: () => ProjectSymbols,
    seeded: readonly ModulePayload[],
    projectId: string,
    moduleName: string,
    source: string,
    candidate: unknown,
): string | null {
    const asked = (candidate ?? {}) as Record<string, unknown>;
    const kind = kindOf(candidate);
    const missing = `The ${kind} candidate is missing what it needs, so it is not offered.`;

    switch (kind) {
        case 'inlineVariable': {
            const offset = whole(asked.offset);
            return offset === null ? missing : inlineVariableFor(symbolsNow, moduleName, source, offset).refused ?? null;
        }

        case 'introduceParameter': {
            const offset = whole(asked.offset);
            return offset === null ? missing : introduceParameterFor(symbolsNow, moduleName, source, offset).refused ?? null;
        }

        case 'moveToModule': {
            const offset = whole(asked.offset);
            return offset === null ? missing : moveToModuleOffered(symbolsNow(), moduleName, source, offset);
        }

        case 'extractVariable': {
            const startOffset = whole(asked.startOffset);
            const endOffset = whole(asked.endOffset);
            if (startOffset === null || endOffset === null) {
                return missing;
            }

            return extractVariableFor(seeded, {
                projectId, moduleName, source, startOffset, endOffset, newName: unusedName(source, 'value'),
            }).refused ?? null;
        }

        case 'extractMethod': {
            const startLine = whole(asked.startLine);
            const endLine = whole(asked.endLine);
            if (startLine === null || endLine === null) {
                return missing;
            }

            return extractMethodFor(symbolsNow(), moduleName, source, startLine, endLine, unusedName(source, 'Extracted'))
                .refused ?? null;
        }

        case 'encapsulateField': {
            const fieldName = named(asked.fieldName);
            return fieldName === null ? missing : encapsulateFieldFor(moduleName, source, fieldName).refused ?? null;
        }

        case 'implementInterface': {
            const interfaceName = named(asked.interfaceName);
            return interfaceName === null
                ? missing
                : implementInterfaceFor(symbolsNow(), moduleName, source, interfaceName).refused ?? null;
        }

        default:
            return `'${kind}' is not a refactoring this engine knows.`;
    }
}

function kindOf(candidate: unknown): string {
    const kind = (candidate as { kind?: unknown } | null | undefined)?.kind;
    return typeof kind === 'string' ? kind : '';
}

function whole(value: unknown): number | null {
    return typeof value === 'number' && Number.isInteger(value) && value >= 0 ? value : null;
}

function named(value: unknown): string | null {
    return typeof value === 'string' && value.trim().length > 0 && value.length <= 255 ? value.trim() : null;
}

/**
 * `stemN`, counting up past every one the module already says - the way the page's own dialogs
 * suggest a name - so a planner asked about a selection is never answering about a name.
 */
function unusedName(source: string, stem: string): string {
    for (let number = 1; number < 100_000; number += 1) {
        const candidate = `${stem}${number}`;
        if (!new RegExp(`\\b${candidate}\\b`, 'i').test(source)) {
            return candidate;
        }
    }

    return `${stem}${Date.now()}`;
}
