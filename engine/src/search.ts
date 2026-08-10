// Finds text across the modules the engine already holds.
//
// The engine is the one place every open workbook's every module exists as text - seeded at
// analysis, overlaid by the live keystroke stream - so scoped search belongs here rather than
// in the host, which would have to read every module through the object model per search.

import type { SearchMatchPayload, SearchParams, SearchResult } from './protocol';

/** Enough for a results list; past this the query needs narrowing, not scrolling. */
const MATCH_LIMIT = 500;

// Any letter, not an ASCII one. Whole-word search over `Calculér` matched `Calcul` and then
// refused it for having a letter after it, so whole-word search could not find a name with an
// accent in it (2026-08-09). The `u` flag is required or \p{L} matches nothing.
const WORD_CHARACTER = /[\p{L}\p{M}\p{N}_]/u;

export function searchModules(
    modules: Iterable<{ projectId: string; module: string; source: string }>,
    params: SearchParams,
): SearchResult {
    const matches: SearchMatchPayload[] = [];
    const query = params.query;

    if (query.length === 0) {
        return { matches, truncated: false };
    }

    const needle = params.matchCase ? query : query.toLowerCase();

    for (const entry of modules) {
        const lines = entry.source.split(/\r\n|\n|\r/);

        for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            const line = lines[lineIndex];
            const haystack = params.matchCase ? line : line.toLowerCase();

            let from = 0;
            while (from <= haystack.length - needle.length) {
                const at = haystack.indexOf(needle, from);
                if (at < 0) {
                    break;
                }

                from = at + 1;

                if (params.wholeWord) {
                    const before = at > 0 ? line[at - 1] : undefined;
                    const after = at + needle.length < line.length ? line[at + needle.length] : undefined;
                    if ((before !== undefined && WORD_CHARACTER.test(before))
                        || (after !== undefined && WORD_CHARACTER.test(after))) {
                        continue;
                    }
                }

                matches.push({
                    projectId: entry.projectId,
                    module: entry.module,
                    line: lineIndex + 1,
                    column: at + 1,
                    length: query.length,
                    preview: line.length > 200 ? line.slice(0, 200) : line,
                });

                if (matches.length >= MATCH_LIMIT) {
                    return { matches, truncated: true };
                }
            }
        }
    }

    return { matches, truncated: false };
}
