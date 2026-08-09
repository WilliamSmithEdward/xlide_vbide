/*
 * Import and export, decided by the COMPANION EDITOR'S OWN CODE.
 *
 * `moduleSyncPlan.ts` is imported straight out of the xlide_vscode checkout, exactly the way the
 * analyzer is. That is the whole point of this file: the decisions that have to agree between the
 * two products, namely what a module's file is called, what kind of module a .cls holds, what counts as
 * a change, which files are stale, are made ONCE, by one implementation, and a fix to it fixes
 * both. Anything in here that reimplemented those decisions would be a second opinion waiting to
 * drift.
 *
 * What it needed to be handed instead of a workbook:
 *
 * Their planner reads modules through a bridge, `bridge.call('readModules', { path, full: true })`,
 * which in the extension opens the .xlsm on disk. Here the workbook is OPEN IN EXCEL and the file
 * on disk is stale by definition, so the shim reads the live project over COM and sends the
 * modules with the request. The stand-in below answers that one call from what arrived. Nothing
 * else about their planner changes, and nothing here touches the workbook file.
 */

import {
    buildExportModuleSyncPlan,
    buildImportModuleSyncPlan,
    type ModuleSyncPlan,
} from '../../../xlide_vscode/src/moduleSyncPlan';

/** One module as the shim read it out of the live project. */
export interface SyncLiveModule {
    name: string;
    /** standard, class, document or userform - their planner's own vocabulary. */
    type: string;
    /** The full text INCLUDING the attribute header, so a written file round-trips. */
    source: string;
    documentType?: string;
}

export interface SyncPlanParams {
    direction: 'export' | 'import';
    /** Only ever used as a label and as a lock key; nothing opens it. */
    workbookPath: string;
    folder: string;
    /** exportAll | trueUp, or updateOnly | trueUpStandardClass. */
    mode?: string;
    modules: SyncLiveModule[];
}

/**
 * The bridge their planner expects, with one call answered from what the shim sent.
 *
 * Deliberately narrow: `readModules` and `readModule` are the only two things the planning path
 * asks for, and answering anything else would be inventing a workbook this process cannot see. An
 * unexpected call throws rather than returning something plausible, so a future version of their
 * code that needs more fails loudly here instead of quietly planning against nothing.
 */
function bridgeOver(modules: SyncLiveModule[]): { call: <T>(method: string, params: unknown) => Promise<T> } {
    const byName = new Map(modules.map((mod) => [mod.name.toLowerCase(), mod]));

    return {
        call: async <T>(method: string, params: unknown): Promise<T> => {
            if (method === 'readModules') {
                return modules.map((mod) => ({
                    name: mod.name,
                    type: mod.type,
                    documentType: mod.documentType,
                    source: mod.source,
                })) as T;
            }

            if (method === 'readModule') {
                const wanted = String((params as { module?: string })?.module ?? '').toLowerCase();
                const found = byName.get(wanted);
                if (!found) {
                    throw new Error(`the live project has no module named ${wanted}`);
                }

                return { source: found.source } as T;
            }

            throw new Error(
                `module sync asked the workbook for '${method}', which this host cannot answer: `
                + 'the project is open in Excel and is read over COM, not from the file.');
        },
    };
}

/**
 * A row whose two sides are the same does not send its comparison down the pipe.
 *
 * The planner draws every row line by line, both texts, both line numbers, with the headers and
 * again without: for a project of 81,795 lines that is some 163,000 comparison entries, and the
 * pipe and the two JSON passes either side of it cost 1,417ms of a 1,710ms plan. Measured 2026-08-09
 * by returning the plan with the comparisons stripped: 1,710ms became 293ms.
 *
 * Only where there is NOTHING TO SHOW. An unchanged row is unchanged because the two texts are
 * equal, and drawing equality line by line produces a screenful that the dialog collapses back to
 * "N identical lines" anyway - which the shim writes from the row's own text, exactly as its
 * built-in planner already does. Every row with a real difference keeps its comparison, drawn by
 * this planner, because that is a picture of a decision and the decisions are theirs.
 */
function withoutPointlessComparisons(plan: ModuleSyncPlan): ModuleSyncPlan {
    const rows = plan as unknown as { items?: { status?: string }[] };
    if (!Array.isArray(rows.items)) {
        return plan;
    }

    return {
        ...plan,
        items: rows.items.map((item) =>
            item.status === 'unchanged' ? { ...item, diff: [], diffWithHeaders: [] } : item),
    } as ModuleSyncPlan;
}

/** Works out what an import or export would do, using the companion editor's own planner. */
export async function syncPlan(params: SyncPlanParams): Promise<ModuleSyncPlan> {
    const bridge = bridgeOver(params.modules) as never;

    const plan = params.direction === 'import'
        ? await buildImportModuleSyncPlan(bridge, {
            workbookPath: params.workbookPath,
            importFolder: params.folder,
            importMode: params.mode === 'trueUpStandardClass' ? 'trueUpStandardClass' : 'updateOnly',
            folderPathSource: 'session',
            importModeSource: 'session',
        })
        : await buildExportModuleSyncPlan(bridge, {
            workbookPath: params.workbookPath,
            exportFolder: params.folder,
            exportMode: params.mode === 'trueUp' ? 'trueUp' : 'exportAll',
            folderPathSource: 'session',
            exportModeSource: 'session',
        });

    return withoutPointlessComparisons(plan);
}
