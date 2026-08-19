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

import { existsSync } from 'node:fs';
import { join } from 'node:path';
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

/**
 * A DESIGNATED DEVIATION from the shared planner: a .frm whose .frx sits beside it is a
 * CREATE here, not a refusal.
 *
 * Their planner refuses to create a userform from a file because THEIR applier cannot -
 * `moduleSyncPlan.ts` encodes the capabilities of the extension's own import path. This
 * product's applier is the add-in, which hands the pair to `VBComponents.Import` and gets the
 * whole form back natively - controls, fonts, pictures - a path the designer suite has pinned
 * since 2026-08-16. The refusal reappeared when the engine was rebuilt against xlide_vscode
 * 4.0.0 (whose #21 classifies .frm as userform; the older engine binary predated that and
 * planned the pair as an ordinary create), caught by the suite going 431/6 (2026-08-19, hunt
 * round three).
 *
 * The promotion is deliberately narrow: import direction, a `skipping-import` row, a `.frm`
 * file, and the sidecar PRESENT in the folder - a .frm alone stays refused, because importing
 * it would fail at the VBE with less to say than the planner's warning already says. The ask
 * for a caller-declared capability is filed as xlide_vscode#27, so this layer can retire.
 */
function withFormPairsCreatable(plan: ModuleSyncPlan, folder: string): ModuleSyncPlan {
    const rows = plan as unknown as {
        items?: {
            status?: string;
            /** The planner's own field for the file inside the folder. */
            relativeName?: string;
            moduleName?: string;
            checked?: boolean;
            warning?: string;
            detail?: string;
            unsupportedDirectCreation?: boolean;
            rightTitle?: string;
        }[];
    };
    if (!Array.isArray(rows.items)) {
        return plan;
    }

    return {
        ...plan,
        items: rows.items.map((item) => {
            const file = item.relativeName ?? '';
            if (item.status !== 'skipping-import' || !/\.frm$/i.test(file)) {
                return item;
            }

            const sidecar = `${file.slice(0, -'.frm'.length)}.frx`;
            if (!existsSync(join(folder, sidecar))) {
                return item;
            }

            return {
                ...item,
                status: 'will-create',
                checked: true,
                warning: undefined,
                detail: 'Will create',
                unsupportedDirectCreation: false,
                rightTitle: `File: ${item.moduleName ?? file} (will create)`,
            };
        }),
    } as ModuleSyncPlan;
}

/** Works out what an import or export would do, using the companion editor's own planner. */
export async function syncPlan(params: SyncPlanParams): Promise<ModuleSyncPlan> {
    const bridge = bridgeOver(params.modules) as never;

    const plan = params.direction === 'import'
        ? withFormPairsCreatable(await buildImportModuleSyncPlan(bridge, {
            workbookPath: params.workbookPath,
            importFolder: params.folder,
            importMode: params.mode === 'trueUpStandardClass' ? 'trueUpStandardClass' : 'updateOnly',
            folderPathSource: 'session',
            importModeSource: 'session',
        }), params.folder)
        : await buildExportModuleSyncPlan(bridge, {
            workbookPath: params.workbookPath,
            exportFolder: params.folder,
            exportMode: params.mode === 'trueUp' ? 'trueUp' : 'exportAll',
            folderPathSource: 'session',
            exportModeSource: 'session',
        });

    return withoutPointlessComparisons(plan);
}
