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

/** Works out what an import or export would do, using the companion editor's own planner. */
export async function syncPlan(params: SyncPlanParams): Promise<ModuleSyncPlan> {
    const bridge = bridgeOver(params.modules) as never;

    return params.direction === 'import'
        ? buildImportModuleSyncPlan(bridge, {
            workbookPath: params.workbookPath,
            importFolder: params.folder,
            importMode: params.mode === 'trueUpStandardClass' ? 'trueUpStandardClass' : 'updateOnly',
            folderPathSource: 'session',
            importModeSource: 'session',
        })
        : buildExportModuleSyncPlan(bridge, {
            workbookPath: params.workbookPath,
            exportFolder: params.folder,
            exportMode: params.mode === 'trueUp' ? 'trueUp' : 'exportAll',
            folderPathSource: 'session',
            exportModeSource: 'session',
        });
}
