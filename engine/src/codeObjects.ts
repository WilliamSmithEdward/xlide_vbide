// Serves the code window's Object and Procedure lists (xlide_vbide#101): the objects a module can
// handle events of, and each one's events with their parameters as a handler declares them.
//
// The native editor lists, besides (General): a document module's own object (Workbook,
// Worksheet, Chart, Word's Document), a class module's Class, a UserForm and each of its controls,
// and every module-level WithEvents variable. The events come from where the analyzer's own
// handler-signature rule reads them - the Office type libraries' sweep (HOST_EVENT_SIGNATURES), a
// project class's declared Events, and the VBA UserForm's extras - so the stub a pick writes is
// one that rule accepts.

import { parseModule } from '../../../xlide_vscode/src/analyzer';
import { HOST_EVENT_SIGNATURES } from '../../../xlide_vscode/src/analyzer/host/eventSignaturesData';
import { resolveHostAlias } from '../../../xlide_vscode/src/analyzer/host/hostModel';
import { hostApp } from './hostApp';
import { assembleContext } from './moduleContext';
import type { CodeObject, CodeObjectEvent, CodeObjectsParams, ModulePayload } from './protocol';

/** Each class's events by lowercased qualified name, in the type library's order. */
const EVENTS_BY_CLASS = new Map<string, { name: string; events: CodeObjectEvent[] }>();
for (const [className, events] of Object.entries(HOST_EVENT_SIGNATURES)) {
    EVENTS_BY_CLASS.set(className.toLowerCase(), {
        name: className,
        events: Object.entries(events).map(([name, params]) => ({ name, params })),
    });
}

/** What VBA's own UserForm adds to the MSForms one, which the type-library sweep cannot see. */
const VBA_USERFORM_EVENTS: CodeObjectEvent[] = [
    { name: 'Activate', params: '' },
    { name: 'Deactivate', params: '' },
    { name: 'Initialize', params: '' },
    { name: 'QueryClose', params: 'Cancel As Integer, CloseMode As Integer' },
    { name: 'Resize', params: '' },
    { name: 'Terminate', params: '' },
];

/** A class module's own two events. */
const CLASS_EVENTS: CodeObjectEvent[] = [
    { name: 'Initialize', params: '' },
    { name: 'Terminate', params: '' },
];

/**
 * A document module's own object, the name its handlers take, by the module's Me type. The Me
 * type is the engine's one answer to what a document module is: the host never sends a document
 * type, so the engine infers it from the host and the module, and an unknown host gets no object.
 */
const DOCUMENT_OBJECTS: Readonly<Record<string, string>> = {
    'excel.workbook': 'Workbook',
    'excel.worksheet': 'Worksheet',
    'excel.chart': 'Chart',
    'word.document': 'Document',
};

/** Events of several classes as one list, the first spelling of a name kept, sorted as the native list is. */
function eventsOf(...classes: string[]): CodeObjectEvent[] {
    const byName = new Map<string, CodeObjectEvent>();
    for (const className of classes) {
        const entry = EVENTS_BY_CLASS.get(className.toLowerCase());
        const library = entry?.name.split('.')[0] ?? '';
        for (const event of entry?.events ?? []) {
            if (!byName.has(event.name.toLowerCase())) {
                byName.set(event.name.toLowerCase(), { name: event.name, params: qualified(event.params, library) });
            }
        }
    }
    return sorted([...byName.values()]);
}

/**
 * Each MSForms class's default event, the one its type library flags FUNCFLAG_FUIDEFAULT (read
 * from FM20.DLL). Image's default, Click, is a hidden member the list leaves out, so an Image
 * takes the first event the list shows.
 */
const MSFORMS_DEFAULT_EVENTS: Readonly<Record<string, string>> = {
    'msforms.userform': 'Click',
    'msforms.commandbutton': 'Click',
    'msforms.textbox': 'Change',
    'msforms.combobox': 'Change',
    'msforms.listbox': 'Click',
    'msforms.checkbox': 'Click',
    'msforms.optionbutton': 'Click',
    'msforms.togglebutton': 'Click',
    'msforms.frame': 'Click',
    'msforms.label': 'Click',
    'msforms.multipage': 'Change',
    'msforms.tabstrip': 'Change',
    'msforms.scrollbar': 'Change',
    'msforms.spinbutton': 'Change',
};

/**
 * The event the native editor writes a handler for when an object is picked in the Object list:
 * the class's flagged default, or, for the Office applications' classes, which flag none, the
 * first event of the interface - Worksheet's SelectionChange, Workbook's Open, Document's New.
 */
function defaultEventOf(className: string): string | undefined {
    const key = className.toLowerCase();
    return MSFORMS_DEFAULT_EVENTS[key] ?? EVENTS_BY_CLASS.get(key)?.events[0]?.name;
}

/** Members of a form's designer that are not controls: a MultiPage's pages, a TabStrip's tabs. */
const NOT_CONTROLS = new Set(['page', 'tab']);

/**
 * The extender events (Enter, Exit, BeforeUpdate, AfterUpdate) a control does not have, measured
 * against the VBE's own CreateEventProc, which refuses them: only a control that holds a value
 * updates, and a Label or an Image never takes the focus. A control not named here has all four.
 */
const EXTENDER_EVENTS_ABSENT: Readonly<Record<string, readonly string[]>> = {
    commandbutton: ['AfterUpdate', 'BeforeUpdate'],
    frame: ['AfterUpdate', 'BeforeUpdate'],
    multipage: ['AfterUpdate', 'BeforeUpdate'],
    tabstrip: ['AfterUpdate', 'BeforeUpdate'],
    label: ['AfterUpdate', 'BeforeUpdate', 'Enter', 'Exit'],
    image: ['AfterUpdate', 'BeforeUpdate', 'Enter', 'Exit'],
};

/** A type's name without its library: `MSForms.Label` is `label`. */
const bareType = (type: string): string => type.slice(type.lastIndexOf('.') + 1).toLowerCase();

/** The library each Office host's own types come from. */
const HOST_LIBRARIES: Readonly<Record<string, string>> = {
    excel: 'Excel',
    word: 'Word',
    powerpoint: 'PowerPoint',
    access: 'Access',
    outlook: 'Outlook',
};

const INTRINSIC_TYPES = new Set([
    'boolean', 'byte', 'currency', 'date', 'decimal', 'double', 'integer', 'long', 'longlong',
    'longptr', 'object', 'single', 'string', 'variant',
]);

/**
 * A handler's parameters as the native editor writes them: a type from the host application's
 * own library bare, any other library's type with the library in front - `ByVal Target As
 * Range` in Excel, but `ByVal Cancel As MSForms.ReturnBoolean` and `ByVal SyncEventType As
 * Office.MsoSyncEventType` (measured against the VBE's own CreateEventProc). The type-library
 * sweep keeps only the bare names; the one type it carries from outside its class's own
 * library is Office's MsoSyncEventType, so an Mso type is Office's and every other type is the
 * class's library's.
 */
function qualified(params: string, library: string): string {
    const host = HOST_LIBRARIES[hostApp()];
    return params.replace(/\bAs\s+([A-Za-z_]\w*)\b(?!\.)/g, (whole, type: string) => {
        if (INTRINSIC_TYPES.has(type.toLowerCase())) {
            return whole;
        }
        const owner = /^Mso[A-Z]/.test(type) ? 'Office' : library;
        return owner && owner !== host ? `As ${owner}.${type}` : whole;
    });
}

const sorted = (events: CodeObjectEvent[]): CodeObjectEvent[] =>
    [...events].sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

export function codeObjectsFor(
    seeded: readonly ModulePayload[],
    params: CodeObjectsParams & { source: string },
): CodeObject[] {
    const ctx = assembleContext(seeded, params);
    const objects: CodeObject[] = [];
    const add = (name: string, events: CodeObjectEvent[], preferred: string | undefined): void => {
        if (events.length > 0) {
            // The default the class names, while the list carries it; otherwise the first.
            const fallback = events[0]?.name ?? '';
            const defaultEvent = events.find((one) => one.name.toLowerCase() === preferred?.toLowerCase())?.name ?? fallback;
            objects.push({ name, events, defaultEvent });
        }
    };

    const documentObject = ctx.meType ? DOCUMENT_OBJECTS[ctx.meType.toLowerCase()] : undefined;
    if (ctx.moduleKind === 'document' && ctx.meType && documentObject) {
        add(documentObject, eventsOf(ctx.meType), defaultEventOf(ctx.meType));
    } else if (ctx.moduleKind === 'class') {
        add('Class', CLASS_EVENTS, 'Initialize');
    } else if (ctx.moduleKind === 'userform') {
        const form = new Map(eventsOf('MSForms.UserForm').map((event) => [event.name.toLowerCase(), event]));
        for (const extra of VBA_USERFORM_EVENTS) {
            form.set(extra.name.toLowerCase(), extra);
        }
        add('UserForm', sorted([...form.values()]), defaultEventOf('MSForms.UserForm'));

        // Each control, with the extender events it has. A MultiPage's pages and a TabStrip's tabs
        // are named members of the form but not controls, and raise nothing of their own.
        for (const control of [...(ctx.implicitMembers ?? [])]
            .filter((member) => !NOT_CONTROLS.has(bareType(member.type)))
            .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }))) {
            const extender = eventsOf('MSForms.Control')
                .filter((event) => !(EXTENDER_EVENTS_ABSENT[bareType(control.type)] ?? []).includes(event.name));
            const events = new Map(eventsOf(control.type).map((event) => [event.name.toLowerCase(), event]));
            for (const event of extender) {
                events.set(event.name.toLowerCase(), event);
            }
            add(control.name, sorted([...events.values()]), defaultEventOf(control.type));
        }
    }

    // Every module-level WithEvents variable: a project class's declared Events first, as VBA
    // resolves a name, then a library class's.
    let members: ReturnType<typeof parseModule>['members'] = [];
    try {
        members = parseModule(params.source).members;
    } catch {
        // A module that will not parse lists its own object and controls only.
    }
    for (const member of members) {
        if (member.kind !== 'VariableGroup' || !member.withEvents) {
            continue;
        }
        for (const decl of member.declarations) {
            const fromProject = projectEventsOf(decl.asType, ctx.projectClassMembers);
            if (fromProject) {
                add(decl.name, sorted(fromProject), fromProject[0]?.name);
                continue;
            }
            const className = decl.asType
                ? EVENTS_BY_CLASS.get(decl.asType.toLowerCase())?.name
                    ?? resolveHostAlias(decl.asType, ctx.hostModel)
                : undefined;
            if (className) {
                add(decl.name, eventsOf(className), defaultEventOf(className));
            }
        }
    }

    // By name, as the native list orders them - a form's UserForm among its controls, a class's
    // Class among its WithEvents variables (read off the VBE's own list, 2026-10-12).
    return objects.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
}

/** The Events a project class declares, in declaration order, when a WithEvents variable's type names one. */
function projectEventsOf(
    asType: string | undefined,
    projectClassMembers: ReturnType<typeof assembleContext>['projectClassMembers'],
): CodeObjectEvent[] | undefined {
    const bare = asType?.trim().slice(asType.trim().lastIndexOf('.') + 1).toLowerCase();
    const type = bare ? (projectClassMembers ?? []).find((candidate) => candidate.name.toLowerCase() === bare) : undefined;
    if (!type) {
        return undefined;
    }

    const events: CodeObjectEvent[] = [];
    for (const member of type.members) {
        const signature = member.signature ?? '';
        const open = signature.indexOf('(');
        if (member.kind === 'event' && open >= 0 && signature.endsWith(')')) {
            events.push({ name: member.name, params: signature.slice(open + 1, -1) });
        }
    }
    return events;
}
