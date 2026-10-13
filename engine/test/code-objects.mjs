// The code window's Object and Procedure lists over the pipe (xlide_vbide#101).
//
// The native Object list shows, besides (General), a document module's own object, a class's
// Class, a form and each of its controls, and every WithEvents variable; the Procedure list shows
// each object's events. The host never sends a document type, so a sheet's Worksheet object has
// to come from the engine's own reading of the module - the first live run found Sheet1 listing
// nothing at all, because the lookup read a field the host leaves empty.
//
//   node test/code-objects.mjs
//   node test/code-objects.mjs --exe

import assert from 'node:assert/strict';
import { reporter, startEngine } from './harness.mjs';

const { call, stop } = await startEngine('code-objects');
const { check, done } = reporter();

const APP_CLASS = [
    'Option Explicit',
    '',
    'Private WithEvents app As Application',
    'Public WithEvents tracker As Tracker',
    '',
].join('\r\n');

const TRACKER = [
    'Option Explicit',
    '',
    'Public Event Moved(ByVal steps As Long)',
    '',
].join('\r\n');

await call('initialize', {});
await call('project/open', {
    projectId: 'Lists',
    generation: 1,
    modules: [
        { moduleName: 'ThisWorkbook', source: 'Option Explicit\r\n', type: 'document' },
        { moduleName: 'Sheet1', source: 'Option Explicit\r\n', type: 'document' },
        { moduleName: 'Watcher', source: APP_CLASS, type: 'class' },
        { moduleName: 'Tracker', source: TRACKER, type: 'class' },
        { moduleName: 'Standard', source: 'Option Explicit\r\n', type: 'standard' },
        {
            moduleName: 'EntryForm',
            source: 'Option Explicit\r\n',
            type: 'userform',
            implicitMembers: [
                { name: 'OkButton', type: 'MSForms.CommandButton' },
                { name: 'Caption', type: 'MSForms.Label' },
                { name: 'NameBox', type: 'MSForms.TextBox' },
                { name: 'Steps', type: 'MSForms.MultiPage' },
                { name: 'Page1', type: 'MSForms.Page' },
            ],
        },
    ],
});

const objectsOf = async (moduleName, moduleType) => (await call('textDocument/codeObjects', {
    projectId: 'Lists',
    moduleName,
    moduleType,
})).objects;
const names = (objects) => objects.map((one) => one.name);
const eventsOf = (objects, name) => objects.find((one) => one.name === name)?.events.map((one) => one.name) ?? [];

const workbook = await objectsOf('ThisWorkbook', 'document');
check('ThisWorkbook lists Workbook, with Open and BeforeClose', () => {
    assert.deepEqual(names(workbook), ['Workbook']);
    assert.ok(eventsOf(workbook, 'Workbook').includes('Open'));
    assert.ok(eventsOf(workbook, 'Workbook').includes('BeforeClose'));
});

const sheet = await objectsOf('Sheet1', 'document');
check('Sheet1 lists Worksheet, with Change and SelectionChange', () => {
    assert.deepEqual(names(sheet), ['Worksheet']);
    assert.ok(eventsOf(sheet, 'Worksheet').includes('Change'));
    assert.ok(eventsOf(sheet, 'Worksheet').includes('SelectionChange'));
    const change = sheet[0].events.find((one) => one.name === 'Change');
    assert.equal(change.params, 'ByVal Target As Range');
});

check('the events are in the native list\'s alphabetical order', () => {
    const events = eventsOf(sheet, 'Worksheet');
    const ordered = [...events].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
    assert.deepEqual(events, ordered);
});

const watcher = await objectsOf('Watcher', 'class');
check('a class lists Class and each WithEvents variable', () => {
    assert.deepEqual(names(watcher), ['app', 'Class', 'tracker']);
    assert.deepEqual(eventsOf(watcher, 'Class'), ['Initialize', 'Terminate']);
    assert.ok(eventsOf(watcher, 'app').includes('WorkbookOpen'));
    assert.deepEqual(eventsOf(watcher, 'tracker'), ['Moved']);
});

const form = await objectsOf('EntryForm', 'userform');
check('a form lists UserForm and its controls', () => {
    assert.deepEqual(names(form), ['Caption', 'NameBox', 'OkButton', 'Steps', 'UserForm']);
    assert.ok(eventsOf(form, 'UserForm').includes('Initialize'));
    assert.ok(eventsOf(form, 'UserForm').includes('QueryClose'));
    assert.ok(eventsOf(form, 'OkButton').includes('Click'));
});

check('a type from outside the host\'s library is written with its library, as the VBE writes it', () => {
    const params = (objects, object, event) => objects.find((one) => one.name === object)
        .events.find((one) => one.name === event).params;
    assert.equal(params(workbook, 'Workbook', 'Sync'), 'ByVal SyncEventType As Office.MsoSyncEventType');
    assert.equal(params(workbook, 'Workbook', 'SheetChange'), 'ByVal Sh As Object, ByVal Target As Range');
    assert.equal(params(form, 'OkButton', 'KeyDown'), 'ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer');
    assert.equal(params(form, 'UserForm', 'AddControl'), 'ByVal Control As MSForms.Control');
    assert.equal(params(form, 'UserForm', 'QueryClose'), 'Cancel As Integer, CloseMode As Integer');
});

check('each object names the default event a pick of it writes, as the native list does', () => {
    const defaultOf = (objects, object) => objects.find((one) => one.name === object).defaultEvent;
    assert.equal(defaultOf(sheet, 'Worksheet'), 'SelectionChange');
    assert.equal(defaultOf(workbook, 'Workbook'), 'Open');
    assert.equal(defaultOf(watcher, 'Class'), 'Initialize');
    assert.equal(defaultOf(watcher, 'app'), 'NewWorkbook');
    assert.equal(defaultOf(watcher, 'tracker'), 'Moved');
    assert.equal(defaultOf(form, 'UserForm'), 'Click');
    assert.equal(defaultOf(form, 'OkButton'), 'Click');
});

check('a control has only the extender events the VBE accepts, and a page is no object', () => {
    assert.ok(!names(form).includes('Page1'));
    assert.deepEqual(eventsOf(form, 'Caption').filter((one) => ['Enter', 'Exit', 'BeforeUpdate', 'AfterUpdate'].includes(one)), []);
    assert.deepEqual(eventsOf(form, 'OkButton').filter((one) => ['Enter', 'Exit', 'BeforeUpdate', 'AfterUpdate'].includes(one)), ['Enter', 'Exit']);
    assert.deepEqual(eventsOf(form, 'NameBox').filter((one) => ['Enter', 'Exit', 'BeforeUpdate', 'AfterUpdate'].includes(one)), ['AfterUpdate', 'BeforeUpdate', 'Enter', 'Exit']);
    assert.equal(form.find((one) => one.name === 'NameBox').defaultEvent, 'Change');
});

const standard = await objectsOf('Standard', 'standard');
check('a standard module lists nothing beyond (General)', () => assert.deepEqual(standard, []));

stop();
process.exit(done());
