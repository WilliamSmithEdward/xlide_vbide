// Proves the engine answers real analysis over a real pipe.
//
// Run against the bundle by default, or against the packaged executable:
//   node test/smoke.mjs
//   node test/smoke.mjs --exe

import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const dist = join(here, '..', 'dist');
const useExe = process.argv.includes('--exe');

const pipeName = `xlide-engine-smoke-${process.pid}`;
const command = useExe ? join(dist, 'xlide-engine.exe') : process.execPath;
const args = useExe ? ['--pipe', pipeName] : [join(dist, 'engine.cjs'), '--pipe', pipeName];

console.log(`starting ${useExe ? 'executable' : 'bundle'}: ${command}`);

const engine = spawn(command, args, { stdio: ['ignore', 'pipe', 'pipe'] });
engine.stderr.on('data', (chunk) => process.stderr.write(`engine stderr: ${chunk}`));

/** Resolves once the engine reports the pipe is open, so there is no race and no sleep. */
const listening = new Promise((resolve, reject) => {
    let seen = '';
    engine.stdout.on('data', (chunk) => {
        seen += chunk.toString();
        if (seen.includes('listening')) { resolve(); }
    });
    engine.on('exit', (code) => reject(new Error(`engine exited early with code ${code}`)));
    setTimeout(() => reject(new Error('engine did not report listening within 30s')), 30_000);
});

await listening;

const socket = net.connect(`\\\\.\\pipe\\${pipeName}`);
await new Promise((resolve, reject) => {
    socket.once('connect', resolve);
    socket.once('error', reject);
});

let nextId = 1;
const pending = new Map();
let inbox = '';

socket.on('data', (chunk) => {
    inbox += chunk.toString('utf8');
    let newline = inbox.indexOf('\n');
    while (newline >= 0) {
        const line = inbox.slice(0, newline).trim();
        inbox = inbox.slice(newline + 1);
        newline = inbox.indexOf('\n');
        if (!line) { continue; }

        const message = JSON.parse(line);
        const waiter = pending.get(message.id);
        if (waiter) {
            pending.delete(message.id);
            if (message.error) { waiter.reject(new Error(`${message.error.code}: ${message.error.message}`)); }
            else { waiter.resolve(message.result); }
        }
    }
});

function call(method, params) {
    const id = nextId++;
    return new Promise((resolve, reject) => {
        pending.set(id, { resolve, reject });
        socket.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
        setTimeout(() => {
            if (pending.delete(id)) { reject(new Error(`${method} timed out`)); }
        }, 30_000);
    });
}

const BAD_MODULE = [
    'Option Explicit',
    '',
    'Sub Probe()',
    '    Dim n As Long',
    '    n = "oops"',
    'End Sub',
    '',
].join('\r\n');

const GOOD_MODULE = [
    'Option Explicit',
    '',
    'Sub Fine()',
    '    Dim n As Long',
    '    n = 1',
    'End Sub',
    '',
].join('\r\n');

// Legal only because it is a class: Me, a Friend procedure, an event. Analysed as a standard
// module every line of it is an error, which is exactly the regression this pins.
const CLASS_MODULE = [
    'Option Explicit',
    '',
    'Public Event Renamed(ByVal newName As String)',
    '',
    'Private mName As String',
    '',
    'Friend Sub Adopt(ByVal name As String)',
    '    mName = name',
    '    RaiseEvent Renamed(mName)',
    'End Sub',
    '',
    'Public Property Get Name() As String',
    '    Name = mName',
    'End Property',
    '',
    'Public Sub Describe()',
    '    Debug.Print Me.Name',
    'End Sub',
    '',
].join('\r\n');

let failures = 0;
function check(name, body) {
    try {
        body();
        console.log(`  ok   ${name}`);
    } catch (error) {
        failures++;
        console.log(`  FAIL ${name}: ${error.message}`);
    }
}

try {
    const hello = await call('initialize', {});
    check('initialize reports the engine', () => assert.equal(hello.engine, 'xlide'));

    const opened = await call('project/open', {
        projectId: 'Smoke',
        generation: 1,
        modules: [
            { moduleName: 'BadModule', source: BAD_MODULE, type: 'standard' },
            { moduleName: 'GoodModule', source: GOOD_MODULE, type: 'standard' },
            { moduleName: 'FineClass', source: CLASS_MODULE, type: 'class' },
        ],
    });
    check('project/open accepts the modules', () => assert.equal(opened.modules, 3));

    check('project/open names the project types and procedures', () => {
        assert.deepEqual(opened.types, ['FineClass']);
        assert.deepEqual(opened.procedures, ['Adopt', 'Describe', 'Fine', 'Name', 'Probe']);
    });

    const bad = await call('textDocument/diagnostics', {
        documentKey: 'Smoke/BadModule',
        projectId: 'Smoke',
        generation: 1,
        source: BAD_MODULE,
        moduleName: 'BadModule',
        moduleType: 'standard',
    });

    console.log(`  -> ${bad.diagnostics.length} diagnostic(s) on the bad module:`);
    for (const diagnostic of bad.diagnostics) {
        console.log(`     [${diagnostic.severity}] ${diagnostic.code ?? '(structural)'} ` +
            `at ${diagnostic.span.start}..${diagnostic.span.end}: ${diagnostic.message}`);
    }

    check('assigning a string to a Long is reported', () => {
        assert.ok(bad.diagnostics.length > 0, 'expected at least one diagnostic');
    });

    check('the diagnostic points inside the module', () => {
        const first = bad.diagnostics[0];
        assert.ok(first.span.start >= 0 && first.span.end <= BAD_MODULE.length,
            `span ${first.span.start}..${first.span.end} outside 0..${BAD_MODULE.length}`);
    });

    const good = await call('textDocument/diagnostics', {
        documentKey: 'Smoke/GoodModule',
        projectId: 'Smoke',
        generation: 1,
        source: GOOD_MODULE,
        moduleName: 'GoodModule',
        moduleType: 'standard',
    });

    console.log(`  -> ${good.diagnostics.length} diagnostic(s) on the clean module`);
    check('clean code produces no findings', () => assert.equal(good.diagnostics.length, 0));

    // The module kind must reach the rules: a class using Me, Friend, and an event is legal as
    // a class and one long error list as anything else.
    const fineClass = await call('textDocument/diagnostics', {
        documentKey: 'Smoke/FineClass',
        projectId: 'Smoke',
        generation: 1,
        source: CLASS_MODULE,
        moduleName: 'FineClass',
        moduleType: 'class',
    });

    console.log(`  -> ${fineClass.diagnostics.length} diagnostic(s) on the class module:`);
    for (const diagnostic of fineClass.diagnostics) {
        console.log(`     [${diagnostic.severity}] ${diagnostic.code}: ${diagnostic.message}`);
    }

    check('Me, Friend, and events are legal in a class module', () =>
        assert.equal(fineClass.diagnostics.length, 0));

    const misfiledClass = await call('textDocument/diagnostics', {
        documentKey: 'Smoke/MisfiledClass',
        projectId: 'Smoke',
        generation: 1,
        source: CLASS_MODULE,
        moduleName: 'FineClass',
        moduleType: 'standard',
    });

    check('the same source as a standard module is the error list', () => {
        assert.ok(
            misfiledClass.diagnostics.some((diagnostic) => diagnostic.code === 'me-outside-object-module'),
            'expected me-outside-object-module for the standard kind');
    });

    // Completions: members after a dot against the host model, and identifiers elsewhere.
    const memberSource = GOOD_MODULE.replace('    n = 1', '    ThisWorkbook.');
    const memberOffset = memberSource.indexOf('ThisWorkbook.') + 'ThisWorkbook.'.length;
    const memberAnswer = await call('textDocument/completion', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: memberSource,
        offset: memberOffset,
        moduleType: 'standard',
    });

    console.log(`  -> ${memberAnswer.items.length} member completion(s) after ThisWorkbook.`);
    check('ThisWorkbook. offers workbook members', () => {
        assert.ok(memberAnswer.items.length > 10, 'expected a member surface');
        assert.ok(memberAnswer.items.some((item) => item.label === 'Worksheets'),
            'expected Worksheets among the members');
    });

    const identifierSource = GOOD_MODULE.replace('    n = 1', '    n = ');
    const identifierOffset = identifierSource.indexOf('    n = ') + '    n = '.length;
    const identifierAnswer = await call('textDocument/completion', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: identifierSource,
        offset: identifierOffset,
        moduleType: 'standard',
    });

    console.log(`  -> ${identifierAnswer.items.length} completion(s) at an expression position`);
    check('an expression position offers globals and the local', () => {
        assert.ok(identifierAnswer.items.some((item) => item.label === 'Application'),
            'expected the Application global');
        assert.ok(identifierAnswer.items.some((item) => item.label === 'n'),
            'expected the local variable n');
    });

    check('a procedure from the other module is offered', () => {
        assert.ok(identifierAnswer.items.some((item) => item.label === 'Probe'),
            'expected Probe from BadModule');
    });

    // A class module's own name is a receiver: factory-style classes are addressed by name.
    const classReceiverSource = GOOD_MODULE.replace('    n = 1', '    FineClass.');
    const classReceiverOffset = classReceiverSource.indexOf('FineClass.') + 'FineClass.'.length;
    const classReceiver = await call('textDocument/completion', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: classReceiverSource,
        offset: classReceiverOffset,
        moduleType: 'standard',
    });

    console.log(`  -> ${classReceiver.items.length} member(s) after FineClass.`);
    check('a class name as receiver offers the class members', () => {
        assert.ok(classReceiver.items.some((item) => item.label === 'Describe'),
            'expected Describe among the members');
        assert.ok(classReceiver.items.some((item) => item.label === 'Name'),
            'expected the Name property among the members');
    });

    // didChange: the engine holds the live text, and a request carries an offset and nothing
    // else. The notification has no id and gets no answer; the pipe's order is the contract.
    const liveSource = GOOD_MODULE.replace('    n = 1', '    ThisWorkbook.');
    const liveDot = liveSource.indexOf('ThisWorkbook.') + 'ThisWorkbook.'.length;
    socket.write(`${JSON.stringify({
        jsonrpc: '2.0',
        method: 'textDocument/didChange',
        params: { projectId: 'Smoke', moduleName: 'GoodModule', source: liveSource },
    })}\n`);

    const offsetOnly = await call('textDocument/completion', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        offset: liveDot,
        moduleType: 'standard',
    });

    check('an offset-only completion answers from the live text', () => {
        assert.ok(offsetOnly.items.some((item) => item.label === 'Worksheets'),
            'expected Worksheets from the live dot');
    });

    socket.write(`${JSON.stringify({
        jsonrpc: '2.0',
        method: 'textDocument/didChange',
        params: {
            projectId: 'Smoke',
            moduleName: 'GoodModule',
            edits: [{ start: liveDot - 1, end: liveDot, text: '' }],
        },
    })}\n`);

    const afterEdit = await call('textDocument/completion', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        offset: liveDot - 1,
        moduleType: 'standard',
    });

    check('an edit notification moves the live text', () => {
        // The edit deleted the dot: an answer still offering members would prove the engine
        // held the old text.
        assert.ok(!afterEdit.items.some((item) => item.label === 'Worksheets'),
            'the dot is gone, so members must be too');
    });

    // Hovers: the identifier under the cursor, described from the same project facts.
    const localHover = await call('textDocument/hover', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: GOOD_MODULE,
        offset: GOOD_MODULE.indexOf('n = 1'),
        moduleType: 'standard',
    });

    console.log(`  -> hover on the local: ${localHover.hover?.signature ?? '(none)'}`);
    check('hovering the local names its declaration', () => {
        assert.ok(localHover.hover, 'expected a hover');
        assert.ok(localHover.hover.signature.includes('n As Long'),
            `signature was '${localHover.hover.signature}'`);
    });

    const globalHoverSource = GOOD_MODULE.replace('    n = 1', '    ThisWorkbook.Save');
    const globalHover = await call('textDocument/hover', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: globalHoverSource,
        offset: globalHoverSource.indexOf('ThisWorkbook') + 3,
        moduleType: 'standard',
    });

    console.log(`  -> hover on the global: ${globalHover.hover?.signature ?? '(none)'}`);
    check('hovering ThisWorkbook names the host global', () => {
        assert.ok(globalHover.hover, 'expected a hover');
        assert.ok(globalHover.hover.signature.includes('ThisWorkbook As Workbook'),
            `signature was '${globalHover.hover.signature}'`);
    });

    const probeHoverSource = GOOD_MODULE.replace('    n = 1', '    Probe');
    const probeHover = await call('textDocument/hover', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: probeHoverSource,
        offset: probeHoverSource.indexOf('    Probe') + 6,
        moduleType: 'standard',
    });

    console.log(`  -> hover across modules: ${probeHover.hover?.signature ?? '(none)'}`);
    check('hovering a procedure from the other module names its home', () => {
        assert.ok(probeHover.hover, 'expected a hover');
        assert.ok(probeHover.hover.signature.includes('Sub Probe'),
            `signature was '${probeHover.hover.signature}'`);
        assert.ok(probeHover.hover.details.some((detail) => detail.includes('BadModule')),
            `details were '${probeHover.hover.details.join('; ')}'`);
    });

    const blankHover = await call('textDocument/hover', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: GOOD_MODULE,
        offset: GOOD_MODULE.indexOf('()') + 1,
        moduleType: 'standard',
    });

    check('hovering nothing answers null', () => assert.equal(blankHover.hover, null));

    // Signature help: the call tip inside an argument list.
    const runtimeCallSource = GOOD_MODULE.replace('    n = 1', '    MsgBox(');
    const runtimeTip = await call('textDocument/signatureHelp', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: runtimeCallSource,
        offset: runtimeCallSource.indexOf('MsgBox(') + 'MsgBox('.length,
        moduleType: 'standard',
    });

    console.log(`  -> call tip: ${runtimeTip.signature?.label ?? '(none)'}`);
    check('a runtime call shows its tip with the first parameter active', () => {
        assert.ok(runtimeTip.signature, 'expected a signature');
        assert.ok(runtimeTip.signature.label.includes('MsgBox'),
            `label was '${runtimeTip.signature.label}'`);
        assert.ok(runtimeTip.signature.parameters.length > 0, 'expected parameters');
        assert.equal(runtimeTip.signature.activeParameter, 0);
    });

    const secondArgSource = GOOD_MODULE.replace('    n = 1', '    MsgBox "hi", ');
    const secondArgTip = await call('textDocument/signatureHelp', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: secondArgSource,
        offset: secondArgSource.indexOf(', ') + 2,
        moduleType: 'standard',
    });

    check('a parenless call advances the active parameter past the comma', () => {
        assert.ok(secondArgTip.signature, 'expected a signature');
        assert.equal(secondArgTip.signature.activeParameter, 1);
    });

    const noCallTip = await call('textDocument/signatureHelp', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: GOOD_MODULE,
        offset: GOOD_MODULE.indexOf('n = 1'),
        moduleType: 'standard',
    });

    check('outside any call there is no tip', () => assert.equal(noCallTip.signature, null));

    // Smart Enter: the block completion, header parens, and continuations the extension does.
    const applyEdits = (text, edits) => [...edits]
        .sort((a, b) => b.start - a.start)
        .reduce((out, edit) => out.slice(0, edit.start) + edit.text + out.slice(edit.end), text);

    const subEnterSource = 'Option Explicit\r\n\r\nSub Test\r\n';
    const subEnter = await call('textDocument/smartEnter', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: subEnterSource,
        offset: subEnterSource.indexOf('Sub Test') + 'Sub Test'.length,
        moduleType: 'standard',
    });

    console.log(`  -> smart enter on Sub Test: ${subEnter.edits.length} edit(s), caret ${subEnter.caret}`);
    check('Enter after Sub Test completes the parens and inserts End Sub', () => {
        const after = applyEdits(subEnterSource, subEnter.edits);
        assert.equal(after, 'Option Explicit\r\n\r\nSub Test()\r\n\r\n\t\r\n\r\nEnd Sub');
        assert.equal(after[subEnter.caret - 1], '\t', 'caret should sit at the end of the body line');
    });

    const withEnterSource = 'Sub Fine()\r\n    With Application\r\n\r\nEnd Sub\r\n';
    const withEnter = await call('textDocument/smartEnter', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: withEnterSource,
        offset: withEnterSource.indexOf('With Application') + 'With Application'.length,
        moduleType: 'standard',
    });

    check('Enter after With inserts End With and seeds the member dot', () => {
        const after = applyEdits(withEnterSource, withEnter.edits);
        assert.ok(after.includes('    End With'), `got: ${JSON.stringify(after)}`);
        assert.equal(after[withEnter.caret - 1], '.', 'caret should sit after the seeded dot');
    });

    const commentEnterSource = "' hello\r\n";
    const commentEnter = await call('textDocument/smartEnter', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: commentEnterSource,
        offset: commentEnterSource.indexOf("' hello") + "' hello".length,
        moduleType: 'standard',
    });

    check('Enter at the end of a comment continues the apostrophes', () => {
        assert.equal(commentEnter.edits.length, 1);
        assert.equal(commentEnter.edits[0].text, "' ");
    });

    const closedAheadSource = 'Sub Alpha\r\n\r\nEnd Sub\r\n';
    const closedAhead = await call('textDocument/smartEnter', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: closedAheadSource,
        offset: closedAheadSource.indexOf('Sub Alpha') + 'Sub Alpha'.length,
        moduleType: 'standard',
    });

    check('an already-closed block gets its parens and indent but no second closer', () => {
        const after = applyEdits(closedAheadSource, closedAhead.edits);
        assert.equal(after, 'Sub Alpha()\r\n\t\r\nEnd Sub\r\n');
    });

    const midlineSource = 'Sub Test\r\nleftover\r\n';
    const midline = await call('textDocument/smartEnter', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: midlineSource,
        offset: midlineSource.indexOf('Sub Test') + 'Sub Test'.length,
        moduleType: 'standard',
    });

    check('Enter that pushed text down owes nothing', () => assert.equal(midline.edits.length, 0));

    // Canonical case: keywords respelled, identifiers matched to their declarations.
    const lowerLineSource = GOOD_MODULE.replace('    n = 1', '    dim m as long');
    const lowerLineStart = lowerLineSource.indexOf('    dim m as long');
    const lowerLine = await call('textDocument/canonicalCase', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: lowerLineSource,
        start: lowerLineStart,
        end: lowerLineStart + '    dim m as long'.length,
        moduleType: 'standard',
    });

    console.log(`  -> canonical case on the lower-case line: ${lowerLine.edits.length} edit(s)`);
    check('dim m as long recases to Dim m As Long', () => {
        assert.deepEqual(lowerLine.edits.map((edit) => edit.text), ['Dim', 'As', 'Long']);
    });

    const runtimeCaseSource = GOOD_MODULE.replace('    n = 1', '    msgbox "hi"');
    const runtimeCaseEnd = runtimeCaseSource.indexOf('msgbox') + 'msgbox'.length;
    const runtimeCase = await call('textDocument/canonicalCase', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: runtimeCaseSource,
        start: runtimeCaseEnd,
        end: runtimeCaseEnd,
        single: true,
        moduleType: 'standard',
    });

    check('msgbox recases to MsgBox from the runtime surface', () => {
        assert.equal(runtimeCase.edits.length, 1);
        assert.equal(runtimeCase.edits[0].text, 'MsgBox');
    });

    const headerCaseSource = 'sub tester\r\n\r\nEnd Sub\r\n';
    const headerCase = await call('textDocument/canonicalCase', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: headerCaseSource,
        start: 0,
        end: 'sub tester'.length,
        completeHeader: true,
        moduleType: 'standard',
    });

    check('leaving a bare header recases it and completes the parens', () => {
        const after = applyEdits(headerCaseSource, headerCase.edits);
        assert.ok(after.startsWith('Sub tester()'), `got: ${JSON.stringify(after)}`);
    });

    const canonicalNoop = await call('textDocument/canonicalCase', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: GOOD_MODULE,
        start: GOOD_MODULE.indexOf('    Dim n As Long'),
        end: GOOD_MODULE.indexOf('    Dim n As Long') + '    Dim n As Long'.length,
        moduleType: 'standard',
    });

    check('an already-canonical line produces no edits', () => assert.equal(canonicalNoop.edits.length, 0));

    // Outline: the procedures under a module node, in declaration order, kinds spelled the way
    // the tree spells them.
    const seededOutline = await call('textDocument/outline', {
        projectId: 'Smoke',
        moduleName: 'FineClass',
    });

    console.log(`  -> outline of the seeded class: ${seededOutline.procedures.map((p) => `${p.kind} ${p.name}@${p.line}`).join(', ')}`);
    check('the seeded class outlines its procedures in order', () => {
        assert.deepEqual(
            seededOutline.procedures.map((p) => `${p.kind} ${p.name}`),
            ['Sub Adopt', 'Property Get Name', 'Sub Describe']);
        assert.ok(seededOutline.procedures.every((p) => p.line > 0), 'expected 1-based lines');
    });

    const liveOutline = await call('textDocument/outline', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: GOOD_MODULE.replace('Sub Fine()', 'Sub Renamed()'),
    });

    check('a live source outlines as it stands, not as it was seeded', () => {
        assert.deepEqual(liveOutline.procedures.map((p) => p.name), ['Renamed']);
        assert.equal(liveOutline.procedures[0].line, 3);
    });

    const unknownOutline = await call('textDocument/outline', {
        projectId: 'Smoke',
        moduleName: 'NoSuchModule',
    });

    check('an unknown module outlines empty', () => assert.equal(unknownOutline.procedures.length, 0));

    // Loop iterator sync: renaming the For side renames the Next side.
    const loopSource = 'Sub Fine()\r\n    For i = 1 To 3\r\n        n = 1\r\n    Next j\r\nEnd Sub\r\n';
    const loopSync = await call('textDocument/loopSync', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: loopSource,
        offset: loopSource.indexOf('For i') + 'For i'.length,
        moduleType: 'standard',
    });

    check('editing the For iterator renames its Next', () => {
        const after = applyEdits(loopSource, loopSync.edits);
        assert.ok(after.includes('Next i'), `got: ${JSON.stringify(after)}`);
    });

    // Quick fixes: resolved from the analyzer's own findings, which carry fix data the surface
    // never sees. Three mistakes, each with a different fix, in one module.
    const FIXABLE = [
        'Sub Broken()',
        '    Dim r As Range',
        '    r = ActiveSheet.Range("A1")',
        '    Call Helper 1, 2',
        'End Sub',
        '',
    ].join('\r\n');

    await call('textDocument/diagnostics', {
        documentKey: 'Smoke/Fixable',
        projectId: 'Smoke',
        generation: 1,
        source: FIXABLE,
        moduleName: 'BadModule',
        moduleType: 'standard',
    });

    const wholeModule = await call('textDocument/codeAction', {
        projectId: 'Smoke',
        moduleName: 'BadModule',
        source: FIXABLE,
        moduleType: 'standard',
        start: 0,
        end: FIXABLE.length,
    });

    console.log(`  -> ${wholeModule.actions.length} quick fix(es) over the module:`);
    for (const action of wholeModule.actions) {
        console.log(`     "${action.title}" [${action.code}]`);
    }

    check('a missing Set is offered one', () => {
        const fix = wholeModule.actions.find((action) => action.code === 'set-required');
        assert.ok(fix, 'expected a set-required fix');
        assert.ok(applyEdits(FIXABLE, fix.edits).includes('Set r = ActiveSheet'),
            'expected the fix to insert Set');
    });

    check('a Call without parentheses is offered them', () => {
        const fix = wholeModule.actions.find((action) =>
            action.code === 'call-requires-parens' && !action.title.startsWith('Suppress'));
        assert.ok(fix, 'expected a call-requires-parens fix');
        assert.ok(applyEdits(FIXABLE, fix.edits).includes('Call Helper(1, 2)'),
            `got: ${JSON.stringify(applyEdits(FIXABLE, fix.edits))}`);
    });

    check('a missing Option Explicit is offered one, and it is preferred', () => {
        const fix = wholeModule.actions.find((action) => action.title === 'Add Option Explicit');
        assert.ok(fix, 'expected an option-explicit-missing fix');
        assert.equal(fix.isPreferred, true);
    });

    check('every finding also offers to suppress itself', () => {
        assert.ok(wholeModule.actions.some((action) =>
            action.title === "Suppress 'set-required' on next line"));
    });

    // A caret rather than a selection: only the finding it sits in answers.
    const caret = FIXABLE.indexOf('r = ActiveSheet') + 1;
    const atCaret = await call('textDocument/codeAction', {
        projectId: 'Smoke',
        moduleName: 'BadModule',
        source: FIXABLE,
        moduleType: 'standard',
        start: caret,
        end: caret,
    });

    check('a caret answers only the finding it sits in', () => {
        assert.ok(atCaret.actions.length > 0, 'expected fixes at the caret');
        assert.ok(atCaret.actions.every((action) => action.code === 'set-required'),
            `got codes: ${atCaret.actions.map((action) => action.code).join(', ')}`);
    });

    // Text the diagnostics have never run against still gets fixes: a fix is an edit by offset,
    // so answering from findings made against older text would place it by arithmetic that has
    // stopped holding.
    const EDITED = FIXABLE.replace('Dim r As Range', 'Dim r As Range   ');
    const untouched = await call('textDocument/codeAction', {
        projectId: 'Smoke',
        moduleName: 'BadModule',
        source: EDITED,
        moduleType: 'standard',
        start: 0,
        end: EDITED.length,
    });

    check('never-analysed text is analysed rather than answered stale', () => {
        const fix = untouched.actions.find((action) => action.code === 'set-required');
        assert.ok(fix, 'expected a set-required fix against the edited text');
        assert.ok(applyEdits(EDITED, fix.edits).includes('Set r = ActiveSheet'),
            `got: ${JSON.stringify(applyEdits(EDITED, fix.edits))}`);
    });

    const clean = await call('textDocument/codeAction', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: GOOD_MODULE,
        moduleType: 'standard',
        start: 0,
        end: GOOD_MODULE.length,
    });

    check('code with nothing wrong offers no fixes', () => assert.equal(clean.actions.length, 0));

    // Semantic tokens: what a grammar cannot know. The class, the host global, and a local that
    // shadows a host global's name have to come out differently.
    const COLOURED = [
        'Option Explicit',
        '',
        'Sub Paint()',
        '    Dim maker As FineClass',
        '    Set maker = New FineClass',
        '    Application.Calculate',
        'End Sub',
        '',
    ].join('\r\n');

    const coloured = await call('textDocument/semanticTokens', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: COLOURED,
        moduleType: 'standard',
    });

    console.log(`  -> ${coloured.tokens.length} semantic token(s):`);
    for (const token of coloured.tokens) {
        console.log(`     ${token.type}${token.modifiers ? ` (${token.modifiers.join(',')})` : ''}` +
            ` ${JSON.stringify(COLOURED.slice(token.start, token.end))}`);
    }

    check("the project's own class is coloured as a class", () => {
        const hits = coloured.tokens.filter((token) =>
            COLOURED.slice(token.start, token.end) === 'FineClass');
        assert.equal(hits.length, 2, 'both mentions of FineClass');
        assert.ok(hits.every((token) => token.type === 'class'),
            `types were ${hits.map((token) => token.type).join(', ')}`);
    });

    check('a host global is marked as one', () => {
        const hit = coloured.tokens.find((token) =>
            COLOURED.slice(token.start, token.end) === 'Application');
        assert.ok(hit, 'expected Application');
        assert.equal(hit.type, 'variable');
        assert.deepEqual(hit.modifiers, ['defaultLibrary']);
    });

    check('tokens arrive in position order, which the surface encodes as deltas', () => {
        for (let i = 1; i < coloured.tokens.length; i++) {
            assert.ok(coloured.tokens[i].start >= coloured.tokens[i - 1].start,
                `token ${i} starts before token ${i - 1}`);
        }
    });

    // A local declared with a host global's name is the developer's, not Excel's.
    const SHADOWED = COLOURED.replace('    Application.Calculate', '    Dim Application As Long');
    const shadowed = await call('textDocument/semanticTokens', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: SHADOWED,
        moduleType: 'standard',
    });

    check('a declaration of the same name stops it being a host global', () => {
        assert.ok(!shadowed.tokens.some((token) =>
            SHADOWED.slice(token.start, token.end) === 'Application'
            && token.modifiers?.includes('defaultLibrary')),
            'a shadowed name must not be marked as the host library');
    });

    // Navigation: where a name is declared, and everywhere in THIS workbook it is used. The
    // workbook is the boundary — a second one holding the same names must not appear.
    const NAV_CALLER = [
        'Option Explicit',
        '',
        'Sub Drive()',
        '    Dim w As FineClass',
        '    Set w = New FineClass',
        '    w.Describe',
        '    Probe',
        'End Sub',
        '',
    ].join('\r\n');

    await call('project/open', {
        projectId: 'Elsewhere',
        generation: 1,
        modules: [{ moduleName: 'BadModule', source: BAD_MODULE, type: 'standard' }],
    });

    const memberDefinition = await call('textDocument/definition', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: NAV_CALLER,
        moduleType: 'standard',
        offset: NAV_CALLER.indexOf('w.Describe') + 4,
    });

    check('a class member reached through a receiver resolves to its own module', () => {
        assert.equal(memberDefinition.locations.length, 1);
        assert.equal(memberDefinition.locations[0].module, 'FineClass');
    });

    const typeDefinition = await call('textDocument/definition', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: NAV_CALLER,
        moduleType: 'standard',
        offset: NAV_CALLER.indexOf('As FineClass') + 4,
    });

    check('a type name resolves to the module that is the type', () => {
        assert.equal(typeDefinition.locations.length, 1);
        assert.equal(typeDefinition.locations[0].module, 'FineClass');
    });

    // GoodModule's caller exists only in unsaved text, which is the case that matters: an answer
    // computed from the seeded copy would miss the call and, once rename is built on it, would
    // leave that call behind.
    socket.write(`${JSON.stringify({
        jsonrpc: '2.0',
        method: 'textDocument/didChange',
        params: { projectId: 'Smoke', moduleName: 'GoodModule', source: NAV_CALLER },
    })}\n`);

    const uses = await call('textDocument/references', {
        projectId: 'Smoke',
        moduleName: 'BadModule',
        source: BAD_MODULE,
        moduleType: 'standard',
        offset: BAD_MODULE.indexOf('Sub Probe') + 5,
        includeDeclaration: true,
    });

    console.log(`  -> ${uses.locations.length} use(s) of Probe: ` +
        uses.locations.map((where) => `${where.module}:${where.line}:${where.column}`).join(', '));

    check('a procedure is found where it is declared and where it is called', () => {
        const modules = new Set(uses.locations.map((where) => where.module));
        assert.ok(modules.has('BadModule'), 'the declaration');
        assert.ok(modules.has('GoodModule'), 'the call from the other module, in unsaved text');
    });

    check('the other open workbook is not searched, though it holds the same module name', () => {
        // Elsewhere/BadModule declares Probe too. An answer naming it would be a rename waiting
        // to edit a workbook nobody asked about.
        assert.equal(
            uses.locations.filter((where) => where.module === 'BadModule').length,
            1,
            'exactly one BadModule, this project\'s');
    });

    const withoutDeclaration = await call('textDocument/references', {
        projectId: 'Smoke',
        moduleName: 'BadModule',
        source: BAD_MODULE,
        moduleType: 'standard',
        offset: BAD_MODULE.indexOf('Sub Probe') + 5,
        includeDeclaration: false,
    });

    check('excluding the declaration drops exactly it', () =>
        assert.equal(withoutDeclaration.locations.length, uses.locations.length - 1));

    const nowhere = await call('textDocument/definition', {
        projectId: 'Smoke',
        moduleName: 'GoodModule',
        source: NAV_CALLER,
        moduleType: 'standard',
        offset: NAV_CALLER.indexOf('Option') + 2,
    });

    check('a keyword is not a symbol', () => assert.equal(nowhere.locations.length, 0));

    // Rename: the same set from every entry point.
    //
    // Asked at a qualified call, the member resolver answers with EVERY module declaring that
    // name, and a rename built straight on that went through all of them — renaming a procedure
    // in another module that merely shared a name (found live, 2026-08-06). Anchoring at the one
    // declaration is what fixed it, so the test that matters is that starting from a call site
    // and starting from the declaration reach exactly the same modules.
    const SHARED_A = [
        'Option Explicit',
        '',
        'Public Sub RunTotal()',
        '    Debug.Print "a"',
        'End Sub',
        '',
    ].join('\r\n');

    const SHARED_B = [
        'Option Explicit',
        '',
        'Public Sub RunTotal()',
        '    Debug.Print "b"',
        'End Sub',
        '',
    ].join('\r\n');

    const NAMES_A = [
        'Option Explicit',
        '',
        'Public Sub Test()',
        '    SharedA.RunTotal',
        'End Sub',
        '',
    ].join('\r\n');

    await call('project/open', {
        projectId: 'Shared',
        generation: 1,
        modules: [
            { moduleName: 'SharedA', source: SHARED_A, type: 'standard' },
            { moduleName: 'SharedB', source: SHARED_B, type: 'standard' },
            { moduleName: 'Names', source: NAMES_A, type: 'standard' },
        ],
    });

    const fromCall = await call('textDocument/rename', {
        projectId: 'Shared',
        moduleName: 'Names',
        source: NAMES_A,
        moduleType: 'standard',
        offset: NAMES_A.indexOf('SharedA.RunTotal') + 'SharedA.Run'.length,
        newName: 'RunTotalTest',
    });

    const fromDeclaration = await call('textDocument/rename', {
        projectId: 'Shared',
        moduleName: 'SharedA',
        source: SHARED_A,
        moduleType: 'standard',
        offset: SHARED_A.indexOf('Sub RunTotal') + 6,
        newName: 'RunTotalTest',
    });

    const touched = (answer) => answer.modules.map((entry) => entry.module).sort().join(', ');
    console.log(`  -> from the call site: ${touched(fromCall)}`);
    console.log(`  -> from the declaration: ${touched(fromDeclaration)}`);

    check('a rename from a call site reaches the same modules as one from the declaration', () =>
        assert.equal(touched(fromCall), touched(fromDeclaration)));

    check('a module that merely shares the name is not renamed', () => {
        assert.ok(!fromCall.modules.some((entry) => entry.module === 'SharedB'),
            'SharedB declares its own RunTotal and must be left alone');
        assert.equal(touched(fromCall), 'Names, SharedA');
    });

    // Renaming a MODULE. Its name is not in its own text, so this is about what every other
    // module says: a qualifier reaching into it, and — for a class — a type naming it.
    const MOD_TARGET = ['Option Explicit', '', 'Public Sub Recalc()', 'End Sub', ''].join('\r\n');
    const MOD_CALLER = [
        'Option Explicit',
        '',
        'Sub Drive()',
        '    Target.Recalc',
        '    Dim k As Kind',
        '    Set k = New Kind',
        '    Debug.Print "Target is data"',
        '    TargetExtra.Thing',
        'End Sub',
        '',
    ].join('\r\n');

    await call('project/open', {
        projectId: 'Modules',
        generation: 1,
        modules: [
            { moduleName: 'Target', source: MOD_TARGET, type: 'standard' },
            { moduleName: 'Kind', source: 'Option Explicit\r\n', type: 'class' },
            { moduleName: 'TargetExtra', source: 'Option Explicit\r\n', type: 'standard' },
            { moduleName: 'Caller', source: MOD_CALLER, type: 'standard' },
        ],
    });

    const renamedModule = await call('workspace/renameModule', {
        projectId: 'Modules',
        moduleName: 'Target',
        newName: 'Utils',
    });

    const callerAfter = renamedModule.modules.find((entry) => entry.module === 'Caller')?.source ?? '';
    console.log(`  -> renaming a module rewrote ${renamedModule.modules.length} other module(s)`);

    check('a qualified call into the module follows it', () =>
        assert.ok(callerAfter.includes('Utils.Recalc'), callerAfter));

    check('the name inside a string literal is left alone', () =>
        assert.ok(callerAfter.includes('"Target is data"'),
            'a module name in a string is data, not a reference'));

    check('another module that merely starts with the name is left alone', () =>
        assert.ok(callerAfter.includes('TargetExtra.Thing'), callerAfter));

    const renamedClass = await call('workspace/renameModule', {
        projectId: 'Modules',
        moduleName: 'Kind',
        newName: 'Sort',
    });

    check('a class module is followed by As and New as well', () => {
        const after = renamedClass.modules.find((entry) => entry.module === 'Caller')?.source ?? '';
        assert.ok(after.includes('As Sort'), after);
        assert.ok(after.includes('New Sort'), after);
    });

    for (const [what, name, expected] of [
        ['a name already taken', 'Kind', 'already has a module'],
        ['a name that is not an identifier', '9Bad', 'is not a VBA name'],
        ['a keyword', 'Next', 'is a VBA keyword'],
        ['the name it already has', 'Target', 'already its name'],
    ]) {
        const refused = await call('workspace/renameModule', {
            projectId: 'Modules',
            moduleName: 'Target',
            newName: name,
        });

        check(`renaming to ${what} is refused`, () => {
            assert.equal(refused.modules.length, 0);
            assert.ok(refused.refused?.includes(expected), refused.refused);
        });
    }

    const noSuchModule = await call('workspace/renameModule', {
        projectId: 'Modules',
        moduleName: 'NeverSeeded',
        newName: 'Fine',
    });

    // The symbol assembly invents a module it has not been seeded with, so that one opened before
    // the first seed still answers. A rename must not ride that and quietly rename nothing.
    check('renaming a module the workbook does not have is refused', () =>
        assert.ok(noSuchModule.refused?.includes('not a module'), noSuchModule.refused));

    // Both entry points reach the same operation. The explorer asks by name; renaming the
    // qualifier in code asks by offset, and a module's name is not a symbol, so that path used to
    // refuse. Pinned as a property because the equivalent gap in symbol rename was only found
    // once a developer used the entry point no probe had tried (2026-08-06).
    const fromCode = await call('textDocument/rename', {
        projectId: 'Modules',
        moduleName: 'Caller',
        source: MOD_CALLER,
        moduleType: 'standard',
        offset: MOD_CALLER.indexOf('Target.Recalc') + 3,
        newName: 'Utils',
    });

    check('renaming a module from a qualifier in code is a module rename', () => {
        assert.equal(fromCode.module, 'Target', 'the engine must say which component to rename');
        assert.equal(fromCode.modules.length, renamedModule.modules.length);
        assert.deepEqual(
            fromCode.modules.map((entry) => entry.module).sort(),
            renamedModule.modules.map((entry) => entry.module).sort(),
            'the explorer and the code entry must reach the same modules');
    });

    // Renaming an INTERFACE, and the classes that implement it. VBA names an implemented member
    // `Interface_Member`, so that prefix is part of the contract: rename the interface and leave
    // the prefix and the class stops implementing anything, which the compiler notices.
    const SHAPE = ['Option Explicit', '', 'Public Sub Draw()', 'End Sub', ''].join('\r\n');
    const CIRCLE = [
        'Option Explicit',
        '',
        'Implements IShape',
        '',
        'Private Sub IShape_Draw()',
        'End Sub',
        '',
    ].join('\r\n');
    const SHAPE_USER = [
        'Option Explicit',
        '',
        'Sub Go()',
        '    Dim s As IShape',
        '    Set s = New Circle',
        '    IShapeCounter = 1',
        'End Sub',
        '',
    ].join('\r\n');

    await call('project/open', {
        projectId: 'Shapes',
        generation: 1,
        modules: [
            { moduleName: 'IShape', source: SHAPE, type: 'class' },
            { moduleName: 'Circle', source: CIRCLE, type: 'class' },
            { moduleName: 'ShapeUser', source: SHAPE_USER, type: 'standard' },
        ],
    });

    const renamedInterface = await call('workspace/renameModule', {
        projectId: 'Shapes',
        moduleName: 'IShape',
        newName: 'IDrawable',
    });

    const circleAfter = renamedInterface.modules.find((entry) => entry.module === 'Circle')?.source ?? '';
    const userAfter = renamedInterface.modules.find((entry) => entry.module === 'ShapeUser')?.source ?? '';

    check('the Implements statement follows the interface', () =>
        assert.ok(circleAfter.includes('Implements IDrawable'), circleAfter));

    check('an implemented member keeps its interface prefix', () =>
        assert.ok(circleAfter.includes('Private Sub IDrawable_Draw()'),
            'the Interface_Member prefix is the contract, not a coincidence'));

    check('a declared variable of the interface type follows it', () =>
        assert.ok(userAfter.includes('As IDrawable'), userAfter));

    check('an unrelated name that merely starts with the interface is left alone', () =>
        assert.ok(userAfter.includes('IShapeCounter = 1'),
            'IShapeCounter is its own name in a module that implements nothing'));

    // Analysis against sources the engine was never given must be refused, not answered from
    // whatever it happens to hold.
    let refused = false;
    try {
        await call('textDocument/diagnostics', {
            documentKey: 'Other/Module1',
            projectId: 'NeverOpened',
            generation: 9,
            source: GOOD_MODULE,
            moduleName: 'Module1',
        });
    } catch (error) {
        refused = String(error.message).includes('-32000');
    }
    check('analysis of an unseeded project is refused', () => assert.ok(refused));

    await call('shutdown', {});
} finally {
    socket.end();
    engine.kill();
}

console.log(failures === 0 ? '\nSMOKE PASSED' : `\nSMOKE FAILED (${failures})`);
process.exit(failures === 0 ? 0 : 1);
