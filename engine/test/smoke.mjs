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
        ],
    });
    check('project/open accepts both modules', () => assert.equal(opened.modules, 2));

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
