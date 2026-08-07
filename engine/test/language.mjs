// Non-ASCII text through the real engine over the real pipe, for every path that carries it.
//
// WHY THIS IS A SEPARATE TEST FROM THE COMPANION'S. xlide_vscode has a language matrix too
// (tests/vbaLanguageMatrix.test.ts), and it tests a different thing: that product reads and
// writes the .xlsm's VBA streams itself, so its risk is decoding module bytes against the
// project's PROJECTCODEPAGE. This product never touches those bytes. Text reaches it as a COM
// BSTR, which is already UTF-16, and its risk is what happens AFTER: JSON to the engine, JSON
// back, and — the sharp one — OFFSET ARITHMETIC.
//
// Every language feature here trades in offsets. Navigation, rename, diagnostics, hover,
// signature help and semantic tokens all name a position as a number of units into the source,
// and the two ends have to agree on what a unit is. JavaScript strings count UTF-16 code units;
// so do C# strings; but a byte count, a code-point count, or an Array.from() anywhere in between
// disagrees the moment the text stops being ASCII, and disagrees by MORE the further into the
// module you are. That is a defect that cannot happen in an English module and cannot be seen in
// a short one.
//
// So the matrix below is xlide_vscode's samples — deliberately the same list, so the two
// products are known to agree about the same languages — driven through THIS product's paths.
// Astral characters (emoji, which are surrogate pairs) get their own cases, because they are the
// one input where UTF-16 code units and code points genuinely differ.
//
//   node test/language.mjs
//   node test/language.mjs --exe

import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const dist = join(here, '..', 'dist');
const useExe = process.argv.includes('--exe');

const pipeName = `xlide-engine-language-${process.pid}`;
const command = useExe ? join(dist, 'xlide-engine.exe') : process.execPath;
const args = useExe ? ['--pipe', pipeName] : [join(dist, 'engine.cjs'), '--pipe', pipeName];

const engine = spawn(command, args, { stdio: ['ignore', 'pipe', 'pipe'] });
engine.stderr.on('data', (chunk) => process.stderr.write(`engine stderr: ${chunk}`));

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

let passed = 0;
const failures = [];

function check(what, body) {
    try {
        body();
        passed++;
        console.log(`ok   ${what}`);
    } catch (error) {
        failures.push(`${what}: ${error.message}`);
        console.log(`FAIL ${what}\n     ${error.message}`);
    }
}

/*
 * One native-language sample per script, taken from xlide_vscode's matrix so the two products
 * are tested against the same text. The code page numbers are its concern, not this one's, but
 * the labels are kept: when a row fails here and there, the same row names it in both places.
 */
const LANGUAGES = [
    ['Thai', 'ทดสอบภาษาไทย'],
    ['Japanese', 'テスト用モジュール'],
    ['Chinese, Simplified', '中文测试模块'],
    ['Korean', '한국어 테스트'],
    ['Chinese, Traditional', '繁體中文測試'],
    ['Central European', 'Příliš žluťoučký kůň Zażółć gęślą'],
    ['Cyrillic', 'Проверка русского текста'],
    ['Western European', 'déjà vu € œuvre Straße'],
    ['Greek', 'Δοκιμή ελληνικού κειμένου'],
    ['Turkish', 'Türkçe deneme ğüşiöç İı'],
    ['Hebrew', 'בדיקת עברית'],
    ['Arabic', 'اختبار العربية'],
    ['Baltic', 'Lietuviškas tekstas ąčęėįšųū'],
    ['Vietnamese', 'Tiếng Việt thử nghiệm'],
    ['Russian (KOI8)', 'Тест КОИ-8'],
    ['Ukrainian', 'Тест української ґї'],
    ['Astral (emoji)', '🎯 test 👨‍👩‍👧‍👦 done 🇯🇵'],
    ['Mixed', 'любой текст 中文 déjà ทดสอบ 🎯'],
];

/**
 * A module whose comments and string literals carry the sample, with a call site placed AFTER
 * all of it. The call site's offset is where the arithmetic either holds or does not: everything
 * non-ASCII sits between the start of the file and the position under test, so any unit
 * disagreement shows up as a resolution that lands on the wrong identifier or on nothing.
 */
function moduleWith(sample) {
    return [
        'Option Explicit',
        '',
        `' ${sample}`,
        `' ${sample} ${sample}`,
        '',
        'Public Sub Target()',
        `    Debug.Print "${sample}"`,
        'End Sub',
        '',
        'Public Sub Caller()',
        `    ' ${sample}`,
        '    Target',
        'End Sub',
        '',
    ].join('\r\n');
}

// The dispatcher refuses everything until this, so it comes before the matrix rather than
// inside it: eighteen identical "initialize has not been called" failures name nothing.
const hello = await call('initialize', {});
check('the engine answers the handshake', () => assert.equal(hello.engine, 'xlide'));

console.log(`\n== ${LANGUAGES.length} languages, each through open, diagnose, outline, navigate and rename ==\n`);

for (const [label, sample] of LANGUAGES) {
    const source = moduleWith(sample);
    const projectId = `Lang${label.replace(/[^A-Za-z]/g, '')}`;

    const opened = await call('project/open', {
        projectId,
        generation: 1,
        modules: [{ moduleName: 'Sample', source, type: 'standard' }],
    });

    // 1. The text survives the pipe. A module the engine could not read would fail everything
    //    below, but it would fail it in ways that read as analysis bugs rather than as encoding.
    check(`${label}: the project opens`, () => assert.equal(opened.modules, 1));

    // 2. Nothing in the text is mistaken for code. A comment carrying Arabic or an emoji must
    //    still be a comment, and a string literal must still be a string literal.
    const diagnostics = await call('textDocument/diagnostics', {
        documentKey: `${projectId}/Sample`,
        projectId,
        generation: 1,
        moduleName: 'Sample',
        moduleType: 'standard',
        source,
    });
    check(`${label}: clean text produces no findings`, () =>
        assert.deepEqual((diagnostics.diagnostics ?? []).map((d) => d.message), []));

    // 3. The outline finds both procedures, at their real lines.
    const outline = await call('textDocument/outline', {
        projectId,
        moduleName: 'Sample',
        moduleType: 'standard',
        source,
    });
    check(`${label}: the outline finds both procedures`, () =>
        assert.deepEqual((outline.procedures ?? []).map((p) => p.name).sort(), ['Caller', 'Target']));

    // 4. THE OFFSET TEST, and it is really a LINE AND COLUMN test.
    //
    //    `Target` is called on the last line, past every non-ASCII character in the file. The
    //    engine takes an offset in and answers a line and a column, so this exercises the whole
    //    conversion: an offset counted in anything but UTF-16 code units resolves the wrong
    //    identifier, and a column counted differently points at the wrong part of the right line.
    //    Both are invisible in English and both corrupt an edit.
    const callSite = source.lastIndexOf('    Target') + 4;
    const definition = await call('textDocument/definition', {
        projectId,
        moduleName: 'Sample',
        moduleType: 'standard',
        source,
        offset: callSite + 2,
    });

    const lines = source.split('\r\n');
    const wantLine = lines.findIndex((one) => one.startsWith('Public Sub Target')) + 1;
    const wantColumn = lines[wantLine - 1].indexOf('Target') + 1;

    check(`${label}: the declaration is found at the right line and column`, () => {
        const found = definition.locations ?? [];
        assert.equal(found.length, 1, `expected one location, got ${found.length}`);
        assert.equal(found[0].module, 'Sample');
        assert.equal(found[0].line, wantLine, 'the line is wrong');
        assert.equal(found[0].column, wantColumn,
            `the column is wrong: ${found[0].column} rather than ${wantColumn}. A column counted `
            + 'in bytes or code points rather than UTF-16 code units drifts by the width of the '
            + 'non-ASCII text to its left.');
        assert.equal(found[0].length, 'Target'.length);
    });

    // 5. RENAME PRODUCES WHOLE MODULE TEXTS, so the test is exact rather than approximate: the
    //    engine's answer must be the source with both occurrences of the name replaced and
    //    nothing else touched. Any offset drift shows up as a rewritten comment, and this
    //    compares the entire file, so it cannot be missed.
    const renamed = await call('textDocument/rename', {
        projectId,
        moduleName: 'Sample',
        moduleType: 'standard',
        source,
        offset: callSite + 2,
        newName: 'Renamed',
    });

    check(`${label}: rename rewrites the name and leaves the text alone`, () => {
        const produced = (renamed.modules ?? []).find((entry) => entry.module === 'Sample');
        assert.ok(produced, `rename touched no module: ${JSON.stringify(renamed.refused ?? renamed)}`);

        // Built by hand rather than with replaceAll, so the expectation is not computed the same
        // way the answer might be: only the two real occurrences change.
        const wanted = source
            .replace('Public Sub Target', 'Public Sub Renamed')
            .replace('\r\n    Target', '\r\n    Renamed');

        assert.equal(produced.source, wanted,
            'the produced text differs from the source with only the name changed');
        assert.equal(produced.replaced, 2, 'both occurrences should have been replaced');
    });
}

/*
 * The sharp case on its own, spelled out.
 *
 * An emoji is one code POINT and two UTF-16 code units, and a family emoji is several of each
 * joined by zero-width joiners. A module with one in a comment shifts every offset after it, and
 * by a different amount depending on which unit you count in. This is the case that separates
 * "handles accents" from "handles text".
 */
const ASTRAL = [
    'Option Explicit',
    '',
    "' 👨‍👩‍👧‍👦 🇯🇵 🎯 a comment of surrogate pairs",
    '',
    'Public Sub Anchor()',
    'End Sub',
    '',
    'Public Sub Uses()',
    '    Anchor',
    'End Sub',
    '',
].join('\r\n');

await call('project/open', {
    projectId: 'Astral',
    generation: 1,
    modules: [{ moduleName: 'Sample', source: ASTRAL, type: 'standard' }],
});

const astralNav = await call('textDocument/definition', {
    projectId: 'Astral',
    moduleName: 'Sample',
    moduleType: 'standard',
    source: ASTRAL,
    offset: ASTRAL.lastIndexOf('    Anchor') + 6,
});

check('surrogate pairs do not shift a navigation target', () => {
    const found = astralNav.locations ?? [];
    assert.equal(found.length, 1, 'navigation found nothing');
    const lines = ASTRAL.split('\r\n');
    const wantLine = lines.findIndex((one) => one.startsWith('Public Sub Anchor')) + 1;
    assert.equal(found[0].line, wantLine);
    assert.equal(found[0].column, lines[wantLine - 1].indexOf('Anchor') + 1);
});

check('the code-unit count is what both ends mean by an offset', () => {
    // Stated as an assertion rather than a comment, because it is the assumption everything
    // above rests on: were the engine counting code points, this length would differ.
    assert.notEqual(ASTRAL.length, [...ASTRAL].length,
        'the astral sample must actually contain surrogate pairs, or it tests nothing');
});

console.log(`\n${passed} passed, ${failures.length} failed`);
for (const failure of failures) { console.log(`  ! ${failure}`); }

socket.end();
engine.kill();
process.exit(failures.length === 0 ? 0 : 1);
