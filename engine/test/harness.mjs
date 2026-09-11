// The shared half of every engine test: spawn the engine, wait for its own "listening" line,
// connect the named pipe, frame newline-delimited JSON-RPC, and report checks in the one
// verdict shape the gate parses.
//
// Four suites carried this scaffold at 0.85 to 0.997 similarity, and the copies had already
// drifted where it costs: smoke and language accepted --exe so they could run against the
// packaged executable, while positions and freshness hardcoded the bundle and could not be
// pointed at what the add-in actually launches. One copy, and --exe works everywhere.

import { spawn } from 'node:child_process';
import net from 'node:net';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const dist = join(here, '..', 'dist');

/**
 * Spawns one engine and connects to it. The label names the pipe, so two suites running at
 * once cannot collide; `--exe` on the command line (or useExe) runs the packaged executable
 * instead of the bundle. Answers { call, stop }: call(method, params) is one JSON-RPC round
 * trip with a 30s budget, stop() ends the pipe and the process.
 */
export async function startEngine(label, { useExe = process.argv.includes('--exe') } = {}) {
    const pipeName = `xlide-engine-${label}-${process.pid}`;
    const command = useExe ? join(dist, 'xlide-engine.exe') : process.execPath;
    const args = useExe ? ['--pipe', pipeName] : [join(dist, 'engine.cjs'), '--pipe', pipeName];

    console.log(`starting ${useExe ? 'executable' : 'bundle'}: ${command}`);

    const engine = spawn(command, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    engine.stderr.on('data', (chunk) => process.stderr.write(`engine stderr: ${chunk}`));

    // Resolves once the engine reports the pipe is open, so there is no race and no sleep.
    //
    // EVERY BUDGET TIMER IS CLEARED WHEN ITS ANSWER ARRIVES, here and in call() below. A pending
    // timer holds node open, so a suite that ends by running out of work - rather than by
    // process.exit - sat out the last 30s budget it had started: form-unvouched and
    // class-predeclared took 30 seconds each for 3 and 6 checks, a minute of every gate run and
    // the same tail on pass-cost in CI (2026-09-10).
    await new Promise((resolve, reject) => {
        let seen = '';
        const budget = setTimeout(() => reject(new Error('engine did not report listening within 30s')), 30_000);
        engine.stdout.on('data', (chunk) => {
            seen += chunk.toString();
            if (seen.includes('listening')) {
                clearTimeout(budget);
                resolve();
            }
        });
        engine.on('exit', (code) => {
            clearTimeout(budget);
            reject(new Error(`engine exited early with code ${code}`));
        });
    });

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
                clearTimeout(waiter.budget);
                if (message.error) { waiter.reject(new Error(`${message.error.code}: ${message.error.message}`)); }
                else { waiter.resolve(message.result); }
            }
        }
    });

    function call(method, params) {
        const id = nextId++;
        return new Promise((resolve, reject) => {
            // Cleared by the answer (above), so a finished suite is not held open by it.
            const budget = setTimeout(() => {
                if (pending.delete(id)) { reject(new Error(`${method} timed out`)); }
            }, 30_000);
            pending.set(id, { resolve, reject, budget });
            socket.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
        });
    }

    /** A notification: no id, no answer, the pipe's order is the contract (didChange). */
    function notify(method, params) {
        socket.write(`${JSON.stringify({ jsonrpc: '2.0', method, params })}\n`);
    }

    function stop() {
        socket.end();
        engine.kill();
    }

    return { call, notify, stop };
}

/**
 * The reporter: check(what, body) runs one assertion body, and done() prints the
 * "N passed, M failed" verdict line the gate parses and answers the process exit code.
 */
export function reporter() {
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

    function done() {
        console.log(`\n${passed} passed, ${failures.length} failed`);
        for (const failure of failures) {
            console.log(`  ${failure}`);
        }
        return failures.length === 0 ? 0 : 1;
    }

    return { check, done };
}
