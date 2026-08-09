// Engine entry point. Serves one named pipe until the add-in disconnects or asks it to stop.

import net from 'node:net';
import { StringDecoder } from 'node:string_decoder';
import { Dispatcher, RpcError } from './dispatcher';
import { ErrorCode, type JsonRpcRequest, type JsonRpcResponse } from './protocol';

function pipePathFrom(argv: readonly string[]): string {
    const index = argv.indexOf('--pipe');
    const name = index >= 0 ? argv[index + 1] : undefined;

    if (!name) {
        throw new Error('Usage: xlide-engine --pipe <name>');
    }

    // Accept either a bare name or a full path, so a caller that already knows the convention and
    // one that does not both work.
    return name.startsWith('\\\\') ? name : `\\\\.\\pipe\\${name}`;
}

function respond(socket: net.Socket, response: JsonRpcResponse): void {
    socket.write(`${JSON.stringify(response)}\n`);
}

function serve(socket: net.Socket, onShutdown: () => void): void {
    const dispatcher = new Dispatcher();

    socket.setNoDelay(true);

    // A StringDecoder rather than chunk.toString('utf8'), because a socket splits its chunks on
    // byte boundaries and a multi-byte character straddling two of them decodes to U+FFFD in
    // each half: the character is destroyed, silently, and only for large payloads. Module
    // source is exactly the large payload here.
    //
    // Latent under the current client: the shim serialises with System.Text.Json, whose default
    // encoder escapes every non-ASCII character to a \uXXXX sequence, so the bytes arriving are
    // ASCII and nothing can straddle. Verified by round-tripping a 35KB module of accented text
    // through the pipe intact (2026-08-07). Fixed anyway: the correctness of this loop should
    // not rest on an escaping default in another language on the far side of the pipe.
    const decoder = new StringDecoder('utf8');
    let buffer = '';

    socket.on('data', (chunk) => {
        buffer += decoder.write(chunk);

        // One message per line. A partial line is kept until its newline arrives.
        let newline = buffer.indexOf('\n');
        while (newline >= 0) {
            const line = buffer.slice(0, newline).trim();
            buffer = buffer.slice(newline + 1);
            newline = buffer.indexOf('\n');

            if (line.length > 0) {
                handleLine(socket, dispatcher, line);
            }
        }

        if (dispatcher.shuttingDown) {
            socket.end();
            onShutdown();
        }
    });

    socket.on('error', () => {
        // A client that disappears is an ordinary end of session, not a fault.
        socket.destroy();
    });
}

function handleLine(socket: net.Socket, dispatcher: Dispatcher, line: string): void {
    let request: JsonRpcRequest;

    try {
        request = JSON.parse(line) as JsonRpcRequest;
    } catch {
        respond(socket, {
            jsonrpc: '2.0',
            id: 0,
            error: { code: ErrorCode.ParseError, message: 'The message was not valid JSON.' },
        });
        return;
    }

    if (typeof request.method !== 'string') {
        if (request.id !== undefined) {
            respond(socket, {
                jsonrpc: '2.0',
                id: request.id,
                error: { code: ErrorCode.InvalidRequest, message: 'The message named no method.' },
            });
        }
        return;
    }

    // One shape for a failure, whether it was thrown or rejected.
    const fail = (error: unknown): void => {
        if (request.id === undefined) {
            return;
        }

        respond(socket, {
            jsonrpc: '2.0',
            id: request.id,
            error:
                error instanceof RpcError
                    ? error.toJson()
                    : {
                          code: ErrorCode.InternalError,
                          message: error instanceof Error ? error.message : String(error),
                      },
        });
    };

    try {
        const result = dispatcher.handle(request.method, request.params);

        // A message with no id is a notification and is deliberately not answered.
        if (request.id === undefined) {
            return;
        }

        // AN ANSWER THAT IS NOT READY YET.
        //
        // Every method here used to answer immediately, so the result went straight out. The
        // import/export planner reads a folder, so it answers a promise — and a promise handed to
        // JSON.stringify is `{}`, which is a valid reply carrying nothing. The caller got an empty
        // object, decided the plan was unreadable and quietly used its own planner instead, and
        // every test still passed because both planners agree (2026-08-09).
        //
        // Checked for rather than making everything async: the synchronous methods are the hot
        // path, and an await on each would put a microtask between every keystroke and its answer.
        if (typeof (result as PromiseLike<unknown> | undefined)?.then === 'function') {
            // Held, because the narrowing above does not survive into the callback.
            const id = request.id;
            void Promise.resolve(result).then(
                (settled) => respond(socket, { jsonrpc: '2.0', id, result: settled }),
                fail);
            return;
        }

        respond(socket, { jsonrpc: '2.0', id: request.id, result });
    } catch (error) {
        fail(error);
    }
}

function main(): void {
    const path = pipePathFrom(process.argv);

    const server = net.createServer((socket) => {
        serve(socket, () => {
            server.close();
            // Any other connection keeps the process alive; unref lets it exit once idle.
            server.unref();
        });
    });

    server.on('error', (error) => {
        process.stderr.write(`engine: ${error.message}\n`);
        process.exit(1);
    });

    server.listen(path, () => {
        // The add-in waits for this line before connecting, so the pipe is guaranteed to exist by
        // the time it dials.
        process.stdout.write(`listening ${path}\n`);
    });
}

main();
