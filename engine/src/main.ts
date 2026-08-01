// Engine entry point. Serves one named pipe until the add-in disconnects or asks it to stop.

import net from 'node:net';
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

    let buffer = '';

    socket.on('data', (chunk) => {
        buffer += chunk.toString('utf8');

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

    try {
        const result = dispatcher.handle(request.method, request.params);

        // A message with no id is a notification and is deliberately not answered.
        if (request.id !== undefined) {
            respond(socket, { jsonrpc: '2.0', id: request.id, result });
        }
    } catch (error) {
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
