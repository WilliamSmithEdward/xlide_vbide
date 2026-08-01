// Prints the reference lexer's token stream in the same shape the ported lexer prints, so the two
// can be compared over a corpus. This exists only for the port and leaves with it.

import { readFileSync } from 'node:fs';
import { basename } from 'node:path';
import { tokenize } from '../../../xlide_vscode/src/analyzer/lexer/tokenize';

const files = process.argv.slice(2);

if (files.length === 0) {
    process.stderr.write('Usage: node lexdump.cjs <file> [file...]\n');
    process.exit(2);
}

for (const file of files) {
    const source = readFileSync(file, 'utf8');
    const tokens = tokenize(source);

    process.stdout.write(`${JSON.stringify({ file: basename(file), tokens: tokens.length })}\n`);

    for (const token of tokens) {
        process.stdout.write(
            `${JSON.stringify({
                kind: token.kind,
                start: token.start,
                end: token.end,
                line: token.line,
                character: token.character,
                text: token.rawText,
                canonical: token.canonicalText ?? null,
            })}\n`,
        );
    }
}
