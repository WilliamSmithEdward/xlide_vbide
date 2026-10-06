// Turn the host's type-library snapshot into the analyzer's ordinary object-model shape.
// It is registered under a content-addressed token: hostRegistry caches merged models by token,
// so replacing a library under the same GUID must never reuse an older merged answer.

import { createHash } from 'node:crypto';
import type { HostConstant, HostMember, HostObjectModel, HostType } from '../../../xlide_vscode/src/analyzer/host/excelObjectModel';
import { registerHostObjectModel, type VbaHostToken } from '../../../xlide_vscode/src/analyzer/host/hostRegistry';
import type { ReferenceLibraryPayload } from './protocol';

const registered = new Set<string>();

export function registerReferenceLibraries(
    libraries: readonly ReferenceLibraryPayload[] | undefined,
    referencedGuids: readonly string[] | undefined,
): ReadonlyMap<string, string> {
    const allowed = new Set((referencedGuids ?? []).map((guid) => guid.toUpperCase()));
    const tokens = new Map<string, string>();
    for (const library of libraries ?? []) {
        const guid = library.guid.toUpperCase();
        if (!allowed.has(guid) || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(library.name)
            || library.types.length > 512) {
            continue;
        }

        const digest = createHash('sha256').update(JSON.stringify(library)).digest('hex').slice(0, 16);
        // The upstream registry normalizes lookup tokens to lower case. Keep registration in
        // that same spelling so custom models survive registry normalization (v11.0.1).
        const token = `typelib:${guid.toLowerCase()}:${digest}`;
        if (!registered.has(token)) {
            const model = modelFor(library);
            registerHostObjectModel(token as VbaHostToken, () => model);
            registered.add(token);
        }
        tokens.set(guid, token);
    }
    return tokens;
}

function modelFor(library: ReferenceLibraryPayload): HostObjectModel {
    const types: Record<string, HostType> = {};
    const aliases: Record<string, string> = {};
    const constants: Record<string, HostConstant> = {};
    const enums: NonNullable<HostObjectModel['enums']> = {};
    const knownTypes = new Set(library.types.map((type) => type.name.toLowerCase()));

    for (const type of library.types) {
        if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(type.name)) {
            continue;
        }
        const qualified = `${library.name}.${type.name}`;
        if (type.kind === 'enum') {
            enums[type.name] = { displayName: type.name, library: library.name };
            for (const member of type.members) {
                if (member.kind !== 'Const' || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(member.name)) {
                    continue;
                }
                const value = /^Const\s+\S+\s*=\s*(-?\d+)$/i.exec(member.signature)?.[1];
                constants[member.name] = {
                    name: member.name,
                    type: type.name,
                    ...(value === undefined ? {} : { value: Number(value) }),
                    source: 'external',
                };
            }
            continue;
        }

        aliases[type.name.toLowerCase()] = qualified;
        aliases[qualified.toLowerCase()] = qualified;
        const members: HostMember[] = type.members
            .filter((member) => /^[A-Za-z_][A-Za-z0-9_]*$/.test(member.name))
            .map((member) => {
                const declaredType = /\bAs\s+([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?)(?:\(\))?\s*$/i
                    .exec(member.signature)?.[1];
                const bare = declaredType?.split('.').at(-1);
                const returns = bare && knownTypes.has(bare.toLowerCase())
                    ? `${library.name}.${bare}` : undefined;
                return {
                    name: member.name,
                    kind: member.kind === 'Sub' || member.kind === 'Function' ? 'method' : 'property',
                    signature: member.signature,
                    ...(declaredType ? { declaredType } : {}),
                    ...(returns ? { returns } : {}),
                };
            });
        types[qualified] = { displayName: type.name, members, exhaustive: false };
    }

    return {
        source: `referenced type library ${library.name} (${library.guid})`,
        hostName: library.name,
        types,
        aliases,
        globals: {},
        constants,
        enums,
    };
}
