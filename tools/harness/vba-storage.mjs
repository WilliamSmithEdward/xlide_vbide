/*
 * A workbook's VBA storage, read from the file Excel has ALREADY saved - no export, no COM, no
 * Excel involved at all.
 *
 * This is the instrument for #53 (docs/userform-designer.md: every changed property rides the
 * document). A form's SITED baseline - what a control holds when it sits on a form, as against a
 * bare coclass instance - is exactly what MSForms writes into its own storage, because the format
 * records only properties that differ from the default. Getting at those bytes was stage one, and
 * this is it: a compound file reader small enough to port to C# when the product needs it.
 *
 * The road, proven 2026-08-16 on FormFixture.xlsm:
 *
 *   the .xlsm is a ZIP        ->  xl/vbaProject.bin
 *   vbaProject.bin is a CFB   ->  /<FormName>/f  (the form and its site array)
 *                                 /<FormName>/o  (each control's own property block)
 *                                 /<FormName>/i06, i10/...  (storages for controls that need one)
 *
 * Usage, from a probe:
 *
 *   import { readCfb } from "./vba-storage.mjs";
 *   const cfb = readCfb("vbaProject.bin");
 *   for (const [path, entry] of cfb.paths) console.log(path, entry.size);
 *   const f = cfb.readStream(cfb.paths.get("/EntryForm/f"));
 *
 * This module is the compound file and nothing above it. What the streams MEAN - FormControl,
 * the site array, the per-control PropMasks - is saved-design.mjs, against [MS-OFORMS].
 */
import { readFileSync } from "node:fs";

/** Takes a path, or the bytes themselves when the caller already has them out of a ZIP. */
export function readCfb(pathOrBytes) {
  const bytes = Buffer.isBuffer(pathOrBytes) ? pathOrBytes : readFileSync(pathOrBytes);
  const sig = bytes.readBigUInt64LE(0);
  if (sig !== 0xE11AB1A1E011CFD0n) throw new Error(`not a compound file: ${sig.toString(16)}`);

  const sectorShift = bytes.readUInt16LE(30);
  const miniShift = bytes.readUInt16LE(32);
  const sectorSize = 1 << sectorShift;
  const miniSize = 1 << miniShift;
  const miniCutoff = bytes.readUInt32LE(56);
  const firstDir = bytes.readUInt32LE(48);
  const firstMiniFat = bytes.readUInt32LE(60);
  const difatCount = bytes.readUInt32LE(72);
  const firstDifat = bytes.readUInt32LE(68);

  const at = (sector) => 512 + sector * sectorSize;

  // The FAT, through the DIFAT: the header holds the first 109 entries, then a chain.
  const fatSectors = [];
  for (let i = 0; i < 109; i++) {
    const s = bytes.readUInt32LE(76 + i * 4);
    if (s === 0xFFFFFFFF) break;
    fatSectors.push(s);
  }

  let difat = firstDifat;
  for (let n = 0; n < difatCount && difat !== 0xFFFFFFFF; n++) {
    const base = at(difat);
    for (let i = 0; i < (sectorSize / 4) - 1; i++) {
      const s = bytes.readUInt32LE(base + i * 4);
      if (s !== 0xFFFFFFFF) fatSectors.push(s);
    }
    difat = bytes.readUInt32LE(base + sectorSize - 4);
  }

  const fat = [];
  for (const s of fatSectors) {
    const base = at(s);
    for (let i = 0; i < sectorSize / 4; i++) fat.push(bytes.readUInt32LE(base + i * 4));
  }

  const chain = (start) => {
    const out = [];
    let s = start;
    while (s !== 0xFFFFFFFE && s !== 0xFFFFFFFF && out.length < 100000) {
      out.push(s);
      s = fat[s];
      if (s === undefined) break;
    }
    return out;
  };

  const readChain = (start, size) => {
    const parts = chain(start).map((s) => bytes.subarray(at(s), at(s) + sectorSize));
    const whole = Buffer.concat(parts);
    return size === undefined ? whole : whole.subarray(0, size);
  };

  // The directory: 128-byte entries.
  const dirBytes = readChain(firstDir);
  const entries = [];
  for (let i = 0; i * 128 < dirBytes.length; i++) {
    const o = i * 128;
    const nameLen = dirBytes.readUInt16LE(o + 64);
    if (nameLen === 0) continue;
    const name = dirBytes.subarray(o, o + Math.max(0, nameLen - 2)).toString("utf16le");
    entries.push({
      index: i,
      name,
      type: dirBytes.readUInt8(o + 66),        // 1 storage, 2 stream, 5 root
      left: dirBytes.readUInt32LE(o + 68),
      right: dirBytes.readUInt32LE(o + 72),
      child: dirBytes.readUInt32LE(o + 76),
      start: dirBytes.readUInt32LE(o + 116),
      size: Number(dirBytes.readBigUInt64LE(o + 120)),
    });
  }

  // The mini stream lives in the root entry's chain; small streams are cut from it.
  const root = entries.find((e) => e.type === 5);
  const miniStream = root && root.size > 0 ? readChain(root.start) : Buffer.alloc(0);
  const miniFat = [];
  {
    const fatBytes = firstMiniFat === 0xFFFFFFFE ? Buffer.alloc(0) : readChain(firstMiniFat);
    for (let i = 0; i * 4 < fatBytes.length; i++) miniFat.push(fatBytes.readUInt32LE(i * 4));
  }

  const readStream = (entry) => {
    if (entry.size >= miniCutoff) return readChain(entry.start, entry.size);
    const parts = [];
    let s = entry.start;
    while (s !== 0xFFFFFFFE && s !== 0xFFFFFFFF && parts.length < 100000) {
      parts.push(miniStream.subarray(s * miniSize, (s + 1) * miniSize));
      s = miniFat[s];
      if (s === undefined) break;
    }
    return Buffer.concat(parts).subarray(0, entry.size);
  };

  // Paths, by walking the red-black sibling trees under each storage.
  const paths = new Map();
  const walk = (index, prefix) => {
    if (index === 0xFFFFFFFF) return;
    const e = entries[index];
    if (!e) return;
    walk(e.left, prefix);
    const path = `${prefix}/${e.name}`;
    paths.set(path, e);
    if (e.type === 1) walk(e.child, path);
    walk(e.right, prefix);
  };
  if (root) walk(root.child, "");

  return { entries, paths, readStream };
}
