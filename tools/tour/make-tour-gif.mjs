/*
 * Assembles assets\images\tour.gif from the four stills capture-tour.mjs writes.
 *
 * Everything is done here - PNG decode, palette, quantisation, LZW - because this machine
 * carries no ffmpeg or ImageMagick and an asset generator that only works on the machine it
 * was written on is the lesson TwinFixture already taught. Node's zlib covers the only part
 * that would be real work.
 *
 *     node tools\tour\make-tour-gif.mjs
 *
 * Each scene holds for a beat, then crossfades into the next; the last fades back into the
 * first so the loop has no seam. One global palette serves every frame: the scenes share a
 * theme, and a per-frame palette would make the GIF flicker at each fade.
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { inflateSync } from "node:zlib";

const SCENES = ["editor", "completion", "problems", "debugger"];
const HOLD_CS = 320;      // how long each scene stands, in GIF centiseconds
const FADE_STEPS = 8;     // frames per crossfade
const FADE_CS = 6;        // per fade frame

const imagesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "..", "assets", "images");

// ---------------------------------------------------------------------------- PNG in

/** Enough of a PNG reader for our own stills: 8-bit RGB or RGBA, not interlaced. */
function decodePng(bytes) {
  if (bytes.readUInt32BE(0) !== 0x89504e47) { throw new Error("not a PNG"); }

  let width = 0;
  let height = 0;
  let colorType = -1;
  const idat = [];

  for (let at = 8; at < bytes.length;) {
    const length = bytes.readUInt32BE(at);
    const type = bytes.toString("latin1", at + 4, at + 8);
    const data = bytes.subarray(at + 8, at + 8 + length);
    if (type === "IHDR") {
      width = data.readUInt32BE(0);
      height = data.readUInt32BE(4);
      const bitDepth = data[8];
      colorType = data[9];
      if (bitDepth !== 8 || (colorType !== 2 && colorType !== 6) || data[12] !== 0) {
        throw new Error(`unhandled PNG shape: depth ${bitDepth} colour ${colorType} interlace ${data[12]}`);
      }
    } else if (type === "IDAT") {
      idat.push(data);
    } else if (type === "IEND") {
      break;
    }
    at += 12 + length;
  }

  const sourcePixel = colorType === 6 ? 4 : 3;
  const raw = inflateSync(Buffer.concat(idat));
  const rowBytes = width * sourcePixel;
  const rgb = Buffer.allocUnsafe(width * height * 3);

  const previous = Buffer.alloc(rowBytes);
  const current = Buffer.allocUnsafe(rowBytes);
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (rowBytes + 1)];
    const line = raw.subarray(y * (rowBytes + 1) + 1, (y + 1) * (rowBytes + 1));

    for (let i = 0; i < rowBytes; i++) {
      const left = i >= sourcePixel ? current[i - sourcePixel] : 0;
      const up = previous[i];
      const upLeft = i >= sourcePixel ? previous[i - sourcePixel] : 0;
      let value = line[i];
      switch (filter) {
        case 1: value += left; break;
        case 2: value += up; break;
        case 3: value += (left + up) >> 1; break;
        case 4: {
          const p = left + up - upLeft;
          const pa = Math.abs(p - left);
          const pb = Math.abs(p - up);
          const pc = Math.abs(p - upLeft);
          value += pa <= pb && pa <= pc ? left : pb <= pc ? up : upLeft;
          break;
        }
        default: break;
      }
      current[i] = value & 0xff;
    }

    let target = y * width * 3;
    for (let x = 0; x < width; x++) {
      rgb[target++] = current[x * sourcePixel];
      rgb[target++] = current[x * sourcePixel + 1];
      rgb[target++] = current[x * sourcePixel + 2];
    }
    current.copy(previous);
  }

  return { width, height, rgb };
}

// ---------------------------------------------------------------------------- palette

/** Median cut over a sample of every frame INCLUDING fade midpoints: a fade's in-between
 *  colours belong to no still, and a palette built without them bands at each transition. */
function buildPalette(samples) {
  let boxes = [{ pixels: samples }];

  const widest = (box) => {
    const low = [255, 255, 255];
    const high = [0, 0, 0];
    for (const pixel of box.pixels) {
      for (let c = 0; c < 3; c++) {
        const v = (pixel >> (c * 8)) & 0xff;
        if (v < low[c]) { low[c] = v; }
        if (v > high[c]) { high[c] = v; }
      }
    }
    let axis = 0;
    let range = -1;
    for (let c = 0; c < 3; c++) {
      if (high[c] - low[c] > range) { range = high[c] - low[c]; axis = c; }
    }
    return { axis, range };
  };

  while (boxes.length < 256) {
    let pick = -1;
    let pickScore = 0;
    for (let i = 0; i < boxes.length; i++) {
      const score = boxes[i].pixels.length * (boxes[i].spread ??= widest(boxes[i]).range);
      if (boxes[i].pixels.length > 1 && score > pickScore) { pick = i; pickScore = score; }
    }
    if (pick < 0) { break; }

    const box = boxes[pick];
    const { axis } = widest(box);
    box.pixels.sort((a, b) => ((a >> (axis * 8)) & 0xff) - ((b >> (axis * 8)) & 0xff));
    const half = box.pixels.length >> 1;
    boxes.splice(pick, 1,
      { pixels: box.pixels.slice(0, half) },
      { pixels: box.pixels.slice(half) });
  }

  const palette = Buffer.alloc(256 * 3);
  boxes.forEach((box, index) => {
    let r = 0; let g = 0; let b = 0;
    for (const pixel of box.pixels) {
      r += pixel & 0xff; g += (pixel >> 8) & 0xff; b += (pixel >> 16) & 0xff;
    }
    const n = box.pixels.length || 1;
    palette[index * 3] = Math.round(r / n);
    palette[index * 3 + 1] = Math.round(g / n);
    palette[index * 3 + 2] = Math.round(b / n);
  });
  return { palette, count: boxes.length };
}

function quantiser(palette, count) {
  const cache = new Map();
  return (r, g, b) => {
    const key = (r << 16) | (g << 8) | b;
    let index = cache.get(key);
    if (index !== undefined) { return index; }
    let best = 0;
    let bestDistance = Infinity;
    for (let i = 0; i < count; i++) {
      const dr = r - palette[i * 3];
      const dg = g - palette[i * 3 + 1];
      const db = b - palette[i * 3 + 2];
      const distance = dr * dr + dg * dg + db * db;
      if (distance < bestDistance) { bestDistance = distance; best = i; }
    }
    cache.set(key, best);
    return best;
  };
}

// ---------------------------------------------------------------------------- GIF out

/** GIF's LZW with variable code width, a 12-bit dictionary, and a reset when it fills. */
function lzwEncode(indices) {
  const CLEAR = 256;
  const EOI = 257;
  const out = [];
  let bits = 0;
  let bitCount = 0;
  let codeSize = 9;

  const emit = (code) => {
    bits |= code << bitCount;
    bitCount += codeSize;
    while (bitCount >= 8) {
      out.push(bits & 0xff);
      bits >>>= 8;
      bitCount -= 8;
    }
  };

  // prefix code (12 bits) x next byte (8 bits) -> code
  let dictionary = new Int32Array(4096 * 256).fill(-1);
  let next = 258;
  emit(CLEAR);

  let prefix = indices[0];
  for (let i = 1; i < indices.length; i++) {
    const symbol = indices[i];
    const slot = prefix * 256 + symbol;
    if (dictionary[slot] >= 0) {
      prefix = dictionary[slot];
      continue;
    }

    emit(prefix);
    if (next < 4096) {
      dictionary[slot] = next;
      if (next === 1 << codeSize && codeSize < 12) { codeSize++; }
      next++;
    } else {
      emit(CLEAR);
      dictionary = new Int32Array(4096 * 256).fill(-1);
      next = 258;
      codeSize = 9;
    }
    prefix = symbol;
  }

  emit(prefix);
  emit(EOI);
  if (bitCount > 0) { out.push(bits & 0xff); }
  return Buffer.from(out);
}

function subBlocks(data) {
  const parts = [];
  for (let at = 0; at < data.length; at += 255) {
    const slice = data.subarray(at, at + 255);
    parts.push(Buffer.from([slice.length]), slice);
  }
  parts.push(Buffer.from([0]));
  return Buffer.concat(parts);
}

function gifFrame(width, height, indices, delayCs) {
  const graphicControl = Buffer.from([
    0x21, 0xf9, 0x04,
    0x04,                     // disposal 1: leave the frame in place
    delayCs & 0xff, (delayCs >> 8) & 0xff,
    0, 0,
  ]);
  const descriptor = Buffer.alloc(10);
  descriptor[0] = 0x2c;
  descriptor.writeUInt16LE(width, 5);
  descriptor.writeUInt16LE(height, 7);
  return Buffer.concat([graphicControl, descriptor, Buffer.from([8]), subBlocks(lzwEncode(indices))]);
}

// ---------------------------------------------------------------------------- assembly

const stills = SCENES.map((name) => {
  const image = decodePng(readFileSync(join(imagesDir, `tour-${name}.png`)));
  console.log(`${name}: ${image.width}x${image.height}`);
  return image;
});

const { width, height } = stills[0];
for (const still of stills) {
  if (still.width !== width || still.height !== height) {
    throw new Error("the stills disagree about their size; reshoot with capture-tour.mjs");
  }
}

const blend = (a, b, t) => {
  const mixed = Buffer.allocUnsafe(a.length);
  for (let i = 0; i < a.length; i++) {
    mixed[i] = (a[i] * (1 - t) + b[i] * t + 0.5) | 0;
  }
  return mixed;
};

console.log("palette...");
const samples = [];
const sample = (rgb) => {
  for (let i = 0; i < rgb.length; i += 3 * 5) {
    samples.push(rgb[i] | (rgb[i + 1] << 8) | (rgb[i + 2] << 16));
  }
};
for (let i = 0; i < stills.length; i++) {
  sample(stills[i].rgb);
  sample(blend(stills[i].rgb, stills[(i + 1) % stills.length].rgb, 0.5));
}
const { palette, count } = buildPalette(samples);
const nearest = quantiser(palette, count);

const toIndices = (rgb) => {
  const indices = new Uint8Array(rgb.length / 3);
  for (let i = 0, p = 0; i < rgb.length; i += 3, p++) {
    indices[p] = nearest(rgb[i], rgb[i + 1], rgb[i + 2]);
  }
  return indices;
};

console.log("frames...");
const frames = [];
for (let i = 0; i < stills.length; i++) {
  frames.push(gifFrame(width, height, toIndices(stills[i].rgb), HOLD_CS));
  const following = stills[(i + 1) % stills.length];
  for (let step = 1; step <= FADE_STEPS; step++) {
    const t = step / (FADE_STEPS + 1);
    frames.push(gifFrame(width, height, toIndices(blend(stills[i].rgb, following.rgb, t)), FADE_CS));
  }
  console.log(`  scene ${SCENES[i]} + fade`);
}

const screen = Buffer.alloc(7);
screen.writeUInt16LE(width, 0);
screen.writeUInt16LE(height, 2);
screen[4] = 0xf7;   // global 256-colour table
const loopForever = Buffer.from([
  0x21, 0xff, 0x0b,
  ...Buffer.from("NETSCAPE2.0", "latin1"),
  0x03, 0x01, 0x00, 0x00, 0x00,
]);

const gif = Buffer.concat([
  Buffer.from("GIF89a", "latin1"),
  screen,
  palette,
  loopForever,
  ...frames,
  Buffer.from([0x3b]),
]);

const target = join(imagesDir, "tour.gif");
writeFileSync(target, gif);
console.log(`${target}: ${(gif.length / 1024 / 1024).toFixed(2)} MB, ${frames.length} frames, ${count} colours`);
