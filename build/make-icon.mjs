/**
 * Generates build/icon.png and build/icon.ico — an open book on the warm brown
 * used throughout the UI. Written by hand rather than committed as an opaque
 * binary so the app icon stays editable with no image tooling installed.
 *
 *   node build/make-icon.mjs
 */
import { deflateSync } from 'node:zlib';
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const SIZE = 512;
const here = dirname(fileURLToPath(import.meta.url));

const INK = [0x2c, 0x1c, 0x10, 0xff];
const BROWN = [0x7a, 0x4f, 0x2a, 0xff];
const PAGE = [0xfa, 0xf6, 0xef, 0xff];
const PAGE_EDGE = [0xe2, 0xd7, 0xc5, 0xff];
const RULE = [0xb9, 0xa8, 0x92, 0xff];

/** RGBA canvas, row-major, initialised transparent. */
const canvas = new Uint8Array(SIZE * SIZE * 4);

function blend(x, y, [r, g, b, a]) {
  if (x < 0 || y < 0 || x >= SIZE || y >= SIZE || a === 0) return;
  const i = (y * SIZE + x) * 4;
  const srcA = a / 255;
  const dstA = canvas[i + 3] / 255;
  const outA = srcA + dstA * (1 - srcA);
  if (outA === 0) return;
  for (let c = 0; c < 3; c++) {
    canvas[i + c] = Math.round((( [r, g, b][c] * srcA) + canvas[i + c] * dstA * (1 - srcA)) / outA);
  }
  canvas[i + 3] = Math.round(outA * 255);
}

/** Anti-aliased rounded rectangle via per-pixel coverage sampling. */
function roundRect(x0, y0, w, h, radius, colour) {
  const x1 = x0 + w;
  const y1 = y0 + h;
  const r = Math.min(radius, w / 2, h / 2);
  for (let y = Math.floor(y0) - 1; y <= Math.ceil(y1) + 1; y++) {
    for (let x = Math.floor(x0) - 1; x <= Math.ceil(x1) + 1; x++) {
      let hits = 0;
      // 3x3 supersample keeps the curves smooth at 16px without a blur pass.
      for (let sy = 0; sy < 3; sy++) {
        for (let sx = 0; sx < 3; sx++) {
          const px = x + (sx + 0.5) / 3;
          const py = y + (sy + 0.5) / 3;
          if (px < x0 || px > x1 || py < y0 || py > y1) continue;
          const cx = Math.min(Math.max(px, x0 + r), x1 - r);
          const cy = Math.min(Math.max(py, y0 + r), y1 - r);
          const dx = px - cx;
          const dy = py - cy;
          if (dx * dx + dy * dy <= r * r) hits++;
        }
      }
      if (hits === 0) continue;
      blend(x, y, [colour[0], colour[1], colour[2], Math.round((colour[3] * hits) / 9)]);
    }
  }
}

// --- the mark: a book, open, seen head-on -----------------------------------

roundRect(0, 0, SIZE, SIZE, 112, BROWN);

const PAGE_TOP = 150;
const PAGE_H = 226;
const PAGE_W = 156;
const GUTTER = 14;
const MID = SIZE / 2;

// Two leaves, each tucked slightly under the spine.
for (const side of [-1, 1]) {
  const x = side === -1 ? MID - GUTTER / 2 - PAGE_W : MID + GUTTER / 2;
  roundRect(x, PAGE_TOP - 6, PAGE_W, PAGE_H + 12, 10, PAGE_EDGE);
  roundRect(x, PAGE_TOP, PAGE_W, PAGE_H, 8, PAGE);

  // Ruled lines, shorter on the last row so it reads as running text.
  const rows = 7;
  const step = 26;
  for (let row = 0; row < rows; row++) {
    const last = row === rows - 1;
    const inset = 22;
    const width = (PAGE_W - inset * 2) * (last ? 0.55 : 1);
    roundRect(x + inset, PAGE_TOP + 34 + row * step, width, 7, 3.5, RULE);
  }
}

// The spine sits above both leaves.
roundRect(MID - GUTTER / 2 - 3, PAGE_TOP - 12, GUTTER + 6, PAGE_H + 24, 5, INK);

// --- PNG ---------------------------------------------------------------------

function chunk(type, data) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, 'latin1'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body) >>> 0);
  return Buffer.concat([length, body, crc]);
}

const CRC_TABLE = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  return table;
})();

function crc32(buf) {
  let c = -1;
  for (const byte of buf) c = CRC_TABLE[(c ^ byte) & 0xff] ^ (c >>> 8);
  return c ^ -1;
}

function encodePng(pixels, size) {
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0; // filter: none
    Buffer.from(pixels.buffer, y * size * 4, size * 4).copy(raw, y * (size * 4 + 1) + 1);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // truecolour with alpha
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

/** Box-filter downscale, so the small ICO entries stay legible. */
function resize(pixels, from, to) {
  const out = new Uint8Array(to * to * 4);
  const ratio = from / to;
  for (let y = 0; y < to; y++) {
    for (let x = 0; x < to; x++) {
      const x0 = Math.floor(x * ratio);
      const y0 = Math.floor(y * ratio);
      const x1 = Math.min(from, Math.ceil((x + 1) * ratio));
      const y1 = Math.min(from, Math.ceil((y + 1) * ratio));
      let r = 0, g = 0, b = 0, a = 0, n = 0;
      for (let sy = y0; sy < y1; sy++) {
        for (let sx = x0; sx < x1; sx++) {
          const i = (sy * from + sx) * 4;
          const alpha = pixels[i + 3] / 255;
          r += pixels[i] * alpha;
          g += pixels[i + 1] * alpha;
          b += pixels[i + 2] * alpha;
          a += pixels[i + 3];
          n++;
        }
      }
      const o = (y * to + x) * 4;
      const alphaSum = a / 255;
      out[o] = alphaSum ? Math.round(r / alphaSum) : 0;
      out[o + 1] = alphaSum ? Math.round(g / alphaSum) : 0;
      out[o + 2] = alphaSum ? Math.round(b / alphaSum) : 0;
      out[o + 3] = Math.round(a / n);
    }
  }
  return out;
}

/** ICO carrying PNG-compressed entries, which Windows Vista onward accepts. */
function encodeIco(sizes) {
  const entries = sizes.map((size) => ({
    size,
    png: encodePng(size === SIZE ? canvas : resize(canvas, SIZE, size), size),
  }));
  const header = Buffer.alloc(6 + entries.length * 16);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2); // type: icon
  header.writeUInt16LE(entries.length, 4);
  let offset = header.length;
  entries.forEach((entry, i) => {
    const at = 6 + i * 16;
    header[at] = entry.size >= 256 ? 0 : entry.size;
    header[at + 1] = entry.size >= 256 ? 0 : entry.size;
    header.writeUInt16LE(1, at + 4); // colour planes
    header.writeUInt16LE(32, at + 6); // bits per pixel
    header.writeUInt32LE(entry.png.length, at + 8);
    header.writeUInt32LE(offset, at + 12);
    offset += entry.png.length;
  });
  return Buffer.concat([header, ...entries.map((e) => e.png)]);
}

writeFileSync(join(here, 'icon.png'), encodePng(canvas, SIZE));
writeFileSync(join(here, 'icon.ico'), encodeIco([16, 24, 32, 48, 64, 128, 256]));
console.log('wrote build/icon.png and build/icon.ico');
