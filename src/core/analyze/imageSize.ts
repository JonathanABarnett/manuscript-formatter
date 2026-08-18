/**
 * Pixel dimensions read from the first few bytes of an image, without
 * decoding it. Enough to tell whether a picture has the pixels to print
 * sharply at the size it is placed on the page.
 */
export interface PixelSize {
  width: number;
  height: number;
}

export function readPixelSize(bytes: Uint8Array): PixelSize | null {
  if (bytes.length < 10) return null;
  return readPng(bytes) ?? readJpeg(bytes) ?? readGif(bytes) ?? readBmp(bytes) ?? null;
}

const be32 = (b: Uint8Array, at: number): number =>
  ((b[at] << 24) | (b[at + 1] << 16) | (b[at + 2] << 8) | b[at + 3]) >>> 0;
const be16 = (b: Uint8Array, at: number): number => (b[at] << 8) | b[at + 1];
const le32 = (b: Uint8Array, at: number): number =>
  (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0;
const le16 = (b: Uint8Array, at: number): number => b[at] | (b[at + 1] << 8);

function readPng(b: Uint8Array): PixelSize | null {
  const signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  if (b.length < 24 || !signature.every((v, i) => b[i] === v)) return null;
  // The IHDR chunk is required to come first: length, "IHDR", width, height.
  if (String.fromCharCode(b[12], b[13], b[14], b[15]) !== 'IHDR') return null;
  return { width: be32(b, 16), height: be32(b, 20) };
}

function readJpeg(b: Uint8Array): PixelSize | null {
  if (b[0] !== 0xff || b[1] !== 0xd8) return null;
  let at = 2;
  while (at + 9 < b.length) {
    if (b[at] !== 0xff) {
      at++;
      continue;
    }
    const marker = b[at + 1];
    // Padding bytes and standalone markers carry no length.
    if (marker === 0xff) {
      at++;
      continue;
    }
    if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
      at += 2;
      continue;
    }
    const length = be16(b, at + 2);
    const isFrame =
      marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isFrame) {
      return { height: be16(b, at + 5), width: be16(b, at + 7) };
    }
    if (marker === 0xda || marker === 0xd9) return null; // scan data or end: no frame seen
    at += 2 + length;
  }
  return null;
}

function readGif(b: Uint8Array): PixelSize | null {
  const head = String.fromCharCode(...b.slice(0, 6));
  if (head !== 'GIF87a' && head !== 'GIF89a') return null;
  return { width: le16(b, 6), height: le16(b, 8) };
}

function readBmp(b: Uint8Array): PixelSize | null {
  if (b[0] !== 0x42 || b[1] !== 0x4d || b.length < 26) return null;
  const height = le32(b, 22);
  // Height is signed: negative means top-down rows.
  return { width: le32(b, 18), height: height > 0x7fffffff ? 0x100000000 - height : height };
}
