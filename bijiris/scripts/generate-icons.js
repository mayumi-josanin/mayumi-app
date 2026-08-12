const fs = require("node:fs");
const path = require("node:path");
const zlib = require("node:zlib");

const OUT_DIRS = [
  path.join(__dirname, "..", "icons"),
  path.join(__dirname, "..", "customer-app", "icons"),
  path.join(__dirname, "..", "admin-app", "icons"),
];

function makeCrcTable() {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
}

const CRC_TABLE = makeCrcTable();

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const typeBuffer = Buffer.from(type);
  const length = Buffer.alloc(4);
  const crc = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);
  crc.writeUInt32BE(crc32(Buffer.concat([typeBuffer, data])), 0);
  return Buffer.concat([length, typeBuffer, data, crc]);
}

function insideRoundedRect(x, y, size, inset, radius) {
  const left = inset;
  const right = size - inset;
  const top = inset;
  const bottom = size - inset;
  const clampedX = Math.max(left + radius, Math.min(x, right - radius));
  const clampedY = Math.max(top + radius, Math.min(y, bottom - radius));
  const dx = x - clampedX;
  const dy = y - clampedY;
  return dx * dx + dy * dy <= radius * radius;
}

function makeIcon(size) {
  const rowLength = size * 4 + 1;
  const raw = Buffer.alloc(rowLength * size);
  const inset = Math.round(size * 0.13);
  const radius = Math.round(size * 0.08);

  for (let y = 0; y < size; y += 1) {
    const row = y * rowLength;
    raw[row] = 0;
    for (let x = 0; x < size; x += 1) {
      const offset = row + 1 + x * 4;
      const line = Math.round(size * 0.13);
      const bar = Math.round(size * 0.09);
      const bgShade = Math.round((x / size) * 12);

      let r = 15 + bgShade;
      let g = 118 + bgShade;
      let b = 110 + bgShade;

      if (insideRoundedRect(x, y, size, inset, radius)) {
        r = 246;
        g = 248;
        b = 247;
      }

      const leftBar =
        x > inset + line &&
        x < inset + line + bar &&
        y > inset + line &&
        y < size - inset - line;
      const centerBar =
        x > size / 2 - bar / 2 &&
        x < size / 2 + bar / 2 &&
        y > inset + line &&
        y < size - inset - line;
      const rightBar =
        x > size - inset - line - bar &&
        x < size - inset - line &&
        y > inset + line &&
        y < size - inset - line;
      const goldBand =
        y > size * 0.3 &&
        y < size * 0.3 + bar &&
        x > inset + line &&
        x < size - inset - line;
      const coralBand =
        y > size * 0.59 &&
        y < size * 0.59 + bar &&
        x > inset + line &&
        x < size - inset - line;

      if (leftBar || centerBar || rightBar) {
        r = 15;
        g = 118;
        b = 110;
      }
      if (goldBand) {
        r = 242;
        g = 201;
        b = 76;
      }
      if (coralBand) {
        r = 232;
        g = 93;
        b = 79;
      }

      raw[offset] = r;
      raw[offset + 1] = g;
      raw[offset + 2] = b;
      raw[offset + 3] = 255;
    }
  }

  const signature = Buffer.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
  ]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const png = Buffer.concat([
    signature,
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
  return png;
}

for (const outDir of OUT_DIRS) {
  fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(path.join(outDir, "icon-192.png"), makeIcon(192));
  fs.writeFileSync(path.join(outDir, "icon-512.png"), makeIcon(512));
}
