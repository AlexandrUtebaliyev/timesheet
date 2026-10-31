// Разовый генератор плейсхолдер-иконок PWA (без canvas и внешних зависимостей).
// Запуск: node tests/gen-icons.mjs
import { deflateSync } from "node:zlib";
import { writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const __dirname = dirname(fileURLToPath(import.meta.url));

function crc32(buf) {
  let c;
  const table = crc32.table || (crc32.table = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c;
    }
    return t;
  })());
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const typeBuf = Buffer.from(type, "ascii");
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
  return Buffer.concat([len, typeBuf, data, crcBuf]);
}

function hexToRgb(hex) {
  const n = parseInt(hex.slice(1), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function buildIcon(size) {
  const bg = hexToRgb("#15171A");
  const accent = hexToRgb("#5B9BFF");
  const margin = Math.round(size * 0.22);
  const barWidth = Math.round(size * 0.1);
  const rows = [];

  for (let y = 0; y < size; y++) {
    const row = [0]; // filter byte: none
    for (let x = 0; x < size; x++) {
      const insideFrame = x >= margin && x < size - margin && y >= margin && y < size - margin;
      let r = bg[0], g = bg[1], b = bg[2];
      if (insideFrame) {
        const relY = y - margin;
        const frameH = size - 2 * margin;
        const isBar = relY % Math.round(frameH / 4) < barWidth / 2;
        if (isBar) {
          r = accent[0]; g = accent[1]; b = accent[2];
        } else {
          r = 0x1f; g = 0x23; b = 0x28;
        }
      }
      row.push(r, g, b);
    }
    rows.push(Buffer.from(row));
  }

  const raw = Buffer.concat(rows);
  const compressed = deflateSync(raw);

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // color type RGB
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  return Buffer.concat([
    signature,
    chunk("IHDR", ihdr),
    chunk("IDAT", compressed),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

const iconsDir = join(__dirname, "..", "icons");
for (const size of [192, 512]) {
  const buf = buildIcon(size);
  writeFileSync(join(iconsDir, `icon-${size}.png`), buf);
  console.log(`icons/icon-${size}.png — ${buf.length} байт`);
}
