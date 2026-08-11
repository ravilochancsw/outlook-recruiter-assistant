import { deflateSync } from 'node:zlib';
import { writeFileSync, mkdirSync } from 'node:fs';

/**
 * Generates the toolbar icons.
 *
 * Hand-rolled because the extension has no runtime dependencies and Chrome needs
 * raster icons — an SVG will not load. Without these the toolbar shows a generic
 * puzzle piece, which is a problem when the popup is the place you go to find out
 * why something did not happen.
 *
 * Run with: node scripts/make-icons.mjs
 */

const BG = [15, 108, 189];
const FG = [255, 255, 255];

function render(size) {
  const buf = Buffer.alloc(size * size * 4, 0);
  const radius = Math.round(size * 0.22);

  // Rounded-rect mask.
  const inside = (x, y) => {
    const cx = Math.min(Math.max(x, radius), size - 1 - radius);
    const cy = Math.min(Math.max(y, radius), size - 1 - radius);
    return (x - cx) ** 2 + (y - cy) ** 2 <= radius * radius;
  };

  const ex0 = Math.round(size * 0.22);
  const ex1 = Math.round(size * 0.78);
  const ey0 = Math.round(size * 0.32);
  const ey1 = Math.round(size * 0.68);
  const stroke = Math.max(1, Math.round(size / 14));
  const midX = (ex0 + ex1) / 2;

  for (let y = 0; y < size; y += 1) {
    for (let x = 0; x < size; x += 1) {
      if (!inside(x, y)) continue;
      let colour = BG;

      const onFrame =
        (x >= ex0 && x <= ex1 && (Math.abs(y - ey0) < stroke || Math.abs(y - ey1) < stroke)) ||
        (y >= ey0 && y <= ey1 && (Math.abs(x - ex0) < stroke || Math.abs(x - ex1) < stroke));

      // Envelope flap: two diagonals meeting in the middle.
      const leftY = ey0 + ((x - ex0) / (midX - ex0)) * (ey1 - ey0) * 0.6;
      const rightY = ey0 + ((ex1 - x) / (ex1 - midX)) * (ey1 - ey0) * 0.6;
      const onFlap =
        (x >= ex0 && x <= midX && Math.abs(y - leftY) < stroke) ||
        (x > midX && x <= ex1 && Math.abs(y - rightY) < stroke);

      if (onFrame || onFlap) colour = FG;

      const i = (y * size + x) * 4;
      buf[i] = colour[0];
      buf[i + 1] = colour[1];
      buf[i + 2] = colour[2];
      buf[i + 3] = 255;
    }
  }
  return buf;
}

function crc32(buf) {
  let c = ~0;
  for (const byte of buf) {
    c ^= byte;
    for (let k = 0; k < 8; k += 1) c = (c >>> 1) ^ (0xedb88320 & -(c & 1));
  }
  return ~c >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

function png(size) {
  const rgba = render(size);
  const stride = size * 4 + 1;
  const raw = Buffer.alloc(stride * size);
  for (let y = 0; y < size; y += 1) {
    raw[y * stride] = 0; // filter: none
    rgba.copy(raw, y * stride + 1, y * size * 4, (y + 1) * size * 4);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // colour type: RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

mkdirSync('src/icons', { recursive: true });
for (const size of [16, 32, 48, 128]) {
  writeFileSync(`src/icons/icon${size}.png`, png(size));
  console.log(`wrote src/icons/icon${size}.png`);
}
