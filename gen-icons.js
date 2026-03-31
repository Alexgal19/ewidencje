// One-time icon generator — run: node gen-icons.js
// Creates proper PNG icons using pure Node.js (no external deps)
'use strict';
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

function uint32BE(n) {
  const b = Buffer.alloc(4);
  b.writeUInt32BE(n, 0);
  return b;
}

function crc32(buf) {
  let c = 0xffffffff;
  for (const byte of buf) {
    c ^= byte;
    for (let i = 0; i < 8; i++) c = (c >>> 1) ^ (c & 1 ? 0xedb88320 : 0);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const t = Buffer.from(type, 'ascii');
  const len = uint32BE(data.length);
  const crcBuf = Buffer.concat([t, data]);
  return Buffer.concat([len, t, data, uint32BE(crc32(crcBuf))]);
}

function makePNG(size) {
  // Draw rounded rect with gradient-like green and a white "E" letter
  const rows = [];
  const cx = size / 2, cy = size / 2, r = size * 0.218;

  for (let y = 0; y < size; y++) {
    const row = [];
    for (let x = 0; x < size; x++) {
      // Rounded rectangle mask
      const dx = Math.max(Math.abs(x - cx) - (cx - r), 0);
      const dy = Math.max(Math.abs(y - cy) - (cy - r), 0);
      const dist = Math.sqrt(dx * dx + dy * dy);
      const inside = dist <= r;

      if (!inside) {
        row.push(11, 17, 32, 0); // transparent bg (#0B1120)
      } else {
        // Gradient: top-left #10B981 → bottom-right #065F46
        const t = (x + y) / (size * 2);
        const R = Math.round(16  + (6  - 16)  * t);
        const G = Math.round(185 + (95 - 185) * t);
        const B = Math.round(129 + (70 - 129) * t);

        // Draw a simple white car silhouette (basic pixel art, 40% of icon size)
        const carSize = size * 0.55;
        const carX = (size - carSize) / 2;
        const carY = (size - carSize * 0.6) / 2;
        // Body rectangle
        const bodyT = carY + carSize * 0.3, bodyB = carY + carSize * 0.6;
        const bodyL = carX, bodyR = carX + carSize;
        // Roof
        const roofT = carY + carSize * 0.05, roofB = carY + carSize * 0.3;
        const roofL = carX + carSize * 0.15, roofR = carX + carSize * 0.85;
        // Wheels
        const wR = carSize * 0.12;
        const w1x = carX + carSize * 0.22, w1y = bodyB;
        const w2x = carX + carSize * 0.78, w2y = bodyB;

        const inBody = x >= bodyL && x <= bodyR && y >= bodyT && y <= bodyB;
        const inRoof = x >= roofL && x <= roofR && y >= roofT && y <= roofB;
        const inW1 = Math.hypot(x - w1x, y - w1y) <= wR;
        const inW2 = Math.hypot(x - w2x, y - w2y) <= wR;
        const inCar = inBody || inRoof || inW1 || inW2;

        if (inCar) {
          row.push(255, 255, 255, 220); // white car
        } else {
          row.push(R, G, B, 255);
        }
      }
    }
    rows.push(row);
  }

  // Build raw image data (RGBA, filter byte 0 per row)
  const raw = Buffer.alloc(size * (size * 4 + 1));
  let off = 0;
  for (const row of rows) {
    raw[off++] = 0; // filter none
    for (const b of row) raw[off++] = b;
  }

  const compressed = zlib.deflateSync(raw, { level: 9 });
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const ihdr = chunk('IHDR', Buffer.concat([
    uint32BE(size), uint32BE(size),
    Buffer.from([8, 6, 0, 0, 0]) // 8-bit RGBA
  ]));
  const idat = chunk('IDAT', compressed);
  const iend = chunk('IEND', Buffer.alloc(0));
  return Buffer.concat([sig, ihdr, idat, iend]);
}

const dir = path.join(__dirname, 'public', 'icons');
fs.mkdirSync(dir, { recursive: true });

for (const size of [192, 512]) {
  console.log(`Generating icon-${size}.png...`);
  const buf = makePNG(size);
  fs.writeFileSync(path.join(dir, `icon-${size}.png`), buf);
  console.log(`  → ${buf.length} bytes`);
}
console.log('Done.');
