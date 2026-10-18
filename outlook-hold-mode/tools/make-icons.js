// Draws the Hold Mode icon (a clock with a pause badge) as PNGs in ../assets.
// Run with: node outlook-hold-mode/tools/make-icons.js
const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const SIZES = [16, 20, 32, 64, 80, 128];
const BLUE = [15, 108, 189];     // clock
const AMBER = [201, 132, 0];     // pause badge
const SS = 4;                    // supersampling factor

function render(size) {
  const N = size * SS;
  const px = new Float32Array(size * size * 4);
  const c = N / 2;
  const R = N * 0.40;            // clock radius (outer edge of ring)
  const ring = Math.max(N * 0.085, SS * 1.6);
  const hand = Math.max(N * 0.07, SS * 1.4);
  const badgeR = N * 0.25, bx = N * 0.74, by = N * 0.74;

  function segDist(x, y, x1, y1, x2, y2) {
    const dx = x2 - x1, dy = y2 - y1;
    const t = Math.max(0, Math.min(1, ((x - x1) * dx + (y - y1) * dy) / (dx * dx + dy * dy)));
    return Math.hypot(x - (x1 + t * dx), y - (y1 + t * dy));
  }

  for (let sy = 0; sy < N; sy++) {
    for (let sx = 0; sx < N; sx++) {
      const x = sx + 0.5, y = sy + 0.5;
      let col = null;

      const dc = Math.hypot(x - c, y - c);
      const inBadgeHalo = Math.hypot(x - bx, y - by) <= badgeR + N * 0.06;
      if (!inBadgeHalo) {
        const onRing = dc <= R && dc >= R - ring;
        const onHands =
          segDist(x, y, c, c, c, c - R * 0.55) <= hand / 2 ||
          segDist(x, y, c, c, c + R * 0.42, c) <= hand / 2;
        if (onRing || onHands) col = BLUE;
      }

      const db = Math.hypot(x - bx, y - by);
      if (db <= badgeR) {
        col = AMBER;
        const barW = badgeR * 0.32, barH = badgeR * 1.0, gap = badgeR * 0.28;
        const inBar = (cx) => Math.abs(x - cx) <= barW / 2 && Math.abs(y - by) <= barH / 2;
        if (inBar(bx - gap / 2 - barW / 2) || inBar(bx + gap / 2 + barW / 2)) col = [255, 255, 255];
      }

      if (col) {
        const i = (Math.floor(sy / SS) * size + Math.floor(sx / SS)) * 4;
        px[i] += col[0]; px[i + 1] += col[1]; px[i + 2] += col[2]; px[i + 3] += 1;
      }
    }
  }

  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0;
    for (let x = 0; x < size; x++) {
      const i = (y * size + x) * 4, o = y * (size * 4 + 1) + 1 + x * 4;
      const n = px[i + 3];
      raw[o] = n ? Math.round(px[i] / n) : 0;
      raw[o + 1] = n ? Math.round(px[i + 1] / n) : 0;
      raw[o + 2] = n ? Math.round(px[i + 2] / n) : 0;
      raw[o + 3] = Math.round((n / (SS * SS)) * 255);
    }
  }
  return png(size, size, raw);
}

function png(w, h, raw) {
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc = (buf) => {
    let c = 0xffffffff;
    for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type), data]);
    const c = Buffer.alloc(4); c.writeUInt32BE(crc(td));
    return Buffer.concat([len, td, c]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8; ihdr[9] = 6; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0))
  ]);
}

const out = path.join(__dirname, "..", "assets");
fs.mkdirSync(out, { recursive: true });
for (const s of SIZES) fs.writeFileSync(path.join(out, `icon-${s}.png`), render(s));
console.log("Wrote", SIZES.map((s) => `icon-${s}.png`).join(", "));
