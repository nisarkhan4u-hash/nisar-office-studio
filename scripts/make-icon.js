'use strict';
// Draws the Nisar Office Studio icon (blue tile, white page, text lines) and writes
// build/icon.png (512px) and build/icon.ico (16–256px) without any image libraries.
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

function inRoundRect(x, y, x0, y0, x1, y1, r) {
  if (x < x0 || x > x1 || y < y0 || y > y1) return false;
  const cx = Math.min(Math.max(x, x0 + r), x1 - r);
  const cy = Math.min(Math.max(y, y0 + r), y1 - r);
  return (x - cx) ** 2 + (y - cy) ** 2 <= r * r;
}

// app icon: blue, with a small red PDF badge; PDF file icon: red
const THEMES = {
  app: { bg: (u, v) => { const t = (u + v) / 2; return [Math.round(27 + 20 * t), Math.round(78 + 40 * (1 - t)), Math.round(170 + 60 * (1 - t))]; }, line: [31, 95, 191], flap: [200, 216, 240], badge: true },
  pdf: { bg: (u, v) => { const t = (u + v) / 2; return [Math.round(200 + 20 * (1 - t)), Math.round(40 + 20 * (1 - t)), Math.round(60 + 15 * (1 - t))]; }, line: [196, 49, 75], flap: [244, 205, 212], badge: false },
};
let theme = THEMES.app;
const bg = (u, v) => theme.bg(u, v);

// colour of the icon at normalised point (u, v) in [0,1], or null for transparent
function shade(u, v) {
  if (!inRoundRect(u, v, 0.04, 0.04, 0.96, 0.96, 0.2)) return null;
  if (theme.badge && inRoundRect(u, v, 0.52, 0.6, 0.9, 0.88, 0.06)) {
    // red badge with three white "text" bars
    if (inRoundRect(u, v, 0.505, 0.585, 0.915, 0.895, 0.07) && !inRoundRect(u, v, 0.535, 0.615, 0.885, 0.865, 0.05)) return [255, 255, 255];
    for (const y of [0.66, 0.72, 0.78]) if (v >= y && v <= y + 0.03 && u >= 0.6 && u <= (y === 0.78 ? 0.74 : 0.82)) return [255, 255, 255];
    return [212, 47, 72];
  }
  // page with a folded top-right corner
  const px0 = 0.26, px1 = 0.74, py0 = 0.16, py1 = 0.84, fold = 0.14;
  const onPage = u >= px0 && u <= px1 && v >= py0 && v <= py1;
  if (onPage) {
    const fx = px1 - fold, fy = py0 + fold;
    if (u > fx && v < fy) {
      if ((u - fx) > (v - py0)) return bg(u, v); // cut corner
      return theme.flap;                         // folded flap
    }
    const lines = [0.40, 0.49, 0.58, 0.67];
    for (let i = 0; i < lines.length; i++) {
      const end = i === lines.length - 1 ? 0.56 : 0.66;
      if (v >= lines[i] && v <= lines[i] + 0.04 && u >= 0.34 && u <= end) return theme.line;
    }
    return [255, 255, 255];
  }
  return bg(u, v);
}

function render(size) {
  const ss = 4;
  const px = Buffer.alloc(size * size * 4);
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let r = 0, g = 0, b = 0, a = 0;
      for (let sy = 0; sy < ss; sy++) {
        for (let sx = 0; sx < ss; sx++) {
          const c = shade((x + (sx + 0.5) / ss) / size, (y + (sy + 0.5) / ss) / size);
          if (c) { r += c[0]; g += c[1]; b += c[2]; a++; }
        }
      }
      const i = (y * size + x) * 4;
      if (a) { px[i] = r / a; px[i + 1] = g / a; px[i + 2] = b / a; }
      px[i + 3] = Math.round((a / (ss * ss)) * 255);
    }
  }
  return px;
}

const CRC = new Int32Array(256).map((_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c; });
function crc32(buf) { let c = -1; for (const b of buf) c = CRC[(c ^ b) & 255] ^ (c >>> 8); return (c ^ -1) >>> 0; }
function chunk(type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}
function png(size) {
  const px = render(size);
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) px.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 6; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw, { level: 9 })), chunk('IEND', Buffer.alloc(0))]);
}
function ico(sizes) {
  const images = sizes.map(png);
  const header = Buffer.alloc(6 + 16 * sizes.length);
  header.writeUInt16LE(0, 0); header.writeUInt16LE(1, 2); header.writeUInt16LE(sizes.length, 4);
  let offset = header.length;
  sizes.forEach((s, i) => {
    const e = 6 + i * 16;
    header[e] = s >= 256 ? 0 : s; header[e + 1] = s >= 256 ? 0 : s;
    header[e + 2] = 0; header[e + 3] = 0;
    header.writeUInt16LE(1, e + 4); header.writeUInt16LE(32, e + 6);
    header.writeUInt32LE(images[i].length, e + 8); header.writeUInt32LE(offset, e + 12);
    offset += images[i].length;
  });
  return Buffer.concat([header, ...images]);
}

const out = path.join(__dirname, '..', 'build');
fs.mkdirSync(out, { recursive: true });
const SIZES = [16, 24, 32, 48, 64, 128, 256];
theme = THEMES.app;
fs.writeFileSync(path.join(out, 'icon.png'), png(512));
fs.writeFileSync(path.join(out, 'icon.ico'), ico(SIZES));
theme = { ...THEMES.app, badge: false }; // .docx files: plain blue page
fs.writeFileSync(path.join(out, 'docx.ico'), ico(SIZES));
theme = THEMES.pdf;                      // .pdf files: red page
fs.writeFileSync(path.join(out, 'pdf.ico'), ico(SIZES));
fs.writeFileSync(path.join(out, 'pdf-icon.png'), png(256));
console.log('Icons written to', out);
