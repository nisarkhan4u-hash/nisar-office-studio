'use strict';
// Draws the Nisar Office Studio icon (same design as scripts/make-icon.js: blue rounded tile,
// white page with blue text lines, red badge with three white bars) and writes the Microsoft
// Store (MSIX/APPX) visual assets into build/appx/, which electron-builder picks up for
// `electron-builder --win appx`. No image libraries: plain Node (zlib for PNG compression).
//
//   node scripts/make-store-assets.js
//
// The renderer is exported too (mobile/scripts/make-mobile-icons.mjs uses it for the
// Android/iOS launcher icons and the Play Store graphics).
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const STORE_BG = [0x1f, 0x5f, 0xbf]; // #1f5fbf, the appx backgroundColor in package.json

function inRoundRect(x, y, x0, y0, x1, y1, r) {
  if (x < x0 || x > x1 || y < y0 || y > y1) return false;
  const cx = Math.min(Math.max(x, x0 + r), x1 - r);
  const cy = Math.min(Math.max(y, y0 + r), y1 - r);
  return (x - cx) ** 2 + (y - cy) ** 2 <= r * r;
}
// app colours from make-icon.js: gradient (27,78,170) -> (47,118,230)
const tileBg = (u, v) => { const t = (u + v) / 2; return [Math.round(27 + 20 * t), Math.round(78 + 40 * (1 - t)), Math.round(170 + 60 * (1 - t))]; };
const LINE = [31, 95, 191];
const FLAP = [200, 216, 240];

// colour at normalised (u, v) in [0,1] of the icon, or null for transparent.
// opts.tile: 'round' (desktop/Store tile), 'square' (full bleed: iOS/Play mask it themselves), 'none' (page + badge only)
function shade(u, v, opts) {
  const tile = opts.tile || 'round';
  const back = () => (tile === 'none' ? null : tileBg(u, v));
  if (tile === 'round' && !inRoundRect(u, v, 0.04, 0.04, 0.96, 0.96, 0.2)) return null;
  if (inRoundRect(u, v, 0.52, 0.6, 0.9, 0.88, 0.06)) {
    if (inRoundRect(u, v, 0.505, 0.585, 0.915, 0.895, 0.07) && !inRoundRect(u, v, 0.535, 0.615, 0.885, 0.865, 0.05)) return [255, 255, 255];
    for (const y of [0.66, 0.72, 0.78]) if (v >= y && v <= y + 0.03 && u >= 0.6 && u <= (y === 0.78 ? 0.74 : 0.82)) return [255, 255, 255];
    return [212, 47, 72];
  }
  const px0 = 0.26, px1 = 0.74, py0 = 0.16, py1 = 0.84, fold = 0.14;
  if (u >= px0 && u <= px1 && v >= py0 && v <= py1) {
    const fx = px1 - fold, fy = py0 + fold;
    if (u > fx && v < fy) {
      if ((u - fx) > (v - py0)) return back();
      return FLAP;
    }
    const lines = [0.40, 0.49, 0.58, 0.67];
    for (let i = 0; i < lines.length; i++) {
      const end = i === lines.length - 1 ? 0.56 : 0.66;
      if (v >= lines[i] && v <= lines[i] + 0.04 && u >= 0.34 && u <= end) return LINE;
    }
    return [255, 255, 255];
  }
  return back();
}

// RGBA canvas W x H; background colour (or null = transparent); the icon is drawn centred,
// iconSize px square. opts.tile as above.
function render(W, H, { background = null, iconSize = Math.min(W, H), tile = 'round', ss = 4 } = {}) {
  const px = Buffer.alloc(W * H * 4);
  const ox = (W - iconSize) / 2, oy = (H - iconSize) / 2;
  for (let y = 0; y < H; y++) {
    for (let x = 0; x < W; x++) {
      let r = 0, g = 0, b = 0, a = 0;
      for (let sy = 0; sy < ss; sy++) {
        for (let sx = 0; sx < ss; sx++) {
          const u = (x + (sx + 0.5) / ss - ox) / iconSize, v = (y + (sy + 0.5) / ss - oy) / iconSize;
          let c = u >= 0 && u <= 1 && v >= 0 && v <= 1 ? shade(u, v, { tile }) : null;
          if (!c && background) c = background;
          if (c) { r += c[0]; g += c[1]; b += c[2]; a++; }
        }
      }
      const i = (y * W + x) * 4;
      if (a) { px[i] = Math.round(r / a); px[i + 1] = Math.round(g / a); px[i + 2] = Math.round(b / a); }
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
// opaque: write RGB without alpha (App Store icons must not have an alpha channel)
function encodePng(W, H, px, { opaque = false } = {}) {
  const bpp = opaque ? 3 : 4;
  const raw = Buffer.alloc(H * (W * bpp + 1));
  for (let y = 0; y < H; y++) {
    const row = y * (W * bpp + 1);
    for (let x = 0; x < W; x++) {
      const s = (y * W + x) * 4, d = row + 1 + x * bpp;
      if (opaque) {
        const a = px[s + 3] / 255; // composite on white
        raw[d] = Math.round(px[s] * a + 255 * (1 - a)); raw[d + 1] = Math.round(px[s + 1] * a + 255 * (1 - a)); raw[d + 2] = Math.round(px[s + 2] * a + 255 * (1 - a));
      } else px.copy(raw, d, s, s + 4);
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(W, 0); ihdr.writeUInt32BE(H, 4);
  ihdr[8] = 8; ihdr[9] = opaque ? 2 : 6; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw, { level: 9 })), chunk('IEND', Buffer.alloc(0))]);
}
function makePng(W, H, opts = {}) { return encodePng(W, H, render(W, H, opts), opts); }

// Microsoft Store assets (names are the ones electron-builder's appx target looks for)
function storeAssets() {
  const icon = (s, pad = 0) => ({ w: s, h: s, opts: { iconSize: Math.round(s * (1 - pad)) } });
  const onBlue = (w, h, frac) => ({ w, h, opts: { background: STORE_BG, iconSize: Math.round(Math.min(w, h) * frac) } });
  return {
    'StoreLogo.png': icon(50),
    'Square44x44Logo.png': icon(44),
    'Square44x44Logo.scale-200.png': icon(88),
    'Square44x44Logo.targetsize-256_altform-unplated.png': icon(256),
    'Square150x150Logo.png': icon(150, 0.2),
    'Square150x150Logo.scale-200.png': icon(300, 0.2),
    'SmallTile.png': icon(71, 0.1),
    'LargeTile.png': icon(310, 0.3),
    'Wide310x150Logo.png': onBlue(310, 150, 0.7),
    'SplashScreen.png': onBlue(620, 300, 0.6),
  };
}

function main() {
  const out = path.join(__dirname, '..', 'build', 'appx');
  fs.mkdirSync(out, { recursive: true });
  for (const [name, a] of Object.entries(storeAssets())) {
    const buf = makePng(a.w, a.h, a.opts);
    fs.writeFileSync(path.join(out, name), buf);
    console.log(`${name.padEnd(54)} ${a.w}x${a.h}  ${buf.length} bytes`);
  }
  console.log('Store assets written to', out);
}

module.exports = { render, encodePng, makePng, STORE_BG };
if (require.main === module) main();
