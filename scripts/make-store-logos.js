'use strict';
// Store listing logos for Partner Center: 1:1 box art (1080 x 1080) and 9:16 poster art (720 x 1080).
// Run: node scripts/make-store-logos.js  -> build/store/*.png
const { createCanvas, loadImage, GlobalFonts } = require('@napi-rs/canvas');
const fs = require('fs');
const path = require('path');

const out = path.join(__dirname, '..', 'build', 'store');
fs.mkdirSync(out, { recursive: true });
const font = ['C:/Windows/Fonts/segoeuib.ttf', 'C:/Windows/Fonts/segoeui.ttf'].find(f => fs.existsSync(f));
if (font) GlobalFonts.registerFromPath(font, 'StoreFont');

async function make(file, w, h, iconSize, withName) {
  const c = createCanvas(w, h);
  const g = c.getContext('2d');
  const grad = g.createLinearGradient(0, 0, w, h);
  grad.addColorStop(0, '#2a6fd6');
  grad.addColorStop(1, '#17489a');
  g.fillStyle = grad;
  g.fillRect(0, 0, w, h);
  const icon = await loadImage(path.join(__dirname, '..', 'build', 'icon.png'));
  const y = withName ? h * 0.36 - iconSize / 2 : (h - iconSize) / 2;
  g.drawImage(icon, (w - iconSize) / 2, y, iconSize, iconSize);
  if (withName) {
    g.fillStyle = '#ffffff';
    g.textAlign = 'center';
    g.font = `${Math.round(w * 0.085)}px StoreFont`;
    g.fillText('Nisar Office', w / 2, h * 0.66);
    g.fillText('Studio', w / 2, h * 0.66 + w * 0.11);
    g.font = `${Math.round(w * 0.04)}px StoreFont`;
    g.fillStyle = 'rgba(255,255,255,0.85)';
    g.fillText('Documents & PDFs, offline', w / 2, h * 0.86);
  }
  fs.writeFileSync(path.join(out, file), c.toBuffer('image/png'));
  console.log(file, w, h);
}

(async () => {
  await make('box-art-1080.png', 1080, 1080, 640, false);
  await make('poster-art-720x1080.png', 720, 1080, 380, true);
})();
