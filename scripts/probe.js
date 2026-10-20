'use strict';
// Developer probe: runs each snippet from a JS file (separated by lines "//---") in the
// editor window and prints the results. Run: electron scripts/probe.js <snippets.js>
const { app } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const file = process.argv.find(a => a.endsWith('.js') && !a.endsWith('probe.js'));
const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'nw-probe-'));
app.setPath('userData', profile);
if (process.env.NW_SEED_RECOVERY) {
  // pretend a previous session crashed with unsaved work
  fs.mkdirSync(path.join(profile, 'recovery'), { recursive: true });
  fs.writeFileSync(path.join(profile, 'recovery', `${Date.now() - 60000}-1.json`), JSON.stringify({ html: '<h1>Recovered work</h1><p>This text was not saved before the crash.</p>', page: null, name: 'Essay draft.docx', path: null, time: Date.now() - 60000 }));
}
app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion');
require('../main.js');

app.on('browser-window-created', (e, win) => {
  if (win.__p) return;
  win.__p = true;
  win.webContents.on('console-message', ev => { if (ev.level === 'error') console.log('CONSOLE ERROR:', ev.message); });
  win.webContents.once('did-finish-load', async () => {
    await new Promise(r => setTimeout(r, 1200));
    const snippets = fs.readFileSync(file, 'utf8').split(/^\/\/---.*$/m).map(s => s.trim()).filter(Boolean);
    for (const s of snippets) {
      if (s.startsWith('SHOT ')) {
        await new Promise(r => setTimeout(r, 400));
        win.webContents.invalidate();
        await new Promise(r => setTimeout(r, 150));
        fs.writeFileSync(s.slice(5).trim(), (await win.webContents.capturePage()).toPNG());
        continue;
      }
      if (s.startsWith('SIZE ')) { const [w, h] = s.slice(5).split('x').map(Number); win.unmaximize(); win.setBounds({ x: 30, y: 30, width: w, height: h }); await new Promise(r => setTimeout(r, 500)); continue; }
      if (s.startsWith('TYPE ')) { win.webContents.insertText(s.slice(5)); await new Promise(r => setTimeout(r, 200)); continue; }
      if (s.startsWith('KEY ')) {
        const [key, ...mods] = s.slice(4).split('+').reverse();
        for (const type of ['keyDown', 'char', 'keyUp']) {
          if (type === 'char' && key.length > 1 && key !== 'Enter') continue;
          win.webContents.sendInputEvent({ type, keyCode: type === 'char' && key === 'Enter' ? '\r' : key, modifiers: mods });
        }
        await new Promise(r => setTimeout(r, 200));
        continue;
      }
      try { console.log('>', JSON.stringify(await win.webContents.executeJavaScript(s, true))); } catch (err) { console.log('! error', err.message); }
    }
    app.exit(0);
  });
});
