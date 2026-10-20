'use strict';
// Real-file test: opens every .docx / .pdf listed in a text file (one path per line), checks it
// renders without errors, then saves a copy to a temporary folder and reopens that copy.
// The original files are only read, never changed.
// Run: electron scripts/corpus.js <list.txt>
const { app, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const listFile = process.argv.find(a => a.endsWith('.txt'));
const files = fs.readFileSync(listFile, 'utf8').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'nos-corpus-'));
app.setPath('userData', path.join(tmp, 'profile'));
app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion');
let n = 0;
let lastSave = null;
dialog.showSaveDialog = async (w, o) => { o = o && !w?.webContents ? w : o || {}; const ext = ((o.filters || [])[0]?.extensions || ['bin'])[0]; lastSave = path.join(tmp, `copy-${++n}.${ext}`); return { canceled: false, filePath: lastSave }; };
dialog.showMessageBox = async (w, o) => ({ response: ((o && !w?.webContents ? w : o) || {}).cancelId ?? 0 });

const sleep = ms => new Promise(r => setTimeout(r, ms));
let errors = [];
async function js(win, code) { try { return await win.webContents.executeJavaScript(code, true); } catch (e) { errors.push(`exec: ${e.message}`); return null; } }
async function waitLoad(win) { if (win.webContents.isLoading()) await new Promise(r => win.webContents.once('did-finish-load', r)); await sleep(1500); }

require('../main.js');
app.once('browser-window-created', async (e, win) => {
  win.webContents.on('console-message', ev => { if (ev.level === 'error') errors.push(ev.message.slice(0, 160)); });
  await waitLoad(win);
  const rows = [];
  const docx = files.filter(f => /\.docx$/i.test(f));
  const pdfs = files.filter(f => /\.pdf$/i.test(f));
  for (const f of docx) {
    errors = [];
    const t0 = Date.now();
    const r = await js(win, `(async () => {
      const d = await api.openPath(${JSON.stringify(f)});
      if (d.error) return { error: d.error };
      loadDoc(d);
      await new Promise(r => setTimeout(r, 900));
      const text = editor.innerText.replace(/\\s+/g, ' ').trim();
      await saveDoc(true);
      return { chars: text.length, pages: document.querySelectorAll('.sheet, .page-sheet').length, tables: editor.querySelectorAll('table').length, imgs: editor.querySelectorAll('img').length, text: text.slice(0, 4000) };
    })()`);
    let round = null;
    if (r && !r.error && lastSave && fs.existsSync(lastSave)) {
      round = await js(win, `(async () => { const d = await api.openPath(${JSON.stringify(lastSave)}); if (d.error) return { error: d.error }; loadDoc(d); await new Promise(r => setTimeout(r, 700)); return { text: editor.innerText.replace(/\\s+/g, ' ').trim().slice(0, 4000), imgs: editor.querySelectorAll('img').length }; })()`);
    }
    const same = r && round && !round.error ? similarity(r.text, round.text) : null;
    rows.push({ file: f, ms: Date.now() - t0, ...(r || { error: 'no result' }), text: undefined, roundTrip: round?.error || (same == null ? 'n/a' : `${Math.round(same * 100)}% text kept, imgs ${r.imgs}->${round.imgs}`), errors });
    lastSave = null;
  }
  // PDFs: switch this window to the PDF editor
  await js(win, `api.switchEditor('pdf', null)`);
  await waitLoad(win);
  for (const f of pdfs) {
    errors = [];
    const t0 = Date.now();
    const r = await js(win, `(async () => {
      const P = window.__pdf;
      const got = await api.readFileBytes(${JSON.stringify(f)});
      if (got.error) return { error: got.error };
      await P.openBytes(got.bytes, ${JSON.stringify(path.basename(f))}, null);
      await new Promise(r => setTimeout(r, 1200));
      if (!P.S.open) return { error: 'did not open' + (document.querySelector('dialog[open]') ? ': ' + document.querySelector('dialog[open]').textContent.trim().slice(0, 80) : '') };
      const pages = P.S.entries.length;
      const rendered = document.querySelectorAll('.pg canvas').length;
      let saved = 'n/a';
      try {
        const out = await P.buildPdf();
        const doc = await P.pdfjsLib.getDocument({ data: out.slice(), ...P.PDF_OPTS }).promise;
        saved = doc.numPages === pages ? 'ok' : 'page count ' + doc.numPages + ' != ' + pages;
      } catch (e) { saved = 'save failed: ' + e.message; }
      document.querySelectorAll('dialog[open]').forEach(d => d.close());
      return { pages, rendered, saved };
    })()`);
    rows.push({ file: f, ms: Date.now() - t0, ...(r || { error: 'no result' }), errors });
  }
  let bad = 0;
  for (const r of rows) {
    const problem = r.error || r.errors.length || (r.saved && r.saved !== 'ok') || (r.roundTrip && /^\d+%/.test(r.roundTrip) && parseInt(r.roundTrip, 10) < 95);
    if (problem) bad++;
    console.log(`${problem ? 'FAIL' : 'PASS'}  ${path.basename(r.file)}  ${r.ms}ms  ${r.error ? 'ERROR ' + r.error : r.saved != null ? `pages ${r.pages} rendered ${r.rendered} save ${r.saved}` : `chars ${r.chars} tables ${r.tables} imgs ${r.imgs} | ${r.roundTrip}`}${r.errors.length ? '\n      ' + [...new Set(r.errors)].slice(0, 4).join('\n      ') : ''}`);
  }
  fs.writeFileSync(path.join(tmp, 'corpus-results.json'), JSON.stringify(rows, null, 1));
  console.log(`\n${rows.length - bad}/${rows.length} files OK. Details: ${path.join(tmp, 'corpus-results.json')}`);
  app.exit(0);
});

// share of words of a that also appear (in order-insensitive bag) in b
function similarity(a, b) {
  const bag = new Map();
  for (const w of b.split(' ')) bag.set(w, (bag.get(w) || 0) + 1);
  let hit = 0; const words = a.split(' ');
  for (const w of words) { const c = bag.get(w); if (c) { hit++; bag.set(w, c - 1); } }
  return words.length ? hit / words.length : 1;
}
