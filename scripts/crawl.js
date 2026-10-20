'use strict';
// "Click everything" test: presses every ribbon button, every menu item, every backstage pane,
// every PDF tool and the main keyboard shortcuts in both editors, answers every dialog with its
// default (OK) and records errors, failed toasts and crashes per control.
// Native file dialogs are answered with test files, printing is stubbed.
// Run: electron scripts/crawl.js            (NW_TRACE=1 prints every step)
const { app, dialog, shell, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'nos-crawl-'));
app.setPath('userData', path.join(tmp, 'profile'));
app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion');
const OUT = path.join(tmp, 'out');
fs.mkdirSync(OUT);
const FIX = { png: path.join(__dirname, '..', 'build', 'icon.png'), pdf: path.join(tmp, 'sample.pdf'), docx: path.join(tmp, 'sample.docx') };

// ---- fixtures
async function makeFixtures() {
  const { PDFDocument, StandardFonts } = require('pdf-lib');
  const d = await PDFDocument.create();
  const f = await d.embedFont(StandardFonts.Helvetica);
  for (let i = 1; i <= 3; i++) {
    const p = d.addPage([612, 792]);
    p.drawText(`Sample heading ${i}`, { x: 60, y: 700, size: 24, font: f });
    p.drawText('The quick brown fox jumps over the lazy dog. Secret 1001.', { x: 60, y: 650, size: 12, font: f });
  }
  const form = d.getForm();
  form.createTextField('name').addToPage(d.getPage(0), { x: 60, y: 560, width: 200, height: 22 });
  form.createCheckBox('agree').addToPage(d.getPage(0), { x: 60, y: 520, width: 16, height: 16 });
  d.addOutline?.();
  fs.writeFileSync(FIX.pdf, await d.save());
  const JSZip = require('jszip');
  const z = new JSZip();
  z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Imported heading</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Imported </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve"> text.</w:t></w:r></w:p></w:body></w:document>');
  fs.writeFileSync(FIX.docx, await z.generateAsync({ type: 'nodebuffer' }));
}

// ---- native dialogs answered automatically
let saveN = 0;
const pickFor = opts => {
  const props = opts.properties || [];
  if (props.includes('openDirectory')) { const dir = path.join(OUT, `folder-${++saveN}`); fs.mkdirSync(dir, { recursive: true }); return [dir]; }
  const exts = (opts.filters || []).flatMap(f => f.extensions || []).map(e => e.toLowerCase());
  const multi = props.includes('multiSelections');
  if (exts.includes('pdf') && !exts.includes('docx')) return multi ? [FIX.pdf, FIX.pdf] : [FIX.pdf];
  if (exts.some(e => ['png', 'jpg', 'jpeg'].includes(e))) return multi ? [FIX.png, FIX.png] : [FIX.png];
  if (exts.includes('docx')) return [FIX.docx];
  return [FIX.pdf];
};
dialog.showOpenDialog = async (w, o) => { o = o && !(w instanceof BrowserWindow) && !w?.webContents ? w : o || {}; const p = pickFor(o); step.dialogs.push(`open:${path.basename(p[0])}`); return { canceled: false, filePaths: p }; };
dialog.showOpenDialogSync = (w, o) => pickFor(o || w || {});
dialog.showSaveDialog = async (w, o) => {
  o = o && !w?.webContents ? w : o || {};
  const ext = ((o.filters || [])[0]?.extensions || [path.extname(o.defaultPath || '.bin').slice(1)])[0] || 'bin';
  const p = path.join(OUT, `save-${++saveN}.${ext}`);
  step.dialogs.push(`save:${path.basename(p)}`);
  return { canceled: false, filePath: p };
};
dialog.showSaveDialogSync = () => path.join(OUT, `save-${++saveN}.bin`);
const answer = o => { const b = o.buttons || ['OK']; const i = o.cancelId ?? b.length - 1; step.dialogs.push(`msg:${(o.message || '').slice(0, 50)}`); return i; };
dialog.showMessageBox = async (w, o) => ({ response: answer(o && !w?.webContents ? w : o || {}), checkboxChecked: false });
dialog.showMessageBoxSync = (w, o) => answer(o || w || {});
dialog.showErrorBox = (t, c) => step.errors.push(`errorBox: ${t} ${c}`);
shell.openExternal = async () => {};
shell.openPath = async () => '';
shell.showItemInFolder = () => {};
app.on('web-contents-created', (e, wc) => {
  wc.print = (o, cb) => { step.dialogs.push('print'); (typeof o === 'function' ? o : cb)?.(true, ''); };
  wc.on('render-process-gone', (ev, d) => step.errors.push(`renderer gone: ${d.reason}`));
});

// ---- results
let step = { name: 'startup', errors: [], dialogs: [] };
const results = [];
const trace = (...a) => { if (process.env.NW_TRACE) console.log(...a); };
let snapWin = null;
async function begin(name) { step = { name, errors: [], dialogs: [], modals: [], toasts: [] }; trace('>', name); }
function end(extra = {}) { Object.assign(step, extra); results.push(step); if (step.errors.length) trace('  ERR', step.errors.join(' | ')); }
const SNAP = `(() => { const e = document.getElementById('editor'); const P = window.__pdf; return JSON.stringify([e ? e.innerHTML : '', document.body.className, P ? P.S.entries.map(x => [x.id, x.rotation, JSON.stringify(x.objects, (k, v) => (k === 'src' ? 0 : v)), JSON.stringify(x.crop || 0), (x.ocr || []).length]) : 0, P ? [P.S.zoom, P.S.tool, JSON.stringify(P.S.style), P.S.protect ? 1 : 0, P.S.sanitize ? 1 : 0, P.S.flattenForms ? 1 : 0] : 0, (() => { try { return typeof state !== 'undefined' ? JSON.stringify([state.page, state.zoom, state.trackChanges, state.markup, state.watermark]) : ''; } catch { return ''; } })(), (document.querySelector('#zoom-label, .zoom-label') || {}).textContent || '', [...document.querySelectorAll('.sheet')].map(x => x.getAttribute('style')).join('|')]); })()`;
const sleep = ms => new Promise(r => setTimeout(r, ms));

const HELPER = `(() => {
  if (window.__crawl) return;
  window.__crawl = { toasts: [] };
  const tt = document.getElementById('toast'); if (tt) new MutationObserver(() => { const t = tt.textContent.trim(); if (t) __crawl.toasts.push(t); }).observe(tt, { childList: true, subtree: true, characterData: true });
  HTMLInputElement.prototype.click = function () { if (this.type === 'file' || this.type === 'color') { this.dispatchEvent(new Event('crawl-picker')); return; } HTMLElement.prototype.click.call(this); };
  window.print = () => {};
  window.alert = () => {}; window.confirm = () => true; window.prompt = () => '';
  __crawl.settle = async (log) => {
    for (let i = 0; i < 6; i++) {
      await new Promise(r => setTimeout(r, 350));
      const d = [...document.querySelectorAll('dialog[open]')].pop();
      if (!d) return;
      const t = (d.querySelector('h2, h1, .title') || d).textContent.trim().slice(0, 60);
      log.push(t);
      const ok = d.querySelector('button[type=submit], .btn.primary, [data-ok], [data-yes]');
      if (ok && i < 4) ok.click(); else d.close();
    }
  };
})()`;

async function js(win, code) {
  try { return await win.webContents.executeJavaScript(code, true); } catch (err) { step.errors.push(`exec: ${err.message}`); return null; }
}

// ---- document editor
const DOC_SAMPLE = `<h1>Report title</h1><p id="p1">The quick brown fox jumps over the lazy dog with a <a href="https://example.com">link</a>.</p><h2>Section</h2><ul><li>First item</li><li>Second item</li></ul><table><tbody><tr><td id="c1">A1</td><td>B1</td></tr><tr><td>A2</td><td>B2</td></tr></tbody></table><p>Last paragraph.</p>`;
const resetDoc = (where) => `(async () => {
  document.querySelectorAll('dialog[open]').forEach(d => d.close());
  try { hideBackstage(); } catch {}
  try { closeMenu(); } catch {}
  editor.innerHTML = ${JSON.stringify(DOC_SAMPLE)};
  markDirty();
  const r = document.createRange();
  if (${JSON.stringify(where)} === 'table') { const c = document.getElementById('c1'); r.setStart(c.firstChild, 1); r.collapse(true); }
  else { const t = document.getElementById('p1').firstChild; r.setStart(t, 4); r.setEnd(t, 15); }
  editor.focus(); getSelection().removeAllRanges(); getSelection().addRange(r);
  try { saveSelection && saveSelection(); } catch {}
  __crawl.toasts.length = 0;
  return 1;
})()`;

async function crawlEditor(win, kind) {
  await js(win, HELPER);
  const controls = await js(win, `(() => {
    const out = [];
    const tabOf = el => { const p = el.closest('[data-panel]'); return p ? p.dataset.panel : null; };
    for (const el of document.querySelectorAll('[data-cmd]')) {
      if (el.closest('.menu')) continue;
      out.push({ kind: 'cmd', key: el.dataset.cmd + (el.dataset.arg ? ':' + el.dataset.arg : ''), sel: el.dataset.arg ? '[data-cmd="' + el.dataset.cmd + '"][data-arg="' + el.dataset.arg + '"]' : '[data-cmd="' + el.dataset.cmd + '"]', tab: tabOf(el) });
    }
    for (const el of document.querySelectorAll('[data-menu]')) out.push({ kind: 'menu', key: el.dataset.menu, sel: '[data-menu="' + el.dataset.menu + '"]', tab: tabOf(el) });
    for (const el of document.querySelectorAll('[data-pane]')) out.push({ kind: 'pane', key: el.dataset.pane, sel: '[data-pane="' + el.dataset.pane + '"]' });
    for (const el of document.querySelectorAll('[data-tool]')) out.push({ kind: 'tool', key: el.dataset.tool, sel: '[data-tool="' + el.dataset.tool + '"]', tab: tabOf(el) });
    const seen = new Set();
    return out.filter(c => { const k = c.kind + c.sel; if (seen.has(k)) return false; seen.add(k); return true; });
  })()`);
  const SKIP = new Set(['closeDoc']); // closes the window: tested at the very end
  trace(`${kind}: ${controls.length} controls`);
  const TABLE = /^(rowAbove|rowBelow|colLeft|colRight|delRow|delCol|delTable|headerRow|table.*)$/i;
  const reset = async where => { await (kind === 'doc' ? js(win, resetDoc(where)) : js(win, resetPdf())); step.before = await js(win, SNAP); };
  const clickSel = c => `(async () => {
    const el = document.querySelector(${JSON.stringify(c.sel)});
    if (!el) return 'missing';
    ${c.tab ? `const tab = document.querySelector('[data-tab="${c.tab}"]'); if (tab) tab.click();` : ''}
    ${c.kind === 'pane' ? `try { ${kind === 'doc' ? "showBackstage('home')" : "document.querySelector('[data-cmd=backstage]').click()"}; } catch {}` : ''}
    el.click();
    return 'ok';
  })()`;
  for (const c of controls) {
    if (SKIP.has(c.key)) continue;
    if (c.kind === 'menu') {
      await reset(TABLE.test(c.key) ? 'table' : 'text');
      await js(win, clickSel(c));
      await sleep(250);
      const items = await js(win, `[...document.querySelectorAll('#menu-layer .menu [data-i]')].map(b => b.textContent.trim().replace(/\\s+/g, ' ').slice(0, 40))`) || [];
      if (!items.length) {
        // palettes, the table grid and the symbol picker: click their first and a middle cell
        const cells = await js(win, `(() => { const m = document.querySelector('#menu-layer .menu'); return m ? m.querySelectorAll('button, td, [data-color], [data-sym], .cell, .sw').length : -1; })()`);
        await js(win, `closeMenu && closeMenu()`);
        if (cells < 1) { begin(`${kind} menu ${c.key}`); step.errors.push('menu did not open or has no items'); end(); continue; }
        for (const i of [...new Set([0, Math.floor(cells / 2), cells - 1])]) {
          begin(`${kind} menu ${c.key} > cell ${i + 1}/${cells}`);
          await reset(TABLE.test(c.key) ? 'table' : 'text');
          await js(win, clickSel(c));
          await sleep(200);
          const what = await js(win, `(() => { const m = document.querySelector('#menu-layer .menu'); const el = m && m.querySelectorAll('button, td, [data-color], [data-sym], .cell, .sw')[${i}]; if (!el) return 'missing'; el.dispatchEvent(new MouseEvent('mousemove', { bubbles: true })); el.dispatchEvent(new MouseEvent('mouseover', { bubbles: true })); el.click(); return (el.title || el.textContent || el.dataset.color || '').trim().slice(0, 30); })()`);
          step.note = what;
          const before = await js(win, `typeof editor !== 'undefined' ? editor.innerHTML.length : 0`);
          await finishStep(win);
          if (c.key === 'table') { const after = await js(win, `editor.querySelectorAll('table').length`); step.note += ` tables now ${after}`; }
        }
        continue;
      }
      await js(win, `closeMenu && closeMenu()`);
      for (let i = 0; i < items.length; i++) {
        begin(`${kind} menu ${c.key} > ${items[i]}`);
        await reset(TABLE.test(c.key) ? 'table' : 'text');
        await js(win, clickSel(c));
        await sleep(200);
        await js(win, `(() => { const b = document.querySelectorAll('#menu-layer .menu [data-i]')[${i}]; if (b) b.click(); })()`);
        await finishStep(win);
      }
      continue;
    }
    begin(`${kind} ${c.kind} ${c.key}`);
    await reset(TABLE.test(c.key) ? 'table' : 'text');
    const r = await js(win, clickSel(c));
    if (r === 'missing') step.errors.push('element missing');
    if (c.kind === 'tool') await dragOnPage(win);
    await finishStep(win);
  }
}
async function finishStep(win) {
  await sleep(300);
  const modals = [];
  await js(win, `(async () => { const log = []; await __crawl.settle(log); return log; })()`).then(l => modals.push(...(l || [])));
  await sleep(400);
  const toasts = await js(win, `__crawl.toasts.splice(0)`) || [];
  for (const t of toasts) if (/went wrong|failed|error|could not|couldn't|cannot/i.test(t)) step.errors.push(`toast: ${t}`);
  const alive = await js(win, `!!document.body && document.querySelectorAll('.pg, #editor').length`);
  if (step.before != null) { const after = await js(win, SNAP); step.changed = after !== step.before; delete step.before; }
  if (!alive) step.errors.push('page empty after step');
  end({ modals, toasts });
}

// ---- PDF editor
const resetPdf = () => `(async () => {
  document.querySelectorAll('dialog[open]').forEach(d => d.close());
  const bs = document.querySelector('#backstage'); if (bs) bs.hidden = true;
  const P = window.__pdf;
  if (!window.__crawlBytes) window.__crawlBytes = (await api.readFileBytes(${JSON.stringify(FIX.pdf)})).bytes;
  if (P.S.editing) P.finishEditing();
  await P.openBytes(window.__crawlBytes.slice(), 'sample.pdf', null);
  P.setTool('select');
  const v = P.views.get(P.S.entries[0].id);
  const o = P.createObject(v, 'text', 60, 300, 200, 30, { text: 'Crawl text', size: 14 });
  P.select(P.S.entries[0].id, o.id);
  __crawl.toasts.length = 0;
  return 1;
})()`;
async function dragOnPage(win) {
  const r = await js(win, `(() => { const c = document.querySelector('.pg'); if (!c) return null; c.scrollIntoView({ block: 'start' }); const b = c.getBoundingClientRect(); return [b.left + b.width * 0.3, b.top + Math.min(b.height * 0.5, 400)]; })()`);
  if (!r) return;
  const [x, y] = r.map(Math.round);
  const before = await js(win, `window.__pdf ? window.__pdf.S.entries.reduce((n, e) => n + e.objects.length, 0) : 0`);
  const wc = win.webContents;
  wc.sendInputEvent({ type: 'mouseDown', x, y, button: 'left', clickCount: 1 });
  for (let k = 1; k <= 6; k++) { wc.sendInputEvent({ type: 'mouseMove', x: x + k * 25, y: y + k * 12, button: 'left', modifiers: ['leftButtonDown'] }); await sleep(30); }
  wc.sendInputEvent({ type: 'mouseUp', x: x + 150, y: y + 72, button: 'left', clickCount: 1 });
  await sleep(500);
  const after = await js(win, `window.__pdf ? window.__pdf.S.entries.reduce((n, e) => n + e.objects.length, 0) : 0`);
  step.note = `objects ${before} -> ${after}`;
}

// ---- shortcuts
async function keys(win, list, reset) {
  for (const k of list) {
    begin(`shortcut ${k}`);
    await reset();
    step.before = await js(win, SNAP);
    const [key, ...mods] = k.split('+').reverse();
    const modifiers = mods.map(m => m.toLowerCase());
    for (const type of ['keyDown', 'char', 'keyUp']) {
      if (type === 'char' && (key.length > 1 || modifiers.includes('control'))) continue;
      win.webContents.sendInputEvent({ type, keyCode: key, modifiers });
    }
    await finishStep(win);
  }
}

// ---- run
const extraWindows = new Set();
app.on('browser-window-created', (e, w) => {
  if (!mainWin) return;
  extraWindows.add(w);
  w.webContents.on('console-message', ev => { if (ev.level === 'error') step.errors.push(`[new window] ${ev.message}`); });
  w.webContents.once('did-finish-load', () => setTimeout(() => { step.dialogs.push(`window:${path.basename(w.webContents.getURL().split('?')[0])}`); if (!w.isDestroyed()) w.destroy(); }, 1200));
});
let mainWin = null;
async function waitLoad(win) {
  if (win.webContents.isLoading()) await new Promise(r => win.webContents.once('did-finish-load', r));
  await sleep(1500);
}
async function run() {
  await makeFixtures();
  require('../main.js');
  await new Promise(r => app.once('browser-window-created', (e, w) => { mainWin = w; r(); }));
  const win = mainWin;
  win.webContents.on('console-message', ev => { if (ev.level === 'error') step.errors.push(`console: ${ev.message.slice(0, 200)}`); });
  win.webContents.on('preload-error', (e, p, err) => step.errors.push(`preload: ${err.message}`));
  await waitLoad(win);
  await js(win, HELPER);
  await crawlEditor(win, 'doc');
  await keys(win, ['Control+B', 'Control+I', 'Control+U', 'Control+Z', 'Control+Y', 'Control+F', 'Control+H', 'Control+D', 'Control+K', 'Control+S', 'Control+P', 'Control+Enter', 'Control+Alt+M', 'Control+Alt+F', 'Control+Shift+E', 'Control+A', 'Tab', 'Escape', 'F7'], () => js(win, resetDoc('text')));
  // switch to the PDF editor in the same window
  begin('switch to PDF editor');
  await js(win, `api.switchEditor('pdf', ${JSON.stringify(FIX.pdf)})`);
  await waitLoad(win);
  await sleep(2500);
  const opened = await js(win, `!!(window.__pdf && window.__pdf.S.open)`);
  if (!opened) step.errors.push('PDF did not open');
  end();
  await js(win, HELPER);
  await crawlEditor(win, 'pdf');
  await keys(win, ['Control+Z', 'Control+Y', 'Delete', 'Control+F', 'Control+S', 'Control+P', 'Control+C', 'Control+V', 'Control+Plus', 'Control+-', 'Escape', 'Right', 'PageDown'], () => js(win, resetPdf()));
  const bad = results.filter(r => r.errors.length);
  for (const r of results) {
    if (process.env.NW_TRACE || r.errors.length) console.log(`${r.errors.length ? 'FAIL' : 'PASS'}${r.changed === false ? ' (no change)' : ''}  ${r.name}${r.note ? `  (${r.note})` : ''}${r.modals?.length ? `  [dialogs: ${r.modals.join(' / ')}]` : ''}${r.dialogs?.length ? `  [native: ${r.dialogs.join(', ')}]` : ''}${r.errors.length ? '\n      ' + r.errors.join('\n      ') : ''}`);
  }
  fs.writeFileSync(path.join(tmp, 'crawl-results.json'), JSON.stringify(results, null, 1));
  console.log(`\n${results.length - bad.length}/${results.length} steps without errors. Saved files: ${fs.readdirSync(OUT).length}. Details: ${path.join(tmp, 'crawl-results.json')}`);
  // last: File > Close must close the window
  const closed = new Promise(r => win.once('closed', () => r(true)));
  await js(win, `document.querySelector('[data-cmd=closeDoc]').click()`);
  const ok = await Promise.race([closed, sleep(4000).then(() => false)]);
  console.log(`${ok ? 'PASS' : 'FAIL'}  pdf closeDoc closes the window`);
  app.exit(bad.length || !ok ? 1 : 0);
}
app.whenReady().then(() => {}); // main.js handles ready
run().catch(err => { console.error('CRAWL CRASHED', err); app.exit(2); });
