'use strict';
// End-to-end check of the PDF editor: npm run pdf-smoke  (electron scripts/pdf-smoke.js [--shots=<dir>])
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const L = require('pdf-lib');

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'nisarstudio-pdf-'));
const shots = process.argv.find(a => a.startsWith('--shots='))?.slice(8);
if (shots) fs.mkdirSync(shots, { recursive: true });
app.setPath('userData', path.join(tmp, 'profile'));
app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion');
require('../main.js');

const results = [];
const errors = [];
const ok = (name, cond, extra = '') => {
  const line = `${cond ? 'PASS' : 'FAIL'}  ${name}${extra && !cond ? `  (${extra})` : ''}`;
  results.push(line);
  if (process.env.NW_TRACE) console.log('TRACE', line);
};
const wait = ms => new Promise(r => setTimeout(r, ms));

async function fixtures() {
  const doc = await L.PDFDocument.create();
  const font = await doc.embedFont(L.StandardFonts.Helvetica);
  for (let i = 1; i <= 3; i++) {
    const p = doc.addPage([612, 792]);
    p.drawText(`Original page ${i}`, { x: 72, y: 700, size: 28, font });
    p.drawText('The quick brown fox jumps over the lazy dog.', { x: 72, y: 650, size: 14, font });
    p.drawText(`Secret code ${1000 + i}`, { x: 72, y: 620, size: 14, font });
  }
  const land = doc.addPage([612, 792]);
  land.setRotation(L.degrees(90));
  land.drawText('Rotated page', { x: 72, y: 700, size: 24, font });
  const form = doc.getForm();
  const tf = form.createTextField('full_name'); tf.addToPage(doc.getPage(0), { x: 72, y: 540, width: 220, height: 22 });
  const cb = form.createCheckBox('agree'); cb.addToPage(doc.getPage(0), { x: 72, y: 500, width: 16, height: 16 });
  const rg = form.createRadioGroup('size'); rg.addOptionToPage('S', doc.getPage(0), { x: 72, y: 460, width: 14, height: 14 }); rg.addOptionToPage('L', doc.getPage(0), { x: 110, y: 460, width: 14, height: 14 });
  const dd = form.createDropdown('country'); dd.addOptions(['Pakistan', 'India', 'UK']); dd.addToPage(doc.getPage(0), { x: 72, y: 420, width: 150, height: 22 });
  fs.writeFileSync(path.join(tmp, 'sample.pdf'), await doc.save());
  const other = await L.PDFDocument.create();
  const f2 = await other.embedFont(L.StandardFonts.Helvetica);
  for (let i = 1; i <= 2; i++) other.addPage([595, 842]).drawText(`Inserted ${i}`, { x: 72, y: 760, size: 24, font: f2 });
  fs.writeFileSync(path.join(tmp, 'other.pdf'), await other.save());
}

let started = false;
app.on('browser-window-created', (e, win) => {
  if (started) return;
  started = true;
  win.webContents.on('console-message', ev => { if (ev.level === 'error') errors.push(`${ev.message} @${ev.lineNumber}`); });
  win.webContents.once('did-finish-load', () => run(win).catch(err => { results.push(`FAIL  crashed: ${err.stack || err}`); finish(); }));
});

async function run(win) {
  const wc = win.webContents;
  await fixtures();
  win.unmaximize();
  win.setBounds({ x: 30, y: 30, width: 1500, height: 950 });
  await wait(800);
  // switch this (empty) document window to the PDF editor with the sample file
  const loaded = new Promise(r => wc.once('did-finish-load', r));
  await wc.executeJavaScript(`api.switchEditor('pdf', ${JSON.stringify(path.join(tmp, 'sample.pdf'))})`);
  await loaded;
  await wait(2500);
  const js = c => wc.executeJavaScript(c, true);
  const P = 'window.__pdf';
  let n = 0;
  const shot = async name => { if (!shots) return; await wait(400); wc.invalidate(); await wait(150); fs.writeFileSync(path.join(shots, `${String(++n).padStart(2, '0')}-${name}.png`), (await wc.capturePage()).toPNG()); };
  // click an element (scrolled into view); opts.at = [fx, fy] picks a point inside it (fractions)
  const click = async (selector, opts = {}) => {
    const at = opts.at || [0.5, 0.5];
    const r = await js(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return null; if (!${!!opts.noScroll}) el.scrollIntoView({ block: 'center', inline: 'center' }); const b = el.getBoundingClientRect(); const top = Math.max(b.top, 0), bottom = Math.min(b.bottom, innerHeight); return { x: b.left + b.width * ${at[0]}, y: ${opts.at ? `b.top + b.height * ${at[1]}` : '(top + bottom) / 2'} }; })()`);
    if (!r) throw new Error(`no element ${selector}`);
    await wait(150);
    const x = Math.round(r.x + (opts.dx || 0)), y = Math.round(r.y + (opts.dy || 0));
    wc.sendInputEvent({ type: 'mouseMove', x, y });
    wc.sendInputEvent({ type: 'mouseDown', x, y, button: 'left', clickCount: opts.count || 1 });
    wc.sendInputEvent({ type: 'mouseUp', x, y, button: 'left', clickCount: opts.count || 1 });
    await wait(250);
  };
  const drag = async (x0, y0, x1, y1) => {
    wc.sendInputEvent({ type: 'mouseMove', x: x0, y: y0 });
    wc.sendInputEvent({ type: 'mouseDown', x: x0, y: y0, button: 'left', clickCount: 1 });
    for (let i = 1; i <= 8; i++) { wc.sendInputEvent({ type: 'mouseMove', x: Math.round(x0 + ((x1 - x0) * i) / 8), y: Math.round(y0 + ((y1 - y0) * i) / 8), button: 'left', modifiers: ['leftbuttondown'] }); await wait(20); }
    wc.sendInputEvent({ type: 'mouseUp', x: x1, y: y1, button: 'left', clickCount: 1 });
    await wait(300);
  };

  ok('PDF editor opened', (await js('document.title')).includes('sample.pdf'), await js('document.title'));
  ok('4 pages', await js(`${P}.S.entries.length === 4`));
  ok('first page rendered', await js('!!document.querySelector(".pg canvas.render")'));
  ok('text layer built', (await js('document.querySelector(".pg .textLayer").innerText')).includes('quick brown fox'));
  ok('thumbnails', (await js('document.querySelectorAll(".pthumb").length')) === 4);
  ok('form fields shown', (await js('document.querySelectorAll(".fld").length')) >= 5);
  ok('rotated page geometry', await js(`(() => { const g = ${P}.geom(${P}.S.entries[3]); return g.rotation === 90 && g.w1 > g.h1; })()`));
  ok('viewport math matches pdf.js', await js(`(() => { const e = ${P}.S.entries[3]; const pg = ${P}.S.sources[0].pages[e.index]; const vp = pg.getViewport({ scale: 1.5, rotation: 90 }); const mine = ${P}.viewportMatrix(pg.view, 90, 1.5); return vp.transform.every((x, i) => Math.abs(x - mine[i]) < 1e-6); })()`));
  await shot('opened');

  // fill the form like a user
  await click('.fld[data-name="full_name"]');
  wc.insertText('Nisar Khan');
  await wait(200);
  await js('document.activeElement.blur(); true');
  await click('.fld[data-name="agree"]');
  await js('document.querySelectorAll(\'.fld[data-name="size"]\')[1].id = "radio-l"; true');
  await click('#radio-l');
  await js(`(() => { const s = document.querySelector('.fld[data-name="country"]'); s.value = 'Pakistan'; s.dispatchEvent(new Event('change')); return true; })()`);
  ok('form values recorded', await js(`JSON.stringify(${P}.S.form) === JSON.stringify({ full_name: 'Nisar Khan', agree: true, size: '1', country: 'Pakistan' })`), await js(`JSON.stringify(${P}.S.form)`));

  // add items through the editor's own functions
  const png = fs.readFileSync(path.join(__dirname, '..', 'build', 'icon.png')).toString('base64');
  await js(`(() => {
    const v = [...${P}.views.values()][0];
    const C = ${P}.createObject;
    C(v, 'text', 300, 120, 200, 20, { text: 'Hello PDF', size: 16, autoW: true });
    C(v, 'text', 300, 160, 200, 20, { text: 'سلام دنیا', size: 18, autoW: true });
    C(v, 'rect', 320, 300, 120, 60);
    C(v, 'ellipse', 460, 300, 100, 60);
    C(v, 'line', 320, 380, 120, 20, { pts: [[0, 0.5], [1, 0.5]] });
    C(v, 'arrow', 460, 380, 100, 30, { pts: [[0, 1], [1, 0]] });
    C(v, 'ink', 320, 420, 100, 40, { strokes: [[[0, 1], [0.3, 0], [0.6, 1], [1, 0]]] });
    C(v, 'highlight', 70, 86, 260, 20);
    C(v, 'underline', 70, 136, 300, 16);
    C(v, 'strike', 70, 160, 120, 16);
    C(v, 'whiteout', 400, 40, 100, 20);
    C(v, 'note', 560, 60, 22, 22, { text: 'Please check this' });
    C(v, 'stamp', 380, 480, 150, 36, { text: 'APPROVED', color: '#1e8e3e' });
    C(v, 'image', 70, 560, 80, 80, { src: 'data:image/png;base64,${png}' });
    const v2 = [...${P}.views.values()][1];
    C(v2, 'redact', 70, 160, 200, 22);
    ${P}.commit();
    return true;
  })()`);
  await wait(500);
  ok('items added', await js(`${P}.S.entries[0].objects.length === 14 && ${P}.S.entries[1].objects.length === 1`));
  ok('items drawn on page', (await js('document.querySelectorAll(".pg .obj").length')) >= 15);
  await js(`${P}.select(null); true`);
  await shot('items-added');

  // real mouse: select and drag an added rectangle
  const before = await js(`${P}.S.entries[0].objects.find(o => o.type === 'rect').m.join(',')`);
  const rr = await js(`(() => { const o = ${P}.S.entries[0].objects.find(o => o.type === 'rect'); const el = document.querySelector('.obj[data-oid="' + o.id + '"]'); el.scrollIntoView({ block: 'center' }); const b = el.getBoundingClientRect(); return [b.left + 20, b.top + 20]; })()`);
  await wait(200);
  if (process.env.NW_DEBUG) console.log('DEBUG', rr, await js(`(() => { const el = document.elementFromPoint(${Math.round(rr[0])}, ${Math.round(rr[1])}); return (el ? el.className + '|' + (el.closest('.obj') ? 'obj' : '-') : 'none') + ' win=' + innerWidth + 'x' + innerHeight + ' scrollY=' + document.scrollingElement.scrollTop + ' tool=' + document.body.dataset.tool; })()`));
  await drag(Math.round(rr[0]), Math.round(rr[1]), Math.round(rr[0]) + 60, Math.round(rr[1]) + 40);
  ok('drag moves an item', (await js(`${P}.S.entries[0].objects.find(o => o.type === 'rect').m.join(',')`)) !== before);
  ok('item selected with toolbar', await js('!document.getElementById("objbar").hidden'));
  await shot('item-selected');

  // real mouse: Add Text tool, click and type
  await js(`${P}.setTool('text'); document.querySelector('.pg').scrollIntoView({ block: 'start' }); true`);
  await wait(200);
  await click('.pg', { at: [0.55, 0.06], noScroll: true });
  wc.insertText('Typed on the page');
  await wait(300);
  await js('document.activeElement.blur(); true');
  await wait(200);
  ok('Add Text by clicking', await js(`${P}.S.entries[0].objects.some(o => o.type === 'text' && o.text === 'Typed on the page')`));

  // real mouse: highlight by selecting text in the text layer
  await js(`${P}.setTool('highlight'); true`);
  const span = await js(`(() => { const s = [...document.querySelectorAll('.pg .textLayer span')].find(s => s.textContent.includes('quick brown')); const b = s.getBoundingClientRect(); return [b.left + 2, b.top + b.height / 2, b.right - 2]; })()`);
  const hlBefore = await js(`${P}.S.entries[0].objects.filter(o => o.type === 'highlight').length`);
  await drag(Math.round(span[0]), Math.round(span[1]), Math.round(span[2]), Math.round(span[1]));
  await wait(300);
  ok('highlight by selecting text', (await js(`${P}.S.entries[0].objects.filter(o => o.type === 'highlight').length`)) === hlBefore + 1);
  await js(`${P}.setTool('select'); true`);

  // edit existing text
  await js(`${P}.setTool('edittext'); true`);
  await click('.pg .textLayer span');
  await wait(200);
  ok('Edit Text creates whiteout + text', await js(`(() => { const o = ${P}.S.entries[0].objects; return o.filter(x => x.type === 'whiteout').length === 2 && ${P}.S.editing; })()`));
  wc.insertText('Edited heading');
  await wait(200);
  await js('document.activeElement.blur(); true');
  await wait(200);
  ok('edited text kept', await js(`${P}.S.entries[0].objects.some(o => o.type === 'text' && o.text === 'Edited heading')`));
  await shot('edited');

  // find
  await js(`document.getElementById('find-input').value = 'Secret code'; ${P}.runFind()`);
  await wait(800);
  ok('find across pages', (await js(`${P}.find.matches.length`)) === 3, String(await js(`${P}.find.matches.length`)));
  ok('find highlight painted', (await js('document.querySelectorAll(".search-layer .hit").length')) >= 1);
  await shot('find');
  await js(`document.querySelector('[data-find="close"]').click(); true`);

  // pages: rotate, duplicate, move, insert, delete
  await js(`${P}.S.current = 2; ${P}.rotatePages(1); true`);
  ok('rotate page', await js(`${P}.S.entries[2].rot === 90`));
  const other = fs.readFileSync(path.join(tmp, 'other.pdf')).toString('base64');
  await js(`${P}.S.current = 3; ${P}.insertPdfBytes({ bytes: Uint8Array.fromBase64('${other}'), name: 'other.pdf' })`);
  await wait(500);
  ok('insert pages from another PDF', await js(`${P}.S.entries.length === 6 && ${P}.S.entries[4].src === 1`));
  await js(`${P}.S.current = 0; ${P}.S.selPages.clear(); ${P}.duplicatePages(); true`);
  ok('duplicate page (with its items)', await js(`${P}.S.entries.length === 7 && ${P}.S.entries[1].objects.length === ${P}.S.entries[0].objects.length`));
  await js(`${P}.S.selPages.clear(); ${P}.S.current = 1; ${P}.deletePages()`);
  await wait(300);
  ok('delete page', await js(`${P}.S.entries.length === 6`));
  await js(`${P}.S.current = 5; ${P}.movePages(-1); true`);
  ok('move page', await js(`${P}.S.entries[4].src === 1 && ${P}.S.entries[4].index === 1`));
  await js(`${P}.undo(); true`);
  ok('undo page move', await js(`${P}.S.entries[5].src === 1 && ${P}.S.entries[5].index === 1`));
  await js(`${P}.redo(); true`);
  ok('redo page move', await js(`${P}.S.entries[4].src === 1 && ${P}.S.entries[4].index === 1`));
  await wait(400);
  await shot('pages');

  // save
  const out = path.join(tmp, 'saved.pdf');
  await js(`${P}.S.path = ${JSON.stringify(out)}; ${P}.S.name = 'saved.pdf'; true`);
  ok('save', (await js(`${P}.save(false)`)) === true && fs.existsSync(out));
  ok('not dirty after save', !(await js(`${P}.S.dirty`)));
  const saved = await L.PDFDocument.load(fs.readFileSync(out));
  ok('saved page count', saved.getPageCount() === 6, String(saved.getPageCount()));
  ok('saved rotation', saved.getPage(2).getRotation().angle === 90, String(saved.getPage(2).getRotation().angle));
  ok('saved inserted page size', Math.round(saved.getPage(5).getSize().width) === 595 || Math.round(saved.getPage(4).getSize().width) === 595);
  const f = saved.getForm();
  ok('saved form text', f.getTextField('full_name').getText() === 'Nisar Khan');
  ok('saved form checkbox', f.getCheckBox('agree').isChecked());
  ok('saved form radio', f.getRadioGroup('size').getSelected() === 'L');
  ok('saved form dropdown', f.getDropdown('country').getSelected()[0] === 'Pakistan');
  const annots = saved.getPage(0).node.Annots();
  const subtypes = annots ? annots.asArray().map(r => saved.context.lookup(r).get(L.PDFName.of('Subtype'))?.toString()) : [];
  ok('sticky note saved as a real PDF comment', subtypes.includes('/Text'), subtypes.join(','));

  // check the saved file's text with pdf.js
  const textCheck = await js(`(async () => {
    const { pdfjsLib, PDF_OPTS } = ${P};
    const bytes = await (await fetch('file:///' + ${JSON.stringify(out.replace(/\\/g, '/'))})).arrayBuffer().catch(() => null);
    return null;
  })()`);
  void textCheck;
  const savedBytes = fs.readFileSync(out).toString('base64');
  const texts = await js(`(async () => {
    const { pdfjsLib, PDF_OPTS } = ${P};
    const doc = await pdfjsLib.getDocument({ data: Uint8Array.fromBase64('${savedBytes}'), ...PDF_OPTS }).promise;
    const out = [];
    for (let i = 1; i <= doc.numPages; i++) { const tc = await (await doc.getPage(i)).getTextContent(); out.push(tc.items.map(t => t.str).join(' ')); }
    const att = await doc.getAttachments();
    return { pages: out, att: att ? Object.keys(att) : [] };
  })()`);
  ok('added text is real PDF text', texts.pages[0].includes('Hello PDF') && texts.pages[0].includes('Typed on the page'), texts.pages[0].slice(0, 200));
  ok('stamp text in PDF', texts.pages[0].includes('APPROVED'));
  ok('redacted page has no hidden text', !texts.pages.some(t => t.includes('Secret code 1002')), texts.pages.join(' | ').slice(0, 300));
  ok('other pages keep their text', texts.pages.some(t => t.includes('Secret code 1003')));
  ok('edit layer stored in file', texts.att.some(a => a.includes('nisarstudio-edits')), texts.att.join(','));

  // reopen: added items come back editable
  const objCount = await js(`${P}.S.entries.reduce((n, e) => n + e.objects.length, 0)`);
  await js(`(async () => { ${P}.S.open = false; await ${P}.openBytes(Uint8Array.fromBase64('${savedBytes}'), 'saved.pdf', ${JSON.stringify(out)}); return true; })()`);
  await wait(1200);
  ok('reopen keeps items editable', (await js(`${P}.S.entries.reduce((n, e) => n + e.objects.length, 0)`)) === objCount, `${objCount}`);
  ok('reopen keeps page order', await js(`${P}.S.entries.length === 6 && ${P}.S.entries[2].rot === 90`));
  ok('reopen keeps form values', await js(`${P}.S.form.full_name === 'Nisar Khan'`));
  await shot('reopened');

  // export to Word opens a document window with the PDF's text
  const before2 = BrowserWindow.getAllWindows().length;
  await js(`${P}.exportToWord()`);
  await wait(2500);
  const docWin = BrowserWindow.getAllWindows().find(w => w !== win && w.webContents.getURL().endsWith('index.html'));
  ok('export to Word opens a document', !!docWin && BrowserWindow.getAllWindows().length === before2 + 1);
  if (docWin) {
    const txt = await docWin.webContents.executeJavaScript('editor.innerText', true);
    ok('converted text present', txt.includes('Original page 1') && txt.includes('quick brown fox'), txt.slice(0, 120));
    ok('converted doc marked unsaved', await docWin.webContents.executeJavaScript('state.dirty', true));
    docWin.destroy();
  }

  // print preview window
  await js(`document.querySelector('[data-cmd="print"]').click(); true`);
  await wait(3000);
  const pv = BrowserWindow.getAllWindows().find(w => w !== win && w.getTitle().includes('Print Preview'));
  ok('print preview opens', !!pv);
  if (pv) { if (shots) fs.writeFileSync(path.join(shots, `${String(++n).padStart(2, '0')}-print-preview.png`), (await pv.webContents.capturePage()).toPNG()); pv.destroy(); }

  // ---- Acrobat-style tools
  const answer = values => js(`(() => { const d = document.getElementById('modal'); if (!d.open) return 'no dialog'; const f = d.querySelector('form'); for (const [k, v] of Object.entries(${JSON.stringify(values)})) { const el = f.elements[k]; if (!el) continue; if (el.type === 'checkbox') el.checked = !!v; else el.value = v; } f.requestSubmit(); return 'ok'; })()`);
  const cmd = async name => { await js(`document.querySelector('[data-cmd="${name}"]').click(); true`); await wait(300); };
  await js(`${P}.S.current = 0; ${P}.S.selPages.clear(); true`);
  await cmd('headerFooter');
  ok('header & footer dialog', (await answer({ hl: 'ACME Ltd', fc: 'Page {page} of {pages}' })) === 'ok');
  await wait(300);
  ok('header & footer on every page', await js(`${P}.S.entries.every(e => e.objects.filter(o => o.hf).length === 2) && ${P}.S.entries[1].objects.some(o => o.hf && o.text === 'Page 2 of 6')`));
  await cmd('watermark');
  await answer({ text: 'DRAFT', angle: '45' });
  await wait(300);
  ok('watermark on every page', await js(`${P}.S.entries.every(e => e.objects.filter(o => o.wm && o.text === 'DRAFT').length === 1)`));
  await js(`${P}.S.current = 1; ${P}.S.selPages.clear(); true`);
  const wBefore = await js(`${P}.geom(${P}.S.entries[1]).w1`);
  await cmd('crop');
  await answer({ top: 1, bottom: 1, left: 1, right: 1, scope: 'current' });
  await wait(600);
  ok('crop page', await js(`!!${P}.S.entries[1].crop && ${P}.geom(${P}.S.entries[1]).w1 < ${wBefore} - 50`), String(await js(`${P}.geom(${P}.S.entries[1]).w1`)));
  await cmd('findRedact');
  await answer({ q: 'Secret code 1003' });
  await wait(500);
  const redactPage = await js(`${P}.S.entries.findIndex(e => e.objects.some(o => o.type === 'redact') && e.index === 2)`);
  ok('find & redact marks matches', redactPage >= 0);
  await js(`(() => { const v = [...${P}.views.values()][0]; ${P}.createObject(v, 'field', 300, 640, 180, 22, { kind: 'text', name: 'MyField' }); ${P}.createObject(v, 'field', 300, 670, 14, 14, { kind: 'check', name: 'MyCheck' }); ${P}.commit(); return true; })()`);
  ok('form fields drawn', await js(`${P}.S.entries[0].objects.filter(o => o.type === 'field').length === 2`));
  // OCR: a page that is only a picture of text
  await js(`(async () => {
    const c = document.createElement('canvas'); c.width = 1700; c.height = 600; const x = c.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.fillStyle = '#000'; x.font = '90px Arial';
    x.fillText('Invoice Number 48213', 60, 200); x.fillText('Total due by Friday', 60, 380);
    const url = c.toDataURL('image/png');
    const S = ${P}.S;
    const ent = { id: 'scan1', src: 0, index: 0, rot: 0, blank: { w: 612, h: 792 }, objects: [] };
    ent.objects.push({ id: 'scanimg', type: 'image', m: [1, 0, 0, -1, 30, 762], w: 552, h: 195, src: url, kind: 'image' });
    S.entries.push(ent); S.current = S.entries.length - 1; ${P}.commit(); return true; })()`);
  await Promise.race([js(`${P}.recognizeText('current')`), wait(120000)]);
  await wait(500);
  const ocrWords = await js(`(${P}.S.entries.find(e => e.id === 'scan1').ocr || []).map(w => w.t).join(' ')`);
  ok('OCR recognises text in a picture', /Invoice/i.test(ocrWords) && /48213/.test(ocrWords), ocrWords);
  // password protection
  await cmd('protect');
  await answer({ user: 'open123', owner: 'owner456', print: true, copy: false, modify: false });
  await wait(200);
  ok('protect settings', await js(`${P}.S.protect && ${P}.S.protect.user === 'open123'`));
  const prot = path.join(tmp, 'protected.pdf');
  await js(`${P}.S.path = ${JSON.stringify(prot)}; ${P}.S.name = 'protected.pdf'; true`);
  ok('save protected', (await js(`${P}.save(false)`)) === true && fs.existsSync(prot));
  let encrypted = false;
  try { await L.PDFDocument.load(fs.readFileSync(prot)); } catch (err) { encrypted = /encrypt/i.test(err.message); }
  ok('saved file is encrypted', encrypted);
  const protB64 = fs.readFileSync(prot).toString('base64');
  const reopen = js(`(async () => { ${P}.S.open = false; await ${P}.openBytes(Uint8Array.fromBase64('${protB64}'), 'protected.pdf', ${JSON.stringify(prot)}); return ${P}.S.entries.length; })()`);
  await wait(1500);
  ok('password prompt on open', (await answer({ pw: 'open123' })) === 'ok');
  const pagesAfter = await reopen;
  await wait(1200);
  ok('opens with password, edits intact', pagesAfter === 7 && (await js(`${P}.S.entries.find(e => e.id === 'scan1') ? 1 : 0`)) === 1, String(pagesAfter));
  ok('stays protected', await js(`!!${P}.S.protect && ${P}.S.protect.user === 'open123'`));
  // what's inside the saved file (decrypted with qpdf, as Acrobat would with the password)
  const plainOut = await js(`(async () => { const r = await api.qpdf({ op: 'decrypt', bytes: Uint8Array.fromBase64('${protB64}'), password: 'open123' }); return r.bytes ? r.bytes.toBase64() : r.error; })()`);
  const plainDoc = await L.PDFDocument.load(Buffer.from(plainOut, 'base64'));
  const cropIdx = await js(`${P}.S.entries.findIndex(e => e.crop)`);
  // cropped through the crop box, or (for a page that also has redactions) rendered at the cropped size
  ok('crop saved', cropIdx >= 0 && Math.round(plainDoc.getPage(cropIdx).getCropBox().width) < 612 - 50,`entry ${cropIdx}: crop ${cropIdx >= 0 ? plainDoc.getPage(cropIdx).getCropBox().width : '-'} media ${cropIdx >= 0 ? plainDoc.getPage(cropIdx).getMediaBox().width : '-'}`);
  ok('new form fields saved', plainDoc.getForm().getFields().map(f => f.getName()).includes('MyField') && plainDoc.getForm().getFields().map(f => f.getName()).includes('MyCheck'));
  const plainTexts = await js(`(async () => { const { pdfjsLib, PDF_OPTS } = ${P}; const doc = await pdfjsLib.getDocument({ data: Uint8Array.fromBase64('${plainOut}'), ...PDF_OPTS }).promise; const out = []; for (let i = 1; i <= doc.numPages; i++) out.push((await (await doc.getPage(i)).getTextContent()).items.map(t => t.str).join(' ')); return out; })()`);
  ok('header/footer and watermark are real text', plainTexts[0].includes('ACME Ltd') && plainTexts[0].includes('DRAFT') && plainTexts[2].includes('Page 3 of'));
  ok('OCR text searchable in saved PDF', plainTexts[plainTexts.length - 1].includes('Invoice'), plainTexts[plainTexts.length - 1]);
  ok('find & redact removed the text', !plainTexts.some(t => t.includes('Secret code 1003')));
  // remove hidden info
  await js(`${P}.S.sanitize = true; true`);
  const sanB64 = await js(`(async () => (await ${P}.buildPdf({ keepEdits: true })).toBase64())()`);
  const sanDoc = await L.PDFDocument.load(Buffer.from(sanB64, 'base64'));
  ok('remove hidden info: no attachments or title', !sanDoc.catalog.lookupMaybe(L.PDFName.of('Names'), L.PDFDict)?.get(L.PDFName.of('EmbeddedFiles')) && !sanDoc.getTitle());
  await js(`${P}.S.sanitize = false; true`);
  // reduce size: a big photo gets smaller
  const bigJpeg = await js(`(async () => { const c = document.createElement('canvas'); c.width = 4000; c.height = 3000; const x = c.getContext('2d'); for (let i = 0; i < 400; i++) { x.fillStyle = 'hsl(' + (i * 37 % 360) + ',70%,50%)'; x.fillRect(Math.random() * 4000, Math.random() * 3000, 300, 300); } const b = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.95)); const doc = await PDFLib.PDFDocument.create(); const img = await doc.embedJpg(new Uint8Array(await b.arrayBuffer())); doc.addPage([612, 792]).drawImage(img, { x: 0, y: 0, width: 612, height: 459 }); const before = await doc.save(); const after = await ${P}.downsampleImages(before, 1600, 0.72); return [before.length, after.length]; })()`);
  ok('reduce file size shrinks photos', bigJpeg[1] < bigJpeg[0] * 0.6, bigJpeg.join(' -> '));
  // side panel
  await js(`${P}.showSide('comments'); true`);
  await wait(200);
  ok('comments panel lists items', (await js('document.querySelectorAll("#comment-list .cm-item").length')) > 3);
  await js(`${P}.showSide('bookmarks'); true`);
  await wait(300);
  ok('bookmarks panel', await js('!!document.querySelector("#bookmark-list .side-empty, #bookmark-list .bm")'));
  await js(`${P}.showSide('pages'); true`);
  await shot('acrobat-tools');

  // tabs & menus render
  for (const tab of ['comment', 'edit', 'pages', 'home']) { await js(`document.querySelector('[data-tab="${tab}"]').click(); true`); await shot(`tab-${tab}`); }
  for (const m of ['zoom', 'stamps', 'color', 'width', 'opacity', 'images']) {
    const c = await js(`(() => { const b = document.querySelector('[data-menu="${m}"]'); openMenu('${m}', b); const n = document.querySelectorAll('.menu').length; closeMenu(); return n; })()`);
    if (c !== 1) ok(`menu ${m}`, false);
  }
  ok('menus open', true);
  await js(`document.querySelector('[data-cmd="signature"]').click(); true`);
  await wait(400);
  ok('signature dialog', await js('document.getElementById("modal").open && !!document.querySelector(".sig-pad")'));
  await shot('signature');
  await js('document.getElementById("modal").close(); true');
  await js(`document.querySelector('[data-cmd="backstage"]').click(); true`);
  await shot('backstage');
  await js(`document.querySelector('[data-cmd="closeBackstage"]').click(); true`);

  ok('no renderer errors', errors.length === 0, errors.join(' | '));
  finish();
}

function finish() {
  console.log('\n' + results.join('\n'));
  if (errors.length) console.log('\nRenderer errors:\n' + errors.join('\n'));
  const failed = results.filter(r => r.startsWith('FAIL')).length;
  console.log(`\n${results.length - failed}/${results.length} checks passed`);
  try { fs.rmSync(tmp, { recursive: true, force: true }); } catch { /* ignore */ }
  app.exit(failed ? 1 : 0);
}
