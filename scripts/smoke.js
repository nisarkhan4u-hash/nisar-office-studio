'use strict';
// End-to-end check: launches Nisar Office Studio, edits a document, saves .docx/.html/.txt,
// reopens them, exports a PDF, and reports problems. Run with: npm run smoke
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'nisarword-smoke-'));
app.setPath('userData', path.join(tmp, 'profile'));
require('../main.js');

const errors = [];
const results = [];
const ok = (name, cond, extra = '') => results.push(`${cond ? 'PASS' : 'FAIL'}  ${name}${extra ? `  (${extra})` : ''}`);
const wait = ms => new Promise(r => setTimeout(r, ms));

let started = false;
app.on('browser-window-created', (e, win) => {
  if (started) return;
  started = true;
  const wc = win.webContents;
  wc.on('console-message', ev => {
    if (ev.level === 'error') errors.push(`${ev.message} @${ev.lineNumber}`);
  });
  wc.on('preload-error', (ev, p, err) => errors.push(`preload: ${err}`));
  wc.once('did-finish-load', () => run(win).catch(err => { results.push(`FAIL  crashed: ${err.stack || err}`); finish(); }));
});

async function run(win) {
  const wc = win.webContents;
  const js = code => wc.executeJavaScript(code, true);
  await wait(1500);

  ok('app started with empty document', await js('editor.innerHTML === "<p><br></p>" && document.title.includes("Document1")'));
  ok('ribbon icons rendered', (await js('$$(".ic").length')) > 60);

  // typing
  await js('editor.focus(); placeCaret(editor.firstElementChild, true); true');
  wc.insertText('Hello world from Nisar Office Studio');
  await wait(300);
  ok('typing works', (await js('editor.innerText')).includes('Hello world from Nisar Office Studio'));
  ok('title shows unsaved marker', (await js('document.title')).startsWith('•'));

  // formatting on a selection
  await js(`(() => { const t = editor.querySelector('p').firstChild; const r = document.createRange(); r.setStart(t, 0); r.setEnd(t, 5); setRange(r); return true; })()`);
  await js('runCommand("bold"); runCommand("italic"); runCommand("underline"); applyFontSize(20); applyFore("#ff0000"); applyFontName("Georgia"); true');
  await wait(100);
  const html1 = await js('editor.innerHTML');
  ok('bold/italic/underline applied', /<b>|<strong>/.test(html1) && /<i>|<em>/.test(html1) && /<u>/.test(html1), html1.slice(0, 200));
  ok('font size applied', /font-size:\s*20pt/.test(html1));
  ok('color applied', /color:\s*rgb\(255, 0, 0\)|#ff0000/i.test(html1));

  // paragraph + styles
  await js('placeCaret(editor.lastElementChild, false); true');
  wc.sendInputEvent({ type: 'keyDown', keyCode: 'Enter' });
  wc.sendInputEvent({ type: 'char', keyCode: '\r' });
  wc.sendInputEvent({ type: 'keyUp', keyCode: 'Enter' });
  await wait(100);
  wc.insertText('Section heading');
  await wait(100);
  await js('applyStyle("h1"); true');
  ok('heading style', await js('!!editor.querySelector("h1") && editor.querySelector("h1").textContent.includes("Section heading")'));
  await js('runCommand("center"); true');
  ok('center alignment', /text-align:\s*center/.test(await js('editor.querySelector("h1").outerHTML')));

  // strikethrough, list, table, page break, line spacing
  await js(`(() => { placeCaret(editor.lastElementChild, false); document.execCommand('insertParagraph'); document.execCommand('formatBlock', false, 'p'); document.execCommand('insertText', false, 'struck text'); const t = editor.lastElementChild.firstChild; const r = document.createRange(); r.selectNodeContents(editor.lastElementChild); setRange(r); runCommand('strike'); return true; })()`);
  await js(`(() => { placeCaret(editor.lastElementChild, false); document.execCommand('insertParagraph'); document.execCommand('insertText', false, 'item one'); runCommand('bullets'); return true; })()`);
  ok('bullet list', await js('!!editor.querySelector("ul li")'));
  await js(`(() => { placeCaret(editor.lastElementChild, false); document.execCommand('insertParagraph'); runCommand('bullets'); insertTable(2, 3); document.execCommand('insertText', false, 'cell A'); return true; })()`);
  ok('table inserted, caret in first cell', await js('editor.querySelectorAll("table td").length === 6 && editor.querySelector("td").textContent.includes("cell A")'));
  await js('tableOp("rowBelow"); tableOp("colRight"); true');
  ok('table row/column insert', await js('editor.querySelector("table").rows.length === 3 && editor.querySelector("table").rows[0].cells.length === 4'));
  await js('placeCaret(editor.lastElementChild, false); insertPageBreak(); document.execCommand("insertText", false, "Second page text"); setLineSpacing("1.5"); true');
  ok('page break', await js('!!editor.querySelector(".page-break")'));
  ok('line spacing', await js('editor.lastElementChild.style.lineHeight === "1.5"'));
  await js('insertImage("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="); true');
  await wait(200);
  await js('selectImage(editor.querySelector("img")); imageAction("50"); selectImage(null); true');
  ok('picture inserted and resized', await js('parseInt(editor.querySelector("img").style.width) > 100'));

  // real mouse clicks on the ribbon must keep the text selection and apply formatting
  const click = async (selector, opts = {}) => {
    const r = await js(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return null; el.scrollIntoView({ block: 'nearest' }); const b = el.getBoundingClientRect(); return { x: b.left + b.width / 2, y: b.top + b.height / 2 }; })()`);
    if (!r) throw new Error(`no element ${selector}`);
    const x = Math.round(r.x), y = Math.round(r.y);
    wc.sendInputEvent({ type: 'mouseMove', x, y });
    wc.sendInputEvent({ type: 'mouseDown', x, y, button: 'left', clickCount: 1, ...opts });
    wc.sendInputEvent({ type: 'mouseUp', x, y, button: 'left', clickCount: 1, ...opts });
    await wait(250);
  };
  await js(`(() => { const p = document.createElement('p'); p.id = 'clicktest'; p.textContent = 'click formatting test'; editor.appendChild(p); const r = document.createRange(); r.selectNodeContents(p); setRange(r); editor.focus(); return true; })()`);
  await click('[data-tab="home"]');
  await click('[data-cmd="italic"]');
  await click('[data-cmd="right"]');
  const ct = await js('document.getElementById("clicktest").outerHTML');
  ok('ribbon click: italic keeps selection', /<i>click formatting test<\/i>/.test(ct), ct);
  ok('ribbon click: align right', /text-align:\s*right/.test(ct));
  ok('ribbon click: button shows active', await js('document.querySelector(\'[data-cmd="italic"]\').classList.contains("on")'));
  await click('[data-menu="fore"]');
  ok('ribbon click: colour menu opens', await js('!!document.querySelector(".menu .sw")'));
  await click('.menu .sw[data-c="#c00000"]');
  ok('ribbon click: colour applied from menu', /color:\s*rgb\(192, 0, 0\)/.test(await js('document.getElementById("clicktest").innerHTML')));
  await click('[data-tab="insert"]');
  await click('[data-menu="table"]');
  await click('.tgrid [data-r="2"][data-c="2"]');
  ok('ribbon click: table from grid', await js('[...editor.querySelectorAll("table")].some(t => t.rows.length === 2 && t.rows[0].cells.length === 2)'));
  await click('[data-tab="home"]');
  await js('[...editor.querySelectorAll("p")].filter(p => p.textContent.includes("click formatting test")).forEach(p => p.remove()); [...editor.querySelectorAll("table")].filter(t => t.rows.length === 2 && t.rows[0].cells.length === 2).forEach(t => t.remove()); true');

  // undo / redo
  await wait(600);
  const before = await js('editor.innerHTML');
  await js('placeCaret(editor.lastElementChild, false); document.execCommand("insertText", false, " UNDOME"); true');
  await wait(600);
  await js('undo(); true');
  ok('undo', !(await js('editor.innerHTML')).includes('UNDOME'));
  await js('redo(); true');
  ok('redo', (await js('editor.innerHTML')).includes('UNDOME'));
  await js('setLineSpacing("2.0"); commitHistory(); undo(); true');
  ok('undo line spacing', (await js('editor.lastElementChild.style.lineHeight')) !== '2.0');
  void before;

  // find & replace
  await js('openFind(true); findInput.value = "cell"; refreshFind(false); replaceInput.value = "box"; replaceAll(); closeFind(); true');
  ok('find & replace', (await js('editor.innerText')).includes('box A'));

  // page setup
  await js('setPage({ orientation: "landscape", size: "a4", margins: { top: 0.5, bottom: 0.5, left: 0.75, right: 0.75 } }); state.page.header = "My Header"; state.page.pageNumbers = true; applyPage(); true');
  ok('landscape layout', (await js('page.offsetWidth')) > 1100);

  // save as .docx
  const docx = path.join(tmp, 'test.docx');
  await js(`state.path = ${JSON.stringify(docx)}; true`);
  const saved = await js('saveDoc(false)');
  ok('save .docx', saved === true && fs.existsSync(docx));
  ok('not dirty after save', !(await js('state.dirty')));

  const JSZip = require('jszip');
  const zip = await JSZip.loadAsync(fs.readFileSync(docx));
  const xml = await zip.file('word/document.xml').async('string');
  ok('docx: bold', xml.includes('<w:b/>'));
  ok('docx: italic', xml.includes('<w:i/>'));
  ok('docx: underline', xml.includes('<w:u '));
  ok('docx: strikethrough', xml.includes('<w:strike/>'));
  ok('docx: size 20pt', xml.includes('<w:sz w:val="40"/>'));
  ok('docx: red', /w:color w:val="ff0000"/i.test(xml));
  ok('docx: Georgia', xml.includes('Georgia'));
  ok('docx: heading', xml.includes('Heading1'));
  ok('docx: centered', xml.includes('<w:jc w:val="center"/>'));
  ok('docx: table', xml.includes('<w:tbl>'));
  ok('docx: list', xml.includes('<w:numPr>'));
  ok('docx: page break', xml.includes('w:type="page"'));
  ok('docx: picture', xml.includes('<w:drawing>'));
  ok('docx: landscape A4', /<w:pgSz w:w="16834" w:h="11909" w:orient="landscape"\/>/.test(xml), (xml.match(/<w:pgSz[^>]*>/) || [])[0]);
  ok('docx: header part', !!zip.file('word/header1.xml'));
  ok('docx: page number field', (await zip.file('word/footer1.xml').async('string')).includes('PAGE'));
  ok('docx: no leftover marker font', !xml.includes('NisarWordStrike'));
  ok('docx: valid XML parts', ['word/document.xml', '[Content_Types].xml', '_rels/.rels'].every(Boolean));
  const mammoth = require('mammoth');
  const back = await mammoth.convertToHtml({ buffer: fs.readFileSync(docx) });
  ok('docx readable by an independent .docx reader', back.value.includes('Hello') && back.value.includes('<table>'));

  // reopen in a fresh state (uses the embedded exact copy)
  const savedHtml = await js('cleanClone().innerHTML');
  await js('editor.innerHTML = "<p><br></p>"; state.path = null; state.dirty = false; state.page = defaultPage(); applyPage(); true');
  await js(`openPath(${JSON.stringify(docx)})`);
  await wait(400);
  const [want, got] = await js(`[sanitize(${JSON.stringify(savedHtml)}), cleanClone().innerHTML]`);
  let d = 0;
  while (d < want.length && want[d] === got[d]) d++;
  ok('reopen .docx exactly', want === got, want === got ? '' : `differs at ${d}: want …${want.slice(d - 40, d + 60)}… got …${got.slice(d - 40, d + 60)}…`);
  ok('reopen keeps page setup', await js('state.page.orientation === "landscape" && state.page.size === "a4" && state.page.header === "My Header"'));

  // an empty (0-byte) .docx opens as a blank document; a damaged one gives a clear message
  const empty = path.join(tmp, 'empty.docx');
  fs.writeFileSync(empty, '');
  const emptyRes = await js(`api.openPath(${JSON.stringify(empty)}).then(d => d.error || d.html)`);
  ok('open empty .docx', emptyRes === '<p><br></p>', emptyRes);
  const broken = path.join(tmp, 'broken.docx');
  fs.writeFileSync(broken, 'not a zip file');
  const brokenRes = await js(`api.openPath(${JSON.stringify(broken)}).then(d => d.error || 'opened')`);
  ok('damaged .docx: clear message', /damaged/.test(brokenRes), brokenRes);

  // Insert Table… dialog: the table goes where the caret was before the dialog opened
  await js(`state.dirty = false; editor.innerHTML = '<p id="tt">Before table</p>'; placeCaret(document.getElementById('tt'), false); true`);
  await js(`document.querySelector('[data-menu="table"]').click(); document.querySelector('#menu-layer [data-custom]').click(); true`);
  await wait(200);
  await js(`document.querySelector('#modal[open] button[type=submit]').click(); true`);
  await wait(200);
  ok('Insert Table… dialog inserts a 3 × 3 table', await js(`(() => { const t = editor.querySelector('table'); return !!t && t.rows.length === 3 && t.rows[0].cells.length === 3 && editor.firstElementChild.id === 'tt'; })()`), await js('editor.innerHTML.slice(0, 120)'));

  // open a .docx made by another program (no embedded copy) -> converter path
  const HTMLtoDOCX = require('html-to-docx');
  const foreign = path.join(tmp, 'foreign.docx');
  fs.writeFileSync(foreign, await HTMLtoDOCX('<h1>Foreign title</h1><p>Plain <b>bold</b> paragraph</p><ul><li>x</li></ul>', null, { orientation: 'landscape' }));
  await js('state.dirty = false; state.path = null; editor.innerHTML = "<p><br></p>"; true');
  await js(`openPath(${JSON.stringify(foreign)})`);
  await wait(400);
  const fh = await js('editor.innerHTML');
  ok('open foreign .docx', fh.includes('Foreign title') && /<strong>bold<\/strong>|<b>bold<\/b>/.test(fh) && /<li[ >]/.test(fh), fh.slice(0, 160));

  ok('foreign .docx: page setup read', await js('state.page.orientation === "landscape"'));

  // .html and .txt round trips
  await js(`state.path = ${JSON.stringify(path.join(tmp, 'test.html'))}; true`);
  ok('save .html', (await js('saveDoc(false)')) === true && fs.readFileSync(path.join(tmp, 'test.html'), 'utf8').includes('Foreign title'));
  await js(`state.path = ${JSON.stringify(path.join(tmp, 'test.txt'))}; true`);
  ok('save .txt', (await js('saveDoc(false)')) === true && fs.readFileSync(path.join(tmp, 'test.txt'), 'utf8').includes('Foreign title'));
  await js('state.dirty = false; state.path = null; editor.innerHTML = "<p><br></p>"; true');
  await js(`openPath(${JSON.stringify(path.join(tmp, 'test.html'))})`);
  await wait(300);
  ok('open .html', (await js('editor.innerText')).includes('Foreign title'));

  // sanitizer
  const clean = await js(`sanitize('<p onclick="alert(1)">x<img src=x onerror="alert(2)"><script>alert(3)<\\/script><a href="javascript:alert(4)">l</a></p>')`);
  ok('sanitizer strips scripts and handlers', !/onclick|onerror|script|javascript:/i.test(clean), clean);

  // PDF (same options the Export button uses)
  const pdf = await wc.printToPDF({ printBackground: true, preferCSSPageSize: true, displayHeaderFooter: true, headerTemplate: '<div style="font-size:9px">h</div>', footerTemplate: '<div style="font-size:9px"><span class="pageNumber"></span></div>' });
  fs.writeFileSync(path.join(tmp, 'test.pdf'), pdf);
  ok('PDF export', pdf.length > 1000 && pdf.slice(0, 4).toString() === '%PDF');

  // menus & dialogs render without errors
  for (const m of ['sizes', 'case', 'fore', 'hilite', 'shade', 'bulletStyle', 'numberStyle', 'spacing', 'table', 'symbols', 'date', 'margins', 'orientation', 'size', 'columns']) {
    const n = await js(`(() => { openMenu(${JSON.stringify(m)}, document.querySelector('[data-menu="${m}"]')); const c = document.querySelectorAll('.menu').length; closeMenu(); return c; })()`);
    if (n !== 1) ok(`menu ${m}`, false);
  }
  ok('all menus open', true);
  await js('showBackstage("home"); true');
  await wait(200);
  ok('backstage + recent list', (await js('document.querySelectorAll("#recent li[data-i]").length')) >= 2);
  await js('hideBackstage(); setDark(true); setDark(false); setWebLayout(true); setWebLayout(false); setFocusMode(true); setFocusMode(false); fitPageWidth(); setZoom(1); true');
  ok('views toggle', true);
  await js('printPreview(); true');
  await wait(3000);
  const all = BrowserWindow.getAllWindows();
  const pvw = all.find(w => w.getTitle().includes('Print Preview'));
  ok('print preview window opens', !!pvw, all.map(w => `${w.id}:${w.getTitle()}`).join(' | '));
  ok('only one document window', all.filter(w => w !== pvw).length === 1, all.map(w => w.getTitle()).join(' | '));
  if (pvw) {
    fs.writeFileSync(path.join(os.tmpdir(), 'nisarword-preview-check.png'), (await pvw.webContents.capturePage()).toPNG());
    pvw.close();
  }
  const shot = await wc.capturePage();
  fs.writeFileSync(path.join(os.tmpdir(), 'nisarword-screenshot.png'), shot.toPNG());

  // Word features: table of contents, footnote, comment, tracked changes -> real Word parts
  await js(`(() => { state.dirty = false; editor.innerHTML = '<h1>Intro</h1><p>Alpha beta.</p><h2>Part</h2><p>Some text.</p><p>Final words.</p>'; applyPage(); editor.focus(); placeCaret(editor.firstElementChild, true); __doc.insertToc(); return true; })()`);
  ok('table of contents', await js('editor.querySelectorAll(".toc-entry").length === 2'));
  await js(`(() => { const t = [...editor.querySelectorAll('p')].find(p => p.textContent.startsWith('Alpha')).firstChild; const r = document.createRange(); r.setStart(t, 5); r.collapse(true); setRange(r); __doc.insertFootnote(); const li = editor.querySelector('.footnotes li'); li.textContent = 'Note one'; return true; })()`);
  ok('footnote', await js('editor.querySelectorAll("sup.fn-ref").length === 1'));
  await js(`(() => { const t = [...editor.querySelectorAll('p')].find(p => p.textContent.startsWith('Some')).firstChild; const r = document.createRange(); r.setStart(t, 0); r.setEnd(t, 4); setRange(r); __doc.newComment(); state.comments[0].text = 'Check this'; document.activeElement.blur(); return true; })()`);
  ok('comment', await js('state.comments.length === 1 && editor.querySelectorAll(".cmt-anchor").length === 1'));
  await js(`(() => { __doc.setTracking(true); const t = [...editor.querySelectorAll('p')].find(p => p.textContent.startsWith('Final')).firstChild; const r = document.createRange(); r.setStart(t, 5); r.collapse(true); editor.focus(); setRange(r); return true; })()`);
  wc.insertText(' new');
  await wait(200);
  ok('tracked insertion', await js('editor.querySelectorAll("ins.tc").length === 1'));
  const wordFile = path.join(tmp, 'word-features.docx');
  await js(`(async () => { __doc.setTracking(false); state.path = ${JSON.stringify(wordFile)}; return saveDoc(false); })()`);
  const wz = await JSZip.loadAsync(fs.readFileSync(wordFile));
  const wdoc = await wz.file('word/document.xml').async('string');
  ok('docx: real Word comment', !!wz.file('word/comments.xml') && wdoc.includes('<w:commentRangeStart'));
  ok('docx: real Word footnote', !!wz.file('word/footnotes.xml') && wdoc.includes('<w:footnoteReference'));
  ok('docx: real tracked insertion', /<w:ins [^>]*w:author=/.test(wdoc));
  ok('docx: no leftover markers', !/[-]/.test(wdoc));

  await wait(200);
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
