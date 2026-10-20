'use strict';
/* Nisar Office Studio: editor (renderer process) */

// `api` is provided by preload.js (window.api)
const editor = $('#editor');
const page = $('#page');
const zoomer = $('#zoomer');
const scroller = $('#scroller');
const DPI = 96;

const PAPER = {
  letter: { label: 'Letter', desc: '8.5" × 11"', w: 8.5, h: 11 },
  legal: { label: 'Legal', desc: '8.5" × 14"', w: 8.5, h: 14 },
  a4: { label: 'A4', desc: '21 cm × 29.7 cm', w: 8.27, h: 11.69 },
  a5: { label: 'A5', desc: '14.8 cm × 21 cm', w: 5.83, h: 8.27 },
  executive: { label: 'Executive', desc: '7.25" × 10.5"', w: 7.25, h: 10.5 },
  tabloid: { label: 'Tabloid', desc: '11" × 17"', w: 11, h: 17 },
};
const MARGINS = {
  normal: { label: 'Normal', top: 1, bottom: 1, left: 1, right: 1 },
  narrow: { label: 'Narrow', top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 },
  moderate: { label: 'Moderate', top: 1, bottom: 1, left: 0.75, right: 0.75 },
  wide: { label: 'Wide', top: 1, bottom: 1, left: 2, right: 2 },
};
const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const FALLBACK_FONTS = ['Arial', 'Arial Black', 'Bahnschrift', 'Book Antiqua', 'Calibri', 'Calibri Light', 'Cambria', 'Candara', 'Century Gothic', 'Comic Sans MS', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Ebrima', 'Franklin Gothic Medium', 'Gabriola', 'Garamond', 'Georgia', 'Impact', 'Lucida Console', 'Lucida Sans Unicode', 'Palatino Linotype', 'Segoe Print', 'Segoe Script', 'Segoe UI', 'Sitka Text', 'Sylfaen', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
const BLOCK_SEL = 'p,h1,h2,h3,h4,h5,h6,li,blockquote,pre,div:not(.page-break)';
const STYLE_TAGS = { normal: 'p', title: 'h1', h1: 'h1', h2: 'h2', h3: 'h3', quote: 'blockquote', code: 'pre' };
const STRIKE_FONT = 'NisarWordStrike';
// classes that survive opening/pasting (doc-features.js adds its own)
const KEEP_CLASSES = new Set(['doc-title', 'doc-subtitle', 'page-break', 'no-borders']);
// extension points used by doc-features.js
const hooks = { sanitize: [], export: [], beforeSave: [], extra: [], load: [], layout: [], sheet: [], sheetKey: [], paste: [], internalLink: [] };

const state = {
  path: null,
  name: 'Document1',
  dirty: false,
  page: defaultPage(),
  pw: 8.5, ph: 11,
  zoom: 1,
  savedRange: null,
  painter: null,
  lastFore: '#c00000',
  lastHilite: '#ffff00',
  lastShade: '#d9e2f3',
  pageCount: 1,
  inTable: false,
  img: null,
  find: { ranges: [], index: -1 },
  speaking: false,
  savedAt: null,
};

function defaultPage() {
  return { size: 'letter', orientation: 'portrait', margins: { ...MARGINS.normal, label: undefined }, columns: 1, header: '', footer: '', pageNumbers: false };
}

const baseName = n => String(n || '').replace(/\.(docx|html?|txt)$/i, '');


// ---------------------------------------------------------------- selection helpers
const inEditor = n => !!n && (n === editor || editor.contains(n));
const elOf = n => (n && n.nodeType === 1 ? n : n && n.parentElement);
function closestEl(node, sel) {
  const el = elOf(node);
  const r = el && el.closest(sel);
  return r && inEditor(r) && r !== editor ? r : null;
}
function anchorElement() {
  const sel = getSelection();
  const el = sel.rangeCount ? elOf(sel.anchorNode) : null;
  return inEditor(el) ? el : editor;
}
// the element whose formatting the toolbar should show: the first selected text
function styleAnchor() {
  const sel = getSelection();
  if (!sel.rangeCount || !inEditor(sel.anchorNode)) return editor;
  const r = sel.getRangeAt(0);
  let n = r.startContainer;
  const startsInText = n.nodeType === 3 && r.startOffset < n.length;
  if (!r.collapsed && !startsInText) {
    const seg = textSegments(r).find(s => s.node.data.slice(s.s, s.e).replace(/​/g, '').trim());
    if (seg) n = seg.node;
  }
  const el = elOf(n);
  return inEditor(el) ? el : editor;
}
function setRange(r) {
  const sel = getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
  state.savedRange = r.cloneRange();
}
function placeCaret(node, atStart = true) {
  const r = document.createRange();
  r.selectNodeContents(node);
  r.collapse(atStart);
  setRange(r);
}
function restoreSelection() {
  if (document.activeElement !== editor) {
    editor.focus({ preventScroll: true });
    if (state.savedRange && inEditor(state.savedRange.startContainer)) setRange(state.savedRange);
    else placeCaret(editor.lastElementChild || editor, false);
  }
}
function closestBlock(node) {
  let el = elOf(node);
  while (el && el !== editor) {
    if (el.matches(BLOCK_SEL) || el.matches('td,th')) return el.matches('td,th') ? null : el;
    el = el.parentElement;
  }
  return null;
}
function getSelectedBlocks() {
  const sel = getSelection();
  if (!sel.rangeCount || !inEditor(sel.anchorNode)) return [];
  const r = sel.getRangeAt(0);
  let blocks = $$(BLOCK_SEL, editor).filter(b => r.intersectsNode(b) && !b.querySelector(BLOCK_SEL));
  if (!blocks.length) {
    const b = closestBlock(r.startContainer);
    if (b) blocks = [b];
    else {
      document.execCommand('formatBlock', false, 'p');
      const b2 = closestBlock(getSelection().anchorNode);
      if (b2) blocks = [b2];
    }
  }
  return blocks;
}

// ---------------------------------------------------------------- undo history
// Snapshots of the whole document, so every kind of change (typing, formatting,
// tables, pictures, spacing) can be undone the same way.
const history = { stack: [], index: -1, timer: null, bytes: 0 };
function nodePath(node) {
  const path = [];
  while (node && node !== editor) {
    const p = node.parentNode;
    if (!p) return null;
    path.unshift(Array.prototype.indexOf.call(p.childNodes, node));
    node = p;
  }
  return node === editor ? path : null;
}
function nodeAt(path) {
  let n = editor;
  for (const i of path) { if (!n.childNodes[i]) return null; n = n.childNodes[i]; }
  return n;
}
function serializeSelection() {
  const r = state.savedRange;
  if (!r) return null;
  const s = nodePath(r.startContainer), e = nodePath(r.endContainer);
  return s && e ? { s, so: r.startOffset, e, eo: r.endOffset } : null;
}
function restoreSerialized(sel) {
  try {
    const s = sel && nodeAt(sel.s), e = sel && nodeAt(sel.e);
    if (!s || !e) { placeCaret(editor.lastElementChild || editor, false); return; }
    const r = document.createRange();
    r.setStart(s, Math.min(sel.so, s.nodeType === 3 ? s.length : s.childNodes.length));
    r.setEnd(e, Math.min(sel.eo, e.nodeType === 3 ? e.length : e.childNodes.length));
    setRange(r);
  } catch { placeCaret(editor.lastElementChild || editor, false); }
}
function resetHistory() {
  clearTimeout(history.timer);
  history.timer = null;
  const html = stripLayout(editor.innerHTML);
  history.stack = [{ html, sel: serializeSelection() }];
  history.index = 0;
  history.bytes = html.length;
}
function commitHistory() {
  clearTimeout(history.timer);
  history.timer = null;
  const html = stripLayout(editor.innerHTML);
  const cur = history.stack[history.index];
  if (cur && cur.html === html) { cur.sel = serializeSelection(); return; }
  for (const dropped of history.stack.splice(history.index + 1)) history.bytes -= dropped.html.length;
  history.stack.push({ html, sel: serializeSelection() });
  history.bytes += html.length;
  // keep at most 200 steps and roughly 150 MB of history
  while (history.stack.length > 200 || (history.bytes > 150e6 && history.stack.length > 2)) history.bytes -= history.stack.shift().html.length;
  history.index = history.stack.length - 1;
}
function scheduleHistory() {
  clearTimeout(history.timer);
  history.timer = setTimeout(commitHistory, 450);
}
// typing is grouped into one undo step; anything else starts a new step
function flushHistory() {
  if (history.timer) commitHistory();
}
// run a whole command as a single undo step, however many DOM changes it makes
let historyBusy = 0;
function withHistory(fn) {
  flushHistory();
  historyBusy++;
  const done = () => { historyBusy--; if (!historyBusy) commitHistory(); };
  let r;
  try { r = fn(); } catch (err) { done(); throw err; }
  if (r && typeof r.then === 'function') return r.finally(done);
  done();
  return r;
}
runMenuAction = fn => withHistory(fn);
reportError = err => showError(err);
function applyHistory(step) {
  const s = history.stack[history.index];
  if (!s) return;
  editor.innerHTML = s.html;
  selectImage(null);
  editor.focus({ preventScroll: true });
  restoreSerialized(step.sel || s.sel);
  if (!state.dirty) { state.dirty = true; updateTitle(); }
  scheduleLayout();
  scheduleStats();
  updateUi();
}
function undo() {
  if (history.timer) commitHistory();
  if (history.index <= 0) return;
  const from = history.stack[history.index];
  history.index--;
  applyHistory({ sel: history.stack[history.index].sel || from.sel });
}
function redo() {
  if (history.timer) commitHistory();
  if (history.index >= history.stack.length - 1) return;
  history.index++;
  applyHistory({});
}

// ---------------------------------------------------------------- commands
// Character styles are written as plain tags (<b>, <i>, <u>…) so they survive .docx export.
const TAG_COMMANDS = new Set(['bold', 'italic', 'underline', 'strikeThrough', 'subscript', 'superscript']);
function exec(cmd, val = null) {
  flushHistory();
  restoreSelection();
  if (TAG_COMMANDS.has(cmd)) document.execCommand('styleWithCSS', false, false);
  document.execCommand(cmd, false, val);
  if (TAG_COMMANDS.has(cmd)) document.execCommand('styleWithCSS', false, true);
  updateUi();
}
function markDirty(typing = false) {
  if (!state.dirty) { state.dirty = true; updateTitle(); }
  if (historyBusy) { /* recorded when the command finishes */ } else if (typing) scheduleHistory(); else commitHistory();
  scheduleLayout(true);
  scheduleStats();
}

const COMMANDS = {
  save: () => saveDoc(false),
  saveAs: () => saveDoc(true),
  undo,
  redo,
  cut: () => { restoreSelection(); api.edit('cut'); },
  copy: () => { restoreSelection(); api.edit('copy'); },
  paste: () => { restoreSelection(); api.edit('paste'); },
  pastePlain,
  selectAll: () => exec('selectAll'),
  painter: toggleFormatPainter,
  bold: () => exec('bold'),
  italic: () => exec('italic'),
  underline: () => exec('underline'),
  strike: () => exec('strikeThrough'),
  sub: () => exec('subscript'),
  sup: () => exec('superscript'),
  grow: () => stepFont(1),
  shrink: () => stepFont(-1),
  clear: clearFormatting,
  fore: () => applyFore(state.lastFore),
  hilite: () => applyHilite(state.lastHilite),
  shade: () => applyShade(state.lastShade),
  bullets: () => toggleList('ul'),
  numbers: () => toggleList('ol'),
  indent: () => indent(1),
  outdent: () => indent(-1),
  left: () => exec('justifyLeft'),
  center: () => exec('justifyCenter'),
  right: () => exec('justifyRight'),
  justify: () => exec('justifyFull'),
  style: b => applyStyle(b.dataset.arg),
  find: () => openFind(false),
  replace: () => openFind(true),
  pageBreak: insertPageBreak,
  picture: insertPicture,
  hr: () => exec('insertHorizontalRule'),
  link: insertLink,
  unlink: () => { restoreSelection(); const a = closestEl(getSelection().anchorNode, 'a'); if (a) { const r = document.createRange(); r.selectNodeContents(a); setRange(r); } exec('unlink'); },
  headerFooter: editHeaderFooter,
  pageNumbers: () => { state.page.pageNumbers = !state.page.pageNumbers; applyPage(); markDirty(); toast(state.page.pageNumbers ? 'Page numbers on. They appear in printed pages, PDF and .docx files.' : 'Page numbers removed'); },
  rowAbove: () => tableOp('rowAbove'),
  rowBelow: () => tableOp('rowBelow'),
  colLeft: () => tableOp('colLeft'),
  colRight: () => tableOp('colRight'),
  delRow: () => tableOp('delRow'),
  delCol: () => tableOp('delCol'),
  delTable: () => tableOp('delTable'),
  headerRow: () => tableOp('headerRow'),
  spell: toggleSpell,
  wordCount: showWordCount,
  speak,
  stopSpeak,
  zoomIn: () => setZoom(Math.round(state.zoom * 10 + 1) / 10),
  zoomOut: () => setZoom(Math.round(state.zoom * 10 - 1) / 10),
  zoom100: () => setZoom(1),
  pageWidth: fitPageWidth,
  onePage: fitOnePage,
  zoomDialog,
  printLayout: () => setWebLayout(false),
  webLayout: () => setWebLayout(true),
  focus: () => setFocusMode(!document.body.classList.contains('focus-mode')),
  boundaries: () => { document.body.classList.toggle('no-boundaries'); updateToggles(); layoutPages(); },
  dark: () => setDark(!document.body.classList.contains('dark')),
  collapseRibbon: () => { document.body.classList.toggle('ribbon-collapsed'); },
  goto: gotoPage,
  backstage: () => showBackstage('home'),
  closeBackstage: hideBackstage,
  new: () => api.newWindow(),
  open: openDoc,
  openPdf,
  pdf: exportPdf,
  print: printDoc,
  printPreview,
  closeDoc: () => { hideBackstage(); window.close(); },
};

function runCommand(name, el) {
  const fn = COMMANDS[name];
  if (!fn) return;
  try {
    const r = name === 'undo' || name === 'redo' ? fn() : withHistory(() => fn(el || {}));
    if (r && typeof r.catch === 'function') r.catch(err => showError(err));
  } catch (err) { showError(err); }
}
function showError(err) {
  console.error(err);
  toast(`Something went wrong: ${err && err.message ? err.message : err}`);
}

// ---------------------------------------------------------------- font & character formatting
function applyFontName(name) {
  if (!name) return;
  restoreSelection();
  document.execCommand('fontName', false, name);
  markDirty();
  updateUi();
}

// The zero-width space that holds a style at an empty caret is removed as soon as
// real text is typed next to it.
function dropCaretMarker() {
  const sel = getSelection();
  const n = sel.anchorNode;
  if (!n || n.nodeType !== 3 || !n.data.includes('​') || n.data.length < 2) return;
  const off = sel.anchorOffset;
  const before = (n.data.slice(0, off).match(/​/g) || []).length;
  n.data = n.data.replace(/​/g, '');
  const r = document.createRange();
  r.setStart(n, Math.max(0, off - before));
  r.collapse(true);
  setRange(r);
}
function insertStyledCaret(styles) {
  const sel = getSelection();
  if (!sel.rangeCount) return;
  const r = sel.getRangeAt(0);
  const span = document.createElement('span');
  Object.assign(span.style, styles);
  span.textContent = '​';
  r.insertNode(span);
  const nr = document.createRange();
  nr.setStart(span.firstChild, 1);
  nr.collapse(true);
  setRange(nr);
}

function applyFontSize(pt) {
  pt = Math.round(parseFloat(pt) * 2) / 2;
  if (!pt || pt < 1) return;
  pt = clamp(pt, 1, 1638);
  restoreSelection();
  const sel = getSelection();
  if (!sel.rangeCount) return;
  if (sel.isCollapsed) { insertStyledCaret({ fontSize: `${pt}pt` }); updateUi(); return; }
  document.execCommand('styleWithCSS', false, false);
  document.execCommand('fontSize', false, '7');
  document.execCommand('styleWithCSS', false, true);
  let first = null, last = null;
  for (const f of $$('font[size="7"]', editor)) {
    const s = document.createElement('span');
    s.style.fontSize = `${pt}pt`;
    for (const a of f.getAttributeNames()) if (a !== 'size' && a !== 'style') s.setAttribute(a, f.getAttribute(a));
    if (f.getAttribute('style')) s.style.cssText += f.getAttribute('style');
    s.style.fontSize = `${pt}pt`;
    while (f.firstChild) s.appendChild(f.firstChild);
    f.replaceWith(s);
    for (const e of $$('[style]', s)) {
      if (e.style.fontSize) { e.style.fontSize = ''; if (!e.getAttribute('style')) e.removeAttribute('style'); }
    }
    first = first || s;
    last = s;
  }
  if (first) {
    const r = document.createRange();
    r.setStartBefore(first);
    r.setEndAfter(last);
    setRange(r);
  }
  markDirty();
  updateUi();
}

function currentFontSize() {
  const sel = getSelection();
  const el = sel.rangeCount && inEditor(sel.anchorNode) ? styleAnchor() : (state.savedRange ? elOf(state.savedRange.startContainer) : null);
  return el ? Math.round(parseFloat(getComputedStyle(el).fontSize) * 0.75 * 2) / 2 : (parseFloat($('#font-size').value) || 11);
}
function stepFont(dir) {
  const cur = currentFontSize();
  const next = dir > 0 ? (FONT_SIZES.find(s => s > cur) || cur + 12) : ([...FONT_SIZES].reverse().find(s => s < cur) || Math.max(1, cur - 1));
  applyFontSize(next);
}

function clearFormatting() {
  restoreSelection();
  document.execCommand('removeFormat');
  const blocks = getSelectedBlocks();
  for (const b of blocks) {
    if (!b.matches('li')) {
      b.removeAttribute('style');
      b.removeAttribute('class');
    }
  }
  if (blocks.some(b => /^(H\d|BLOCKQUOTE|PRE)$/.test(b.tagName))) document.execCommand('formatBlock', false, 'p');
  markDirty();
  updateUi();
}

function applyFore(c) {
  if (c) { state.lastFore = c; $('#fore-bar').style.background = c; }
  exec('foreColor', c || '#000000');
}
function applyHilite(c) {
  if (c) { state.lastHilite = c; $('#hilite-bar').style.background = c; }
  restoreSelection();
  document.execCommand('hiliteColor', false, c || 'transparent');
  if (!c) for (const s of $$('span', editor)) if (s.style.backgroundColor === 'transparent') { s.style.backgroundColor = ''; if (!s.getAttribute('style')) unwrap(s); }
  markDirty();
}
function applyShade(c) {
  if (c) { state.lastShade = c; $('#shade-bar').style.background = c; }
  restoreSelection();
  for (const b of getSelectedBlocks()) b.style.backgroundColor = c || '';
  markDirty();
}
function unwrap(el) {
  const p = el.parentNode;
  while (el.firstChild) p.insertBefore(el.firstChild, el);
  el.remove();
}

function textSegments(range) {
  const segs = [];
  const w = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
  let n;
  while ((n = w.nextNode())) {
    if (!range.intersectsNode(n)) continue;
    const s = n === range.startContainer ? range.startOffset : 0;
    const e = n === range.endContainer ? range.endOffset : n.length;
    if (e > s) segs.push({ node: n, s, e });
  }
  return segs;
}
function changeCase(mode) {
  restoreSelection();
  const sel = getSelection();
  if (!sel.rangeCount || sel.isCollapsed) { toast('Select some text first.'); return; }
  const range = sel.getRangeAt(0);
  const segs = textSegments(range);
  const full = segs.map(x => x.node.data.slice(x.s, x.e)).join('');
  const lower = full.toLocaleLowerCase();
  const out = {
    upper: () => full.toLocaleUpperCase(),
    lower: () => lower,
    title: () => lower.replace(/(^|[^\p{L}\p{N}'’])(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase()),
    sentence: () => lower.replace(/(^\s*|[.!?]\s+|\n\s*)(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase()),
    toggle: () => [...full].map(c => (c === c.toLocaleUpperCase() ? c.toLocaleLowerCase() : c.toLocaleUpperCase())).join(''),
  }[mode]();
  if (out.length !== full.length) { document.execCommand('insertText', false, out); markDirty(); return; }
  let i = 0;
  for (const x of segs) {
    const len = x.e - x.s;
    x.node.data = x.node.data.slice(0, x.s) + out.slice(i, i + len) + x.node.data.slice(x.e);
    i += len;
  }
  const r = document.createRange();
  r.setStart(segs[0].node, segs[0].s);
  r.setEnd(segs[segs.length - 1].node, segs[segs.length - 1].e);
  setRange(r);
  markDirty();
}

// Format Painter
function toggleFormatPainter() {
  if (state.painter) { state.painter = null; document.body.classList.remove('painter'); updateToggles(); return; }
  restoreSelection();
  const el = styleAnchor();
  const cs = getComputedStyle(el);
  let bg = null;
  for (let e = el; e && e !== editor; e = e.parentElement) {
    if (e.matches('span,mark,font') && e.style.backgroundColor) { bg = e.style.backgroundColor; break; }
  }
  state.painter = {
    bold: document.queryCommandState('bold'),
    italic: document.queryCommandState('italic'),
    underline: document.queryCommandState('underline'),
    strike: document.queryCommandState('strikeThrough'),
    fontName: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(),
    fontSize: Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2,
    color: cs.color,
    bg,
  };
  document.body.classList.add('painter');
  updateToggles();
  toast('Now select the text you want to apply this formatting to.');
}
function applyPainter() {
  const p = state.painter;
  state.painter = null;
  document.body.classList.remove('painter');
  updateToggles();
  flushHistory();
  historyBusy++;
  document.execCommand('removeFormat');
  document.execCommand('styleWithCSS', false, false);
  for (const [cmd, on] of [['bold', p.bold], ['italic', p.italic], ['underline', p.underline], ['strikeThrough', p.strike]]) {
    if (document.queryCommandState(cmd) !== on) document.execCommand(cmd);
  }
  document.execCommand('styleWithCSS', false, true);
  setTimeout(() => { historyBusy--; if (!historyBusy) commitHistory(); }, 0);
  document.execCommand('fontName', false, p.fontName);
  document.execCommand('foreColor', false, p.color);
  if (p.bg) document.execCommand('hiliteColor', false, p.bg);
  applyFontSize(p.fontSize);
}

// ---------------------------------------------------------------- paragraph formatting
function indent(dir) {
  restoreSelection();
  if (closestEl(getSelection().anchorNode, 'li')) {
    listIndent(dir);
    return;
  } else {
    for (const b of getSelectedBlocks()) {
      const cur = parseFloat(getComputedStyle(b).marginLeft) || 0;
      const next = Math.max(0, Math.round((cur + dir * 48) / 48) * 48);
      b.style.marginLeft = next ? `${next}px` : '';
      if (!b.getAttribute('style')) b.removeAttribute('style');
    }
  }
  markDirty();
}
function setLineSpacing(v) {
  restoreSelection();
  for (const b of getSelectedBlocks()) b.style.lineHeight = v;
  markDirty();
}
function toggleSpaceAfter() {
  restoreSelection();
  const blocks = getSelectedBlocks();
  const hasSpace = blocks.some(b => parseFloat(getComputedStyle(b).marginBottom) > 0);
  for (const b of blocks) b.style.marginBottom = hasSpace ? '0' : '8pt';
  markDirty();
}
function toggleSpaceBefore() {
  restoreSelection();
  const blocks = getSelectedBlocks();
  const hasSpace = blocks.some(b => parseFloat(b.style.marginTop) > 0);
  for (const b of blocks) b.style.marginTop = hasSpace ? '' : '12pt';
  markDirty();
}
function applyStyle(name) {
  const tag = STYLE_TAGS[name];
  if (!tag) return;
  restoreSelection();
  document.execCommand('formatBlock', false, tag);
  for (const b of getSelectedBlocks()) {
    b.classList.toggle('doc-title', name === 'title');
    if (!b.className) b.removeAttribute('class');
  }
  markDirty();
  updateUi();
}
// Chromium's insertUnorderedList nests the list inside the paragraph (<p><ul>…) and
// moves the caret, so lists are built here instead.
function rememberRange() {
  const sel = getSelection();
  if (!sel.rangeCount) return null;
  const r = sel.getRangeAt(0);
  return { sc: r.startContainer, so: r.startOffset, ec: r.endContainer, eo: r.endOffset };
}
function reapplyRange(saved, swap = new Map()) {
  if (!saved) return;
  const fix = (node, off) => {
    const n = swap.get(node) || node;
    if (!editor.contains(n)) return null;
    return [n, Math.min(off, n.nodeType === 3 ? n.length : n.childNodes.length)];
  };
  const s = fix(saved.sc, saved.so), e = fix(saved.ec, saved.eo);
  if (!s || !e) return;
  const r = document.createRange();
  try { r.setStart(...s); r.setEnd(...e); setRange(r); } catch { /* keep current selection */ }
}
function moveChildren(from, to) { while (from.firstChild) to.appendChild(from.firstChild); }
function listItemsOf(blocks) {
  return blocks.map(b => (b.matches('li') ? b : b.closest('li'))).filter((li, i, a) => li && inEditor(li) && a.indexOf(li) === i);
}
function toggleList(tag, styleType) {
  restoreSelection();
  const saved = rememberRange();
  const blocks = getSelectedBlocks();
  if (!blocks.length) return;
  const swap = new Map();
  const items = listItemsOf(blocks);
  const allInLists = items.length && blocks.every(b => b.matches('li') || b.closest('li'));

  if (allInLists && items.every(li => li.parentElement.tagName.toLowerCase() === tag) && !styleType) {
    // remove the list formatting from these items
    for (const li of items) {
      const list = li.parentElement;
      const after = [];
      for (let n = li.nextElementSibling; n; n = n.nextElementSibling) after.push(n);
      if (after.length) {
        const tail = list.cloneNode(false);
        after.forEach(n => tail.appendChild(n));
        list.after(tail);
      }
      const p = document.createElement('p');
      if (li.getAttribute('style')) p.setAttribute('style', li.getAttribute('style'));
      moveChildren(li, p);
      if (!p.firstChild) p.innerHTML = '<br>';
      list.after(p);
      swap.set(li, p);
      li.remove();
      if (!list.children.length) list.remove();
    }
  } else if (allInLists) {
    // switch bullets <-> numbering (or apply a new marker style) for the lists involved
    const lists = [...new Set(items.map(li => li.parentElement))];
    for (const list of lists) {
      let target = list;
      if (list.tagName.toLowerCase() !== tag) {
        target = document.createElement(tag);
        for (const a of list.getAttributeNames()) if (a !== 'style') target.setAttribute(a, list.getAttribute(a));
        moveChildren(list, target);
        list.replaceWith(target);
        swap.set(list, target);
      }
      target.style.listStyleType = styleType || '';
      if (!target.getAttribute('style')) target.removeAttribute('style');
    }
  } else {
    // turn the selected paragraphs into one list
    const list = document.createElement(tag);
    if (styleType) list.style.listStyleType = styleType;
    const first = blocks[0];
    const cell = first.closest('td,th');
    first.before(list);
    for (const b of blocks) {
      if (b.closest('td,th') !== cell) continue;
      const li = document.createElement('li');
      const align = b.style.textAlign;
      if (align) li.style.textAlign = align;
      if (b.matches('li')) { moveChildren(b, li); b.remove(); }
      else {
        moveChildren(b, li);
        b.remove();
      }
      if (!li.firstChild) li.innerHTML = '<br>';
      swap.set(b, li);
      list.appendChild(li);
    }
    // join with a neighbouring list of the same kind
    const prev = list.previousElementSibling, next = list.nextElementSibling;
    if (prev && prev.tagName === list.tagName && (prev.getAttribute('style') || '') === (list.getAttribute('style') || '')) { moveChildren(list, prev); list.remove(); if (next && next.tagName === prev.tagName) { moveChildren(next, prev); next.remove(); } }
    else if (next && next.tagName === list.tagName && (next.getAttribute('style') || '') === (list.getAttribute('style') || '')) { while (list.lastChild) next.prepend(list.lastChild); list.remove(); }
  }
  reapplyRange(saved, swap);
  markDirty();
  updateUi();
}
function setListStyle(kind, type) {
  toggleList(kind, type);
}
function unlistItem(li, swap) {
  const list = li.parentElement;
  const after = [];
  for (let n = li.nextElementSibling; n; n = n.nextElementSibling) after.push(n);
  if (after.length) {
    const tail = list.cloneNode(false);
    after.forEach(n => tail.appendChild(n));
    list.after(tail);
  }
  const p = document.createElement('p');
  if (li.getAttribute('style')) p.setAttribute('style', li.getAttribute('style'));
  // a nested list inside the item stays a list, after the new paragraph
  const nested = [...li.children].filter(c => c.matches('ul,ol'));
  nested.forEach(n => n.remove());
  moveChildren(li, p);
  if (!p.firstChild) p.innerHTML = '<br>';
  list.after(p, ...nested);
  if (swap) swap.set(li, p);
  li.remove();
  if (!list.children.length) list.remove();
  return p;
}
// Tab / Increase Indent inside a list: nest the item under the one above it (Word style).
function listIndent(dir) {
  restoreSelection();
  const saved = rememberRange();
  const items = listItemsOf(getSelectedBlocks());
  if (!items.length) return false;
  const swap = new Map();
  if (dir > 0) {
    for (const li of items) {
      const prev = li.previousElementSibling;
      if (!prev || prev.tagName !== 'LI') continue;
      let sub = prev.lastElementChild;
      if (!sub || !sub.matches('ul,ol')) {
        sub = document.createElement(li.parentElement.tagName.toLowerCase());
        prev.appendChild(sub);
      }
      sub.appendChild(li);
      // children nested under this item move up a level with it
      const own = [...li.children].filter(c => c.matches('ul,ol'));
      for (const o of own) { while (o.firstChild) sub.appendChild(o.firstChild); o.remove(); }
    }
  } else {
    for (const li of [...items].reverse()) {
      const list = li.parentElement;
      const parentLi = list.parentElement && list.parentElement.matches('li') ? list.parentElement : null;
      if (!parentLi) { unlistItem(li, swap); continue; }
      const after = [];
      for (let n = li.nextElementSibling; n; n = n.nextElementSibling) after.push(n);
      if (after.length) {
        const sub = list.cloneNode(false);
        after.forEach(n => sub.appendChild(n));
        li.appendChild(sub);
      }
      parentLi.after(li);
      if (!list.children.length) list.remove();
    }
  }
  reapplyRange(saved, swap);
  markDirty();
  updateUi();
  return true;
}
// Chromium wraps text in <span style="font-size:…"> copies of inherited styles when it
// merges blocks; drop declarations that change nothing and unwrap empty spans.
function cleanRedundantSpans(root) {
  if (!root) return;
  const saved = rememberRange();
  let changed = false;
  for (const s of [...root.querySelectorAll('span[style]')]) {
    if (s.textContent.includes('​') || !s.parentElement) continue;
    const cs = getComputedStyle(s), ps = getComputedStyle(s.parentElement);
    for (const prop of [...s.style]) {
      if (/^(font-size|font-family|color|font-weight|font-style|line-height|letter-spacing)$/.test(prop) && cs.getPropertyValue(prop) === ps.getPropertyValue(prop)) {
        s.style.removeProperty(prop);
        changed = true;
      }
    }
    if (!s.style.length && [...s.attributes].every(a => a.name === 'style')) { unwrap(s); changed = true; }
  }
  if (changed) reapplyRange(saved);
}
function caretAtStartOf(el) {
  const sel = getSelection();
  if (!sel.rangeCount || !sel.isCollapsed) return false;
  const r = document.createRange();
  r.setStart(el, 0);
  r.setEnd(sel.anchorNode, sel.anchorOffset);
  const frag = r.cloneContents();
  return !frag.textContent.replace(/​/g, '') && !frag.querySelector('img,table,hr');
}
function inList(tag) {
  const li = closestEl(getSelection().anchorNode, 'li');
  return !!li && li.parentElement.tagName.toLowerCase() === tag;
}

// ---------------------------------------------------------------- insert
function insertPageBreak() {
  restoreSelection();
  // inside a list or table the break goes after it, as in Word
  const outer = closestEl(getSelection().anchorNode, 'table') || (() => {
    let li = closestEl(getSelection().anchorNode, 'li');
    let list = li && li.parentElement;
    while (list && list.parentElement && list.parentElement !== editor) list = list.parentElement;
    return list && list.parentElement === editor ? list : null;
  })();
  if (outer) {
    const pb = document.createElement('div');
    pb.className = 'page-break';
    pb.contentEditable = 'false';
    const p = document.createElement('p');
    p.innerHTML = '<br>';
    outer.after(pb, p);
    placeCaret(p, true);
  } else {
    document.execCommand('insertHTML', false, '<div class="page-break" contenteditable="false"></div><p><br></p>');
  }
  markDirty();
}
async function insertPicture() {
  const r = await api.pickImage();
  if (!r) return;
  if (r.error) { await alertBox('Insert Picture', r.error); return; }
  insertImage(r.url);
}
function insertImage(url) {
  restoreSelection();
  document.execCommand('insertHTML', false, `<img src="${esc(url)}" alt="">`);
  markDirty();
}
function insertTable(rows, cols) {
  restoreSelection();
  const row = `<tr>${'<td><br></td>'.repeat(cols)}</tr>`;
  document.execCommand('insertHTML', false, `<table><tbody>${row.repeat(rows)}</tbody></table><p><br></p>`);
  // put the caret in the first cell of the new table
  const sel = getSelection();
  const tables = $$('table', editor);
  const t = tables.find(tb => { const n = tb.nextElementSibling; return n && sel.anchorNode && (n === sel.anchorNode || n.contains(sel.anchorNode)); }) || tables[tables.length - 1];
  if (t && t.rows[0]) placeCaret(t.rows[0].cells[0], true);
  markDirty();
}
async function insertLink() {
  restoreSelection();
  const sel = getSelection();
  const text = sel.toString();
  const existing = closestEl(sel.anchorNode, 'a');
  const range = sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
  const r = await askForm(existing ? 'Edit Link' : 'Insert Link', [
    { id: 'text', label: 'Text to display', value: existing ? existing.textContent : text },
    { id: 'url', label: 'Address (web page or e-mail)', value: existing ? existing.getAttribute('href') : '', placeholder: 'https://example.com' },
  ], existing ? 'Update' : 'Insert');
  if (!r) return;
  let url = r.url.trim();
  if (!url) return;
  if (!/^(https?:|mailto:|#)/i.test(url)) url = /^[^\s/@]+@[^\s/@]+\.[^\s/@]+$/.test(url) ? `mailto:${url}` : `https://${url}`;
  editor.focus({ preventScroll: true });
  if (range) setRange(range);
  if (existing) {
    existing.setAttribute('href', url);
    if (r.text && r.text !== existing.textContent) existing.textContent = r.text;
  } else if (!text || r.text !== text) {
    document.execCommand('insertHTML', false, `<a href="${esc(url)}">${esc(r.text || url)}</a>&nbsp;`);
  } else {
    document.execCommand('createLink', false, url);
  }
  markDirty();
}
function insertText(t) {
  restoreSelection();
  document.execCommand('insertText', false, t);
}
async function editHeaderFooter() {
  const p = state.page;
  const r = await askForm('Header & Footer', [
    { id: 'header', label: 'Header text (top of every page)', value: p.header },
    { id: 'footer', label: 'Footer text (bottom of every page)', value: p.footer },
    { id: 'pageNumbers', label: 'Show page numbers in the footer', type: 'checkbox', value: p.pageNumbers },
  ], 'Apply', 'Headers and footers appear on printed pages, in PDF exports and in saved .docx files.');
  if (!r) return;
  Object.assign(p, { header: r.header.trim(), footer: r.footer.trim(), pageNumbers: r.pageNumbers });
  applyPage();
  markDirty();
}

// ---------------------------------------------------------------- tables
function tableOp(op) {
  restoreSelection();
  const cell = closestEl(getSelection().anchorNode, 'td,th');
  if (!cell) { toast('Click inside a table first.'); return; }
  const row = cell.parentElement;
  const table = cell.closest('table');
  const idx = cell.cellIndex;
  const blank = tag => { const c = document.createElement(tag); c.innerHTML = '<br>'; return c; };
  switch (op) {
    case 'rowAbove':
    case 'rowBelow': {
      const nr = document.createElement('tr');
      for (const c of row.cells) { const n = blank(c.tagName === 'TH' && op === 'rowAbove' ? 'th' : 'td'); if (c.colSpan > 1) n.colSpan = c.colSpan; nr.appendChild(n); }
      if (op === 'rowAbove') row.before(nr); else row.after(nr);
      placeCaret(nr.cells[Math.min(idx, nr.cells.length - 1)]);
      break;
    }
    case 'colLeft':
    case 'colRight':
      for (const tr of table.rows) {
        const ref = tr.cells[Math.min(idx, tr.cells.length - 1)];
        if (!ref) continue;
        const n = blank(ref.tagName.toLowerCase());
        if (op === 'colLeft') ref.before(n); else ref.after(n);
      }
      for (const c of $$('td,th', table)) c.style.width = '';
      break;
    case 'delRow':
      if (table.rows.length <= 1) { removeTable(table); break; }
      { const next = row.nextElementSibling || row.previousElementSibling; row.remove(); if (next) placeCaret(next.cells[0]); }
      break;
    case 'delCol':
      if (row.cells.length <= 1) { removeTable(table); break; }
      for (const tr of [...table.rows]) if (tr.cells[idx]) tr.cells[idx].remove();
      placeCaret(row.cells[Math.max(0, idx - 1)]);
      break;
    case 'delTable':
      removeTable(table);
      break;
    case 'headerRow': {
      const first = table.rows[0];
      const toTh = first.cells[0].tagName === 'TD';
      for (const c of [...first.cells]) {
        const n = document.createElement(toTh ? 'th' : 'td');
        n.innerHTML = c.innerHTML;
        if (c.colSpan > 1) n.colSpan = c.colSpan;
        c.replaceWith(n);
      }
      break;
    }
  }
  markDirty();
  updateUi();
}
function removeTable(table) {
  const p = document.createElement('p');
  p.innerHTML = '<br>';
  table.replaceWith(p);
  placeCaret(p);
}

// ---------------------------------------------------------------- images
function selectImage(img) {
  for (const i of $$('img.nw-selected', editor)) { i.classList.remove('nw-selected'); if (!i.className) i.removeAttribute('class'); }
  state.img = img;
  const bar = $('#imgbar');
  if (!img) { bar.hidden = true; $('#img-handles').hidden = true; return; }
  img.classList.add('nw-selected');
  const r = document.createRange();
  r.selectNode(img);
  setRange(r);
  bar.hidden = false;
  $('#img-handles').hidden = false;
  positionImgBar();
}
function positionImgBar() {
  const img = state.img;
  const bar = $('#imgbar');
  if (!img || bar.hidden) return;
  if (!editor.contains(img)) { selectImage(null); return; }
  const ws = $('#workspace').getBoundingClientRect();
  const r = img.getBoundingClientRect();
  const hd = $('#img-handles');
  Object.assign(hd.style, { left: `${r.left - ws.left}px`, top: `${r.top - ws.top}px`, width: `${r.width}px`, height: `${r.height}px` });
  let top = r.top - ws.top - bar.offsetHeight - 8;
  if (top < 6) top = Math.min(r.bottom - ws.top + 8, ws.height - bar.offsetHeight - 6);
  bar.style.top = `${Math.max(6, top)}px`;
  bar.style.left = `${clamp(r.left - ws.left, 6, ws.width - bar.offsetWidth - 6)}px`;
}
// drag a corner handle to resize the selected picture (keeps its proportions)
$('#img-handles').addEventListener('pointerdown', e => {
  const h = e.target.closest('.hd');
  const img = state.img;
  if (!h || !img) return;
  e.preventDefault();
  h.setPointerCapture(e.pointerId);
  const dir = h.dataset.h.endsWith('e') ? 1 : -1;
  const startX = e.clientX;
  const startW = img.offsetWidth;
  const cw = contentWidthPx();
  document.body.classList.add('resizing');
  const move = ev => {
    const w = clamp(startW + (dir * (ev.clientX - startX)) / state.zoom, 16, cw);
    img.style.width = `${Math.round(w)}px`;
    img.style.height = 'auto';
    img.removeAttribute('width');
    img.removeAttribute('height');
    positionImgBar();
  };
  const up = () => {
    h.removeEventListener('pointermove', move);
    h.removeEventListener('pointerup', up);
    h.removeEventListener('pointercancel', up);
    document.body.classList.remove('resizing');
    markDirty();
    selectImage(img);
  };
  h.addEventListener('pointermove', move);
  h.addEventListener('pointerup', up);
  h.addEventListener('pointercancel', up);
});
function contentWidthPx() {
  const m = state.page.margins;
  return (state.pw - m.left - m.right) * DPI / (state.page.columns || 1);
}
function imageAction(a) {
  const img = state.img;
  if (!img) return;
  const cw = contentWidthPx();
  const curW = img.offsetWidth || img.naturalWidth;
  const setW = w => { img.style.width = `${Math.round(clamp(w, 16, cw))}px`; img.style.height = 'auto'; img.removeAttribute('width'); img.removeAttribute('height'); };
  switch (a) {
    case 'smaller': setW(curW * 0.9); break;
    case 'larger': setW(curW * 1.1); break;
    case 'inline': img.style.float = ''; img.style.display = ''; img.style.margin = ''; break;
    case 'center': img.style.float = ''; img.style.display = 'block'; img.style.margin = '0 auto'; break;
    case 'floatLeft': img.style.display = ''; img.style.float = 'left'; img.style.margin = '0 12px 6px 0'; break;
    case 'floatRight': img.style.display = ''; img.style.float = 'right'; img.style.margin = '0 0 6px 12px'; break;
    case 'delete': { const r = document.createRange(); r.selectNode(img); setRange(r); editor.focus(); document.execCommand('delete'); selectImage(null); markDirty(); return; }
    default: setW(cw * (+a / 100));
  }
  markDirty();
  requestAnimationFrame(positionImgBar);
}

// ---------------------------------------------------------------- page layout
function applyPage() {
  const p = state.page;
  const paper = PAPER[p.size] || PAPER.letter;
  const land = p.orientation === 'landscape';
  state.pw = land ? paper.h : paper.w;
  state.ph = land ? paper.w : paper.h;
  const m = p.margins;
  page.style.width = `${state.pw * DPI}px`;
  page.style.padding = `${m.top * DPI}px ${m.right * DPI}px ${m.bottom * DPI}px ${m.left * DPI}px`;
  editor.style.minHeight = `${(state.ph - m.top - m.bottom) * DPI}px`;
  editor.style.columnCount = p.columns > 1 ? p.columns : '';
  $('#page-rule').textContent = `@page { size: ${state.pw}in ${state.ph}in; margin: ${m.top}in ${m.right}in ${m.bottom}in ${m.left}in; }`;
  state.sheetKey = '';
  layoutPages();
}

// ---- pagination
// The document is one editable flow. Paper sheets are drawn behind it with a gap between
// them, and any block that would cross the bottom margin is pushed (with extra top
// padding, never saved) to the top of the next sheet, so text never runs into the gap.
const PAGE_GAP = 24;
let layoutQueued = false;
let keepCaretVisible = false;
function scheduleLayout(fromTyping) {
  if (fromTyping) keepCaretVisible = true;
  if (layoutQueued) return;
  layoutQueued = true;
  requestAnimationFrame(() => { layoutQueued = false; layoutPages(); });
}
function geometry() {
  const m = state.page.margins;
  const PH = state.ph * DPI;
  const C = (state.ph - m.top - m.bottom) * DPI;
  const b = document.body.classList;
  const paged = !b.contains('web-layout') && !(state.page.columns > 1) && !b.contains('no-boundaries');
  return { m, PH, C, S: PH + PAGE_GAP, paged };
}
function layoutUnits() {
  const units = [];
  const walkList = list => {
    for (const li of list.children) {
      if (li.tagName !== 'LI') continue;
      units.push(li);
      for (const sub of li.children) if (sub.matches('ul,ol')) walkList(sub);
    }
  };
  for (const el of editor.children) {
    if (el.matches('ul,ol')) walkList(el);
    else units.push(el);
  }
  return units;
}
function clearPush(el) {
  el.removeAttribute('data-nw-push');
  el.style.removeProperty('--nw-push');
  el.style.removeProperty('--nw-base');
  if (!el.getAttribute('style')) el.removeAttribute('style');
}
function stripLayout(html) {
  return html.replace(/ data-nw-push=""/g, '').replace(/--nw-(push|base):\s*[^;"]*;?\s*/g, '').replace(/ style=""/g, '');
}
function layoutPages() {
  const g = geometry();
  for (const el of editor.querySelectorAll('[data-nw-push]')) clearPush(el);
  const er = editor.getBoundingClientRect();
  const scale = er.height / (editor.offsetHeight || 1) || 1;
  let n;
  if (g.paged) {
    const units = layoutUnits();
    const info = units.map(el => {
      const r = el.getBoundingClientRect();
      let bottom = r.bottom;
      if (el.tagName === 'LI') { const sub = [...el.children].find(c => c.matches('ul,ol')); if (sub) bottom = sub.getBoundingClientRect().top; }
      return { el, top: (r.top - er.top) / scale, h: Math.max(0, (bottom - r.top) / scale), heading: /^H[1-6]$/.test(el.tagName), pb: el.classList.contains('page-break') };
    });
    const pushes = new Map();
    const accAt = [];
    let acc = 0;
    let forceNext = false;
    const nextStart = top => (Math.floor(top / g.S) + 1) * g.S;
    for (let i = 0; i < info.length; i++) {
      const u = info[i];
      accAt[i] = acc;
      const top = u.top + acc;
      const pageTop = Math.floor(top / g.S) * g.S;
      let target = null;
      if (forceNext) { if (top > pageTop + 1) target = nextStart(top); forceNext = false; }
      else if (top > pageTop + g.C - 0.5) target = nextStart(top);          // starts in the bottom margin
      else if (top + u.h > pageTop + g.C + 0.5 && u.h <= g.C && top > pageTop + 1) {
        // doesn't fit: keep a heading together with the block after it, like Word
        const prev = info[i - 1];
        if (prev && prev.heading && !pushes.has(i - 1) && prev.top + accAt[i - 1] > pageTop + 1 && prev.top + accAt[i - 1] >= pageTop) {
          const ptop = prev.top + accAt[i - 1];
          acc = accAt[i - 1];
          const shift = nextStart(ptop) - ptop;
          pushes.set(i - 1, shift);
          acc += shift;
          i--; // re-evaluate this block after the heading moved
          continue;
        }
        target = nextStart(top);
      }
      if (target !== null) { pushes.set(i, target - top); acc += target - top; }
      if (u.pb) forceNext = true;
    }
    for (const [i, shift] of pushes) {
      const el = info[i].el;
      const cs = getComputedStyle(el);
      let push = shift, base;
      if (el.tagName === 'TABLE') {
        base = parseFloat(cs.marginTop) || 0;
        const prev = el.previousElementSibling;
        const a = prev ? parseFloat(getComputedStyle(prev).marginBottom) || 0 : 0;
        push = shift + Math.max(a, base) - base;   // margins collapse with the block above
      } else base = parseFloat(cs.paddingTop) || 0;
      el.setAttribute('data-nw-push', '');
      el.style.setProperty('--nw-base', `${base}px`);
      el.style.setProperty('--nw-push', `${Math.round(push * 100) / 100}px`);
    }
    const last = editor.lastElementChild;
    const bottom = last ? (last.getBoundingClientRect().bottom - er.top) / scale : 0;
    n = Math.max(1, Math.floor(Math.max(0, bottom - 1) / g.S) + 1);
    editor.style.minHeight = `${(n - 1) * g.S + g.C}px`;
    page.style.height = `${n * g.S - PAGE_GAP}px`;
    $('#boundaries').innerHTML = '';
  } else {
    editor.style.minHeight = `${g.C}px`;
    n = Math.max(1, Math.ceil((editor.offsetHeight - 2) / g.C));
    page.style.height = document.body.classList.contains('web-layout') ? '' : `${n * g.PH}px`;
    const lines = [];
    if (!document.body.classList.contains('no-boundaries')) for (let k = 1; k < n; k++) lines.push(`<div class="pb-line" style="top:${g.m.top * DPI + k * g.C}px"><span>Page ${k + 1}</span></div>`);
    $('#boundaries').innerHTML = lines.join('');
  }
  state.pageCount = n;
  renderSheets(n, g);
  updatePageStatus();
  positionImgBar();
  if (keepCaretVisible) { keepCaretVisible = false; revealCaret(); }
  for (const h of hooks.layout) h();
}
function renderSheets(n, g) {
  const p = state.page;
  const key = [n, g.paged, g.PH, p.header, p.footer, p.pageNumbers, g.m.top, g.m.bottom, ...hooks.sheetKey.map(f => f())].join('|');
  if (key === state.sheetKey) return;
  state.sheetKey = key;
  const sheets = $('#sheets');
  const extra = k => hooks.sheet.map(f => f(k, n, g)).join('');
  if (!g.paged) {
    sheets.innerHTML = `<div class="sheet" style="top:0;height:${n * g.PH}px">${extra(0)}</div>`;
    return;
  }
  const hTop = Math.max(4, Math.min(0.5 * DPI, (g.m.top * DPI) / 2) - 7);
  const fBot = Math.max(4, Math.min(0.5 * DPI, (g.m.bottom * DPI) / 2) - 7);
  const out = [];
  for (let k = 0; k < n; k++) {
    const foot = [p.footer, p.pageNumbers ? `Page ${k + 1} of ${n}` : ''].filter(Boolean).join(' · ');
    out.push(`<div class="sheet" style="top:${k * g.S}px;height:${g.PH}px">${p.header ? `<div class="hf" style="top:${hTop}px;padding:0 ${g.m.right * DPI}px 0 ${g.m.left * DPI}px">${esc(p.header)}</div>` : ''}${foot ? `<div class="hf" style="bottom:${fBot}px;padding:0 ${g.m.right * DPI}px 0 ${g.m.left * DPI}px">${esc(foot)}</div>` : ''}${extra(k)}</div>`);
  }
  sheets.innerHTML = out.join('');
}
function revealCaret() {
  const sel = getSelection();
  if (!sel.rangeCount || !inEditor(sel.anchorNode) || document.activeElement !== editor) return;
  let rect = sel.getRangeAt(0).getBoundingClientRect();
  if (!rect.height) rect = anchorElement().getBoundingClientRect();
  const sr = scroller.getBoundingClientRect();
  if (rect.bottom > sr.bottom - 24) scroller.scrollTop += rect.bottom - sr.bottom + 60;
  else if (rect.top < sr.top + 8) scroller.scrollTop -= sr.top - rect.top + 40;
}
function currentPage() {
  const sel = getSelection();
  if (!sel.rangeCount || !inEditor(sel.anchorNode)) return state.lastPage || 1;
  let rect = sel.getRangeAt(0).getBoundingClientRect();
  if (!rect.height) rect = anchorElement().getBoundingClientRect();
  const er = editor.getBoundingClientRect();
  const scale = er.height / (editor.offsetHeight || 1) || 1;
  const g = geometry();
  const y = (rect.top - er.top) / scale;
  state.lastPage = clamp(Math.floor(y / (g.paged ? g.S : g.C)) + 1, 1, state.pageCount);
  return state.lastPage;
}
function updatePageStatus() {
  $('#st-page').textContent = `Page ${currentPage()} of ${state.pageCount}`;
}
function setPage(changes) {
  Object.assign(state.page, changes);
  applyPage();
  markDirty();
}
async function customMargins() {
  const m = state.page.margins;
  const f = (id, label) => ({ id, label, value: m[id], type: 'number', attrs: 'step="0.05" min="0" max="5"' });
  const r = await askForm('Custom Margins (inches)', [f('top', 'Top'), f('bottom', 'Bottom'), f('left', 'Left'), f('right', 'Right')], 'Apply', null, true);
  if (!r) return;
  const v = k => clamp(parseFloat(r[k]) || 0, 0, 5);
  const next = { top: v('top'), bottom: v('bottom'), left: v('left'), right: v('right') };
  if (next.left + next.right > state.pw - 1 || next.top + next.bottom > state.ph - 1) { await alertBox('Custom Margins', 'Those margins leave no room for text. Try smaller margins.'); return; }
  setPage({ margins: next });
}

// ---------------------------------------------------------------- zoom & view
function setZoom(z) {
  z = clamp(Math.round(z * 100) / 100, 0.25, 3);
  const before = scroller.scrollTop / (scroller.scrollHeight || 1);
  state.zoom = z;
  zoomer.style.zoom = z;
  $('#zoom-range').value = Math.round(z * 100);
  $('#st-zoom').textContent = `${Math.round(z * 100)}%`;
  scroller.scrollTop = before * scroller.scrollHeight;
  positionImgBar();
}
function fitPageWidth() { setWebLayout(false); setZoom((scroller.clientWidth - 64) / (state.pw * DPI)); }
function fitOnePage() { setWebLayout(false); setZoom(Math.min((scroller.clientHeight - 60) / (state.ph * DPI), (scroller.clientWidth - 64) / (state.pw * DPI))); }
async function zoomDialog() {
  const r = await askForm('Zoom', [{ id: 'z', label: 'Percent (25–300)', value: Math.round(state.zoom * 100), type: 'number', attrs: 'min="25" max="300" step="5"' }], 'OK');
  if (r) setZoom((parseFloat(r.z) || 100) / 100);
}
function setWebLayout(on) { document.body.classList.toggle('web-layout', on); updateToggles(); scheduleLayout(); }
function setFocusMode(on) {
  document.body.classList.toggle('focus-mode', on);
  if (on) toast('Focus mode. Press Esc to exit.');
  updateToggles();
}
function setDark(on) {
  document.body.classList.toggle('dark', on);
  api.setSetting('dark', on);
  updateToggles();
}
function toggleSpell() {
  editor.spellcheck = !editor.spellcheck;
  // force Chromium to re-check / clear the squiggles
  const html = stripLayout(editor.innerHTML);
  const r = state.savedRange;
  editor.blur();
  editor.innerHTML = html;
  state.savedRange = null;
  editor.focus();
  if (!r) placeCaret(editor.firstElementChild || editor);
  updateToggles();
  toast(editor.spellcheck ? 'Spell checking is on. Right-click a red-underlined word for suggestions.' : 'Spell checking is off');
}
async function gotoPage() {
  const r = await askForm('Go To Page', [{ id: 'n', label: `Page number (1–${state.pageCount})`, value: currentPage(), type: 'number', attrs: `min="1" max="${state.pageCount}"` }], 'Go');
  if (!r) return;
  const n = clamp(parseInt(r.n, 10) || 1, 1, state.pageCount);
  const g = geometry();
  const y = (g.paged ? (n - 1) * g.S : g.m.top * DPI + (n - 1) * g.C) * state.zoom;
  scroller.scrollTop = page.getBoundingClientRect().top - scroller.getBoundingClientRect().top + scroller.scrollTop + y - 16;
  // put the caret at the start of that page
  const target = layoutUnits().find(u => (u.getBoundingClientRect().top - editor.getBoundingClientRect().top) / state.zoom >= (g.paged ? (n - 1) * g.S : (n - 1) * g.C) - 2);
  if (target) { editor.focus({ preventScroll: true }); placeCaret(target, true); }
}
function updateToggles() {
  const b = document.body.classList;
  const set = (cmd, on) => $$(`[data-cmd="${cmd}"]`).forEach(e => e.classList.toggle('on', !!on));
  set('printLayout', !b.contains('web-layout'));
  set('webLayout', b.contains('web-layout'));
  set('focus', b.contains('focus-mode'));
  set('boundaries', !b.contains('no-boundaries'));
  set('dark', b.contains('dark'));
  set('spell', editor.spellcheck);
  set('painter', !!state.painter);
  set('pageNumbers', state.page.pageNumbers);
  set('speak', state.speaking);
  $('#st-spell').textContent = `Spelling: ${editor.spellcheck ? 'On' : 'Off'}`;
}

// ---------------------------------------------------------------- toolbar state
let uiQueued = false;
function scheduleUi() {
  if (uiQueued) return;
  uiQueued = true;
  requestAnimationFrame(() => { uiQueued = false; updateUi(); });
}
function updateUi() {
  const sel = getSelection();
  if (!sel.rangeCount || !inEditor(sel.anchorNode)) return;
  // command states describe whatever has focus, so only read them from the document itself
  if (document.activeElement !== editor || $('#modal').open) return;
  const q = c => { try { return document.queryCommandState(c); } catch { return false; } };
  const map = { bold: 'bold', italic: 'italic', underline: 'underline', strike: 'strikeThrough', sub: 'subscript', sup: 'superscript', left: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', justify: 'justifyFull' };
  for (const [cmd, name] of Object.entries(map)) {
    const on = q(name);
    for (const b of $$(`[data-cmd="${cmd}"]`)) b.classList.toggle('on', on);
  }
  for (const b of $$('[data-cmd="bullets"]')) b.classList.toggle('on', inList('ul'));
  for (const b of $$('[data-cmd="numbers"]')) b.classList.toggle('on', inList('ol'));
  const el = styleAnchor();
  const cs = getComputedStyle(el);
  const ff = $('#font-family');
  if (document.activeElement !== ff) {
    const fam = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    if (fam && fam !== STRIKE_FONT) {
      if (![...ff.options].some(o => o.value === fam)) addFontOption(fam, true);
      ff.value = fam;
    }
  }
  const fs = $('#font-size');
  if (document.activeElement !== fs) fs.value = String(Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2);

  const block = closestBlock(el);
  let style = 'normal';
  if (block) {
    const t = block.tagName.toLowerCase();
    if (t === 'h1') style = block.classList.contains('doc-title') ? 'title' : 'h1';
    else if (t === 'h2' || t === 'h3') style = t;
    else if (t === 'blockquote') style = 'quote';
    else if (t === 'pre') style = 'code';
  }
  for (const c of $$('.style-card')) c.classList.toggle('on', c.dataset.arg === style);

  const inTable = !!closestEl(el, 'td,th');
  if (inTable !== state.inTable) {
    state.inTable = inTable;
    api.setCtx({ inTable });
    $('.table-tools').classList.toggle('disabled', !inTable);
  }
  updatePageStatus();
}

// ---------------------------------------------------------------- stats
let statsTimer = null;
function scheduleStats() {
  clearTimeout(statsTimer);
  statsTimer = setTimeout(updateStats, 250);
}
function countWords(t) { return (t.match(/\S+/g) || []).length; }
function updateStats() {
  const sel = getSelection();
  const total = countWords(editor.innerText);
  const selText = sel.rangeCount && inEditor(sel.anchorNode) && !sel.isCollapsed ? sel.toString() : '';
  $('#st-words').textContent = selText ? `${countWords(selText)} of ${total.toLocaleString()} words` : `${total.toLocaleString()} word${total === 1 ? '' : 's'}`;
}
async function showWordCount() {
  const sel = getSelection();
  const selText = sel.rangeCount && inEditor(sel.anchorNode) && !sel.isCollapsed ? sel.toString() : '';
  const text = selText || editor.innerText;
  const paras = $$(BLOCK_SEL, editor).filter(b => !b.querySelector(BLOCK_SEL) && b.textContent.trim()).length;
  const rows = [
    ['Pages', state.pageCount],
    ['Words', countWords(text)],
    ['Characters (no spaces)', text.replace(/\s/g, '').length],
    ['Characters (with spaces)', text.replace(/\n/g, '').length],
    ['Paragraphs', selText ? (selText.split(/\n+/).filter(s => s.trim()).length) : paras],
  ];
  await alertBox(selText ? 'Word Count (selection)' : 'Word Count', `<table class="stats">${rows.map(([a, b]) => `<tr><td>${a}</td><td>${Number(b).toLocaleString()}</td></tr>`).join('')}</table>`, true);
}

// ---------------------------------------------------------------- read aloud
function speak() {
  if (!window.speechSynthesis) { toast('Read Aloud is not available on this PC.'); return; }
  const sel = getSelection();
  const t = ((sel.rangeCount && inEditor(sel.anchorNode) && sel.toString().trim()) || editor.innerText).trim();
  if (!t) { toast('There is nothing to read.'); return; }
  speechSynthesis.cancel();
  const chunks = t.match(/[^.!?\n]+[.!?]*\s*/g) || [t];
  state.speaking = true;
  updateToggles();
  chunks.forEach((c, i) => {
    const u = new SpeechSynthesisUtterance(c);
    if (i === chunks.length - 1) u.onend = () => { state.speaking = false; updateToggles(); };
    u.onerror = () => { state.speaking = false; updateToggles(); };
    speechSynthesis.speak(u);
  });
}
function stopSpeak() {
  if (window.speechSynthesis) speechSynthesis.cancel();
  state.speaking = false;
  updateToggles();
}

// ---------------------------------------------------------------- find & replace
const findInput = $('#find-input');
const replaceInput = $('#replace-input');
function openFind(withReplace) {
  const bar = $('#findbar');
  bar.hidden = false;
  $('#replace-row').hidden = !withReplace;
  const sel = getSelection();
  if (sel.rangeCount && inEditor(sel.anchorNode) && !sel.isCollapsed) {
    const t = sel.toString();
    if (t.length < 100 && !t.includes('\n')) findInput.value = t;
  }
  findInput.focus();
  findInput.select();
  refreshFind(false);
}
function closeFind() {
  $('#findbar').hidden = true;
  if (window.CSS && CSS.highlights) CSS.highlights.clear();
  const cur = state.find.ranges[state.find.index];
  state.find = { ranges: [], index: -1 };
  editor.focus({ preventScroll: true });
  if (cur) setRange(cur);
  else if (state.savedRange) setRange(state.savedRange);
}
function findMatches() {
  const q = findInput.value;
  if (!q) return [];
  const mc = $('#find-case').checked;
  const whole = $('#find-word').checked;
  const needle = mc ? q : q.toLocaleLowerCase();
  const isWord = c => !!c && /[\p{L}\p{N}_]/u.test(c);
  const out = [];
  const w = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
  let n;
  while ((n = w.nextNode())) {
    const hay = mc ? n.data : n.data.toLocaleLowerCase();
    let i = 0;
    while ((i = hay.indexOf(needle, i)) !== -1) {
      if (!whole || (!isWord(n.data[i - 1]) && !isWord(n.data[i + q.length]))) {
        const r = document.createRange();
        r.setStart(n, i);
        r.setEnd(n, i + q.length);
        out.push(r);
      }
      i += Math.max(1, q.length);
    }
  }
  return out;
}
function paintFind() {
  if (!(window.CSS && CSS.highlights && window.Highlight)) return;
  const { ranges, index } = state.find;
  CSS.highlights.set('nw-find', new Highlight(...ranges));
  if (ranges[index]) CSS.highlights.set('nw-find-current', new Highlight(ranges[index]));
  else CSS.highlights.delete('nw-find-current');
  $('#find-count').textContent = findInput.value ? (ranges.length ? `${index + 1} of ${ranges.length}` : 'No results') : '';
}
function scrollRangeIntoView(r) {
  const rect = r.getBoundingClientRect();
  const sr = scroller.getBoundingClientRect();
  if (rect.top < sr.top + 60 || rect.bottom > sr.bottom - 40) scroller.scrollTop += rect.top - sr.top - sr.height / 2;
}
function pointCmp(r, node, offset) { const c = document.createRange(); c.setStart(r.startContainer, r.startOffset); return c.comparePoint(node, offset); }
function refreshFind(advance, dir = 1) {
  const ranges = findMatches();
  state.find.ranges = ranges;
  if (!ranges.length) { state.find.index = -1; paintFind(); return; }
  const ref = state.find.index >= 0 && !advance ? null : (state.savedRange || null);
  let idx;
  if (!advance) {
    // first match at or after the caret
    const anchor = state.savedRange;
    idx = anchor ? ranges.findIndex(r => pointCmp(r, anchor.startContainer, anchor.startOffset) <= 0) : 0;
    if (idx < 0) idx = 0;
  } else {
    const cur = ref;
    if (dir > 0) {
      idx = cur ? ranges.findIndex(r => pointCmp(r, cur.endContainer, cur.endOffset) <= 0) : 0;
      if (idx < 0) idx = 0;
    } else {
      idx = -1;
      if (cur) for (let i = ranges.length - 1; i >= 0; i--) {
        const end = document.createRange(); end.setStart(ranges[i].endContainer, ranges[i].endOffset);
        if (end.comparePoint(cur.startContainer, cur.startOffset) >= 0) { idx = i; break; }
      }
      if (idx < 0) idx = ranges.length - 1;
    }
  }
  state.find.index = idx;
  state.savedRange = ranges[idx].cloneRange();
  paintFind();
  scrollRangeIntoView(ranges[idx]);
}
function findStep(dir) {
  if (!state.find.ranges.length) { refreshFind(false); return; }
  refreshFind(true, dir);
}
function replaceOne() {
  const q = findInput.value;
  const cur = state.find.ranges[state.find.index];
  if (!q || !cur) { refreshFind(false); return; }
  editor.focus({ preventScroll: true });
  setRange(cur);
  document.execCommand('insertText', false, replaceInput.value);
  state.savedRange = getSelection().getRangeAt(0).cloneRange();
  findInput.focus();
  refreshFind(true, 1);
}
function replaceAll() {
  const q = findInput.value;
  if (!q) return;
  const ranges = findMatches();
  if (!ranges.length) { toast('No results found.'); return; }
  editor.focus({ preventScroll: true });
  for (let i = ranges.length - 1; i >= 0; i--) {
    setRange(ranges[i]);
    document.execCommand('insertText', false, replaceInput.value);
  }
  findInput.focus();
  refreshFind(false);
  toast(`Replaced ${ranges.length} item${ranges.length === 1 ? '' : 's'}.`);
}

// ---------------------------------------------------------------- menus
const HILITES = ['#ffff00', '#00ff00', '#00ffff', '#ff00ff', '#0000ff', '#ff0000', '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#c0c0c0', '#000000'];
const SYMBOLS = '© ® ™ § ¶ † ‡ • · … – — ‘ ’ “ ” « » ‹ › € £ ¥ ¢ $ ₹ ° ± × ÷ ≈ ≠ ≤ ≥ ∞ √ ∑ π µ Ω α β γ δ θ λ σ φ ½ ¼ ¾ ⅓ ² ³ ← → ↑ ↓ ↔ ⇒ ✓ ✔ ✗ ★ ☆ ♥ ☺ ☐ ☑ ♪ ¡ ¿'.split(' ');

const MENUS = {
  sizes: () => FONT_SIZES.map(s => item(String(s), () => applyFontSize(s), { check: currentFontSize() === s })),
  case: () => [
    item('Sentence case.', () => changeCase('sentence')),
    item('lowercase', () => changeCase('lower')),
    item('UPPERCASE', () => changeCase('upper')),
    item('Capitalize Each Word', () => changeCase('title')),
    item('tOGGLE cASE', () => changeCase('toggle')),
  ],
  fore: () => colorMenu(applyFore, 'Automatic'),
  hilite: () => colorMenu(applyHilite, 'No Color', HILITES, true),
  shade: () => colorMenu(applyShade, 'No Color'),
  bulletStyle: () => [
    item('● Filled circle', () => setListStyle('ul', 'disc')),
    item('○ Hollow circle', () => setListStyle('ul', 'circle')),
    item('■ Square', () => setListStyle('ul', 'square')),
  ],
  numberStyle: () => [
    item('1. 2. 3.', () => setListStyle('ol', 'decimal')),
    item('a. b. c.', () => setListStyle('ol', 'lower-alpha')),
    item('A. B. C.', () => setListStyle('ol', 'upper-alpha')),
    item('i. ii. iii.', () => setListStyle('ol', 'lower-roman')),
    item('I. II. III.', () => setListStyle('ol', 'upper-roman')),
  ],
  spacing: () => {
    const b = getSelectedBlocksSafe();
    const cur = b[0] ? b[0].style.lineHeight : '';
    return [
      { head: 'Line spacing' },
      ...['1.0', '1.15', '1.5', '2.0', '2.5', '3.0'].map(v => item(v, () => setLineSpacing(v === '1.0' ? '1' : v), { check: cur === v || (cur === '1' && v === '1.0') || (!cur && v === '1.15') })),
      '-',
      item(b[0] && parseFloat(b[0].style.marginTop) > 0 ? 'Remove Space Before Paragraph' : 'Add Space Before Paragraph', toggleSpaceBefore),
      item(b[0] && parseFloat(getComputedStyle(b[0]).marginBottom) > 0 ? 'Remove Space After Paragraph' : 'Add Space After Paragraph', toggleSpaceAfter),
    ];
  },
  table: () => ({
    build: el => {
      el.innerHTML = '<div class="tlabel">Insert Table</div><div class="tgrid"></div><hr><button class="item" data-custom>Insert Table…</button>';
      const g = $('.tgrid', el);
      for (let r = 0; r < 8; r++) for (let c = 0; c < 10; c++) { const d = document.createElement('div'); d.dataset.r = r + 1; d.dataset.c = c + 1; g.appendChild(d); }
      g.addEventListener('mouseover', e => {
        const t = e.target.closest('[data-r]');
        if (!t) return;
        for (const d of g.children) d.classList.toggle('hot', +d.dataset.r <= +t.dataset.r && +d.dataset.c <= +t.dataset.c);
        $('.tlabel', el).textContent = `${t.dataset.c} × ${t.dataset.r} Table`;
      });
      g.addEventListener('click', e => { const t = e.target.closest('[data-r]'); if (t) { closeMenu(); withHistory(() => insertTable(+t.dataset.r, +t.dataset.c)); } });
      $('[data-custom]', el).addEventListener('click', async () => {
        closeMenu();
        const r = await askForm('Insert Table', [
          { id: 'cols', label: 'Number of columns', value: 3, type: 'number', attrs: 'min="1" max="30"' },
          { id: 'rows', label: 'Number of rows', value: 3, type: 'number', attrs: 'min="1" max="200"' },
        ], 'Insert', null, true);
        if (r) withHistory(() => insertTable(clamp(parseInt(r.rows, 10) || 1, 1, 200), clamp(parseInt(r.cols, 10) || 1, 1, 30)));
      });
    },
  }),
  symbols: () => ({
    build: el => {
      el.innerHTML = `<div class="symgrid">${SYMBOLS.map(s => `<button data-s="${esc(s)}" title="${esc(s)}">${esc(s)}</button>`).join('')}</div>`;
      el.addEventListener('click', e => { const b = e.target.closest('[data-s]'); if (b) { closeMenu(); withHistory(() => insertText(b.dataset.s)); } });
    },
  }),
  date: () => {
    const d = new Date();
    const f = o => d.toLocaleString(undefined, o);
    return [
      f({ dateStyle: 'full' }),
      f({ dateStyle: 'long' }),
      f({ dateStyle: 'medium' }),
      f({ dateStyle: 'short' }),
      d.toISOString().slice(0, 10),
      f({ timeStyle: 'short' }),
      f({ dateStyle: 'long', timeStyle: 'short' }),
    ].map(s => item(s, () => insertText(s)));
  },
  margins: () => [
    ...Object.entries(MARGINS).map(([k, m]) => item(m.label, () => setPage({ margins: { top: m.top, bottom: m.bottom, left: m.left, right: m.right } }), {
      check: ['top', 'bottom', 'left', 'right'].every(s => state.page.margins[s] === m[s]),
      hint: `T ${m.top}" B ${m.bottom}" L ${m.left}" R ${m.right}"`,
    })),
    '-',
    item('Custom Margins…', customMargins),
  ],
  orientation: () => [
    item('Portrait', () => setPage({ orientation: 'portrait' }), { check: state.page.orientation === 'portrait' }),
    item('Landscape', () => setPage({ orientation: 'landscape' }), { check: state.page.orientation === 'landscape' }),
  ],
  size: () => Object.entries(PAPER).map(([k, p]) => item(p.label, () => setPage({ size: k }), { check: state.page.size === k, hint: p.desc })),
  columns: () => [1, 2, 3].map(n => item(['One', 'Two', 'Three'][n - 1], () => {
    setPage({ columns: n });
    if (n > 1) toast('Columns are shown side by side while you edit. Print Preview shows exactly how pages will look.', printPreview, 'Print Preview');
  }, { check: (state.page.columns || 1) === n })),
};
function getSelectedBlocksSafe() {
  const sel = getSelection();
  if (state.savedRange && (!sel.rangeCount || !inEditor(sel.anchorNode))) return [closestBlock(state.savedRange.startContainer)].filter(Boolean);
  return sel.rangeCount && inEditor(sel.anchorNode) ? [closestBlock(sel.anchorNode)].filter(Boolean) : [];
}
// ---------------------------------------------------------------- sanitizing
function sanitize(html) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  doc.querySelectorAll('script,style,meta,link,iframe,frame,frameset,object,embed,applet,form,input,button,select,textarea,title,base,noscript,template,svg,math,xml,video,audio,canvas').forEach(n => n.remove());
  const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_COMMENT);
  const comments = [];
  while (walker.nextNode()) comments.push(walker.currentNode);
  comments.forEach(c => c.remove());
  for (const el of doc.body.querySelectorAll('*')) {
    if (el.tagName.includes(':')) { el.replaceWith(...el.childNodes); continue; }
    for (const a of el.getAttributeNames()) {
      const v = el.getAttribute(a);
      const n = a.toLowerCase();
      if (n.startsWith('on') || n === 'contenteditable' || n === 'id' || n === 'name' || n === 'srcset' || n === 'formaction') el.removeAttribute(a);
      else if (n === 'href' && !/^(https?:|mailto:|#)/i.test(v.trim())) el.removeAttribute(a);
      else if (n === 'src' && !/^(data:image\/|https?:)/i.test(v.trim())) el.removeAttribute(a);
      else if (n === 'class') {
        const keep = v.split(/\s+/).filter(c => KEEP_CLASSES.has(c));
        if (keep.length) el.setAttribute('class', keep.join(' ')); else el.removeAttribute('class');
      } else if (n === 'style') {
        const decls = v.split(';').map(s => s.trim()).filter(s => s && !/^(mso-|position|z-index|behavior|-moz-binding)/i.test(s) && !/expression\s*\(|url\s*\(/i.test(s));
        if (decls.length) {
          el.setAttribute('style', decls.join('; '));
          const norm = el.style.cssText; // the browser's own spelling, so saved and live copies match
          if (norm) el.setAttribute('style', norm); else el.removeAttribute('style');
        } else el.removeAttribute('style');
      }
    }
    if (el.tagName === 'IMG' && !el.getAttribute('src')) el.remove();
  }
  for (const pb of doc.body.querySelectorAll('.page-break')) { pb.innerHTML = ''; pb.setAttribute('contenteditable', 'false'); }
  for (const h of hooks.sanitize) h(doc.body);
  return doc.body.innerHTML;
}

// ---------------------------------------------------------------- documents
function isPristine() {
  return !state.dirty && !state.path && !editor.innerText.trim() && !editor.querySelector('img,table,hr');
}
function updateTitle() {
  const n = state.name;
  document.title = `${state.dirty ? '• ' : ''}${n} - Nisar Office Studio`;
  const dn = $('#docname');
  dn.textContent = `${n}${state.dirty ? ' • Not saved' : ''}`;
  dn.title = state.path || 'Not saved yet';
}
function loadDoc(doc) {
  let html;
  if (doc.text != null) html = doc.text.split(/\r?\n/).map(l => `<p>${esc(l).replace(/\t/g, '&emsp;') || '<br>'}</p>`).join('');
  else html = sanitize(doc.html || '');
  editor.innerHTML = html.trim() ? html : '<p><br></p>';
  const p = doc.page && typeof doc.page === 'object' ? doc.page : null;
  state.page = { ...defaultPage(), ...(p || {}), margins: { ...MARGINS.normal, ...((p && p.margins) || {}) } };
  delete state.page.margins.label;
  if (!PAPER[state.page.size]) state.page.size = 'letter';
  state.path = doc.path || null;
  state.name = doc.name || state.name;
  state.dirty = false;
  selectImage(null);
  for (const h of hooks.load) h(doc.extra || {});
  applyPage();
  updateTitle();
  editor.focus({ preventScroll: true });
  placeCaret(editor.firstElementChild || editor, true);
  scroller.scrollTop = 0;
  resetHistory();
  scheduleStats();
  updateUi();
  if (doc.imported === 'pdf') {
    // converted from a PDF: not saved anywhere yet
    state.dirty = true;
    updateTitle();
    toast('Converted from PDF. Check the layout, then save it as a document.');
  } else if (doc.imported) toast('Opened. Some advanced Word formatting (like text boxes and comments) may look simpler here.');
}

function cleanClone() {
  const clone = editor.cloneNode(true);
  const w = document.createTreeWalker(clone, NodeFilter.SHOW_TEXT);
  let n;
  while ((n = w.nextNode())) if (n.data.includes('​')) n.data = n.data.replace(/​/g, '');
  for (const i of clone.querySelectorAll('img.nw-selected')) { i.classList.remove('nw-selected'); if (!i.className) i.removeAttribute('class'); }
  for (const el of clone.querySelectorAll('[data-nw-push]')) clearPush(el);
  for (const el of clone.querySelectorAll('.cmt-anchor.active, .tc-current')) { el.classList.remove('active', 'tc-current'); }
  return clone;
}
function buildExportHtml(extras) {
  const clone = cleanClone();
  for (const h of hooks.export) h(clone, extras);
  const liveImgs = $$('img', editor);
  clone.querySelectorAll('img').forEach((img, i) => {
    const live = liveImgs[i];
    if (live && live.offsetWidth) {
      img.style.width = `${live.offsetWidth}px`;
      img.style.height = `${live.offsetHeight}px`;
    }
    img.removeAttribute('width');
    img.removeAttribute('height');
  });
  const setDefault = (el, prop, val) => { if (!el.style[prop]) el.style[prop] = val; };
  for (const el of clone.querySelectorAll('h1,h2,h3,h4')) {
    const t = el.classList.contains('doc-title');
    setDefault(el, 'fontFamily', 'Calibri Light');
    setDefault(el, 'fontSize', t ? '28pt' : { H1: '16pt', H2: '13pt', H3: '12pt', H4: '11pt' }[el.tagName]);
    setDefault(el, 'color', t ? '#000000' : el.tagName === 'H3' ? '#1f3763' : '#2f5496');
  }
  for (const el of clone.querySelectorAll('.doc-subtitle')) { setDefault(el, 'color', '#5a5a5a'); setDefault(el, 'fontSize', '13pt'); }
  for (const el of clone.querySelectorAll('blockquote')) {
    const p = document.createElement('p');
    p.innerHTML = el.innerHTML;
    p.style.cssText = el.style.cssText;
    setDefault(p, 'marginLeft', '58px');
    setDefault(p, 'fontStyle', 'italic');
    setDefault(p, 'color', '#404040');
    el.replaceWith(p);
  }
  for (const el of clone.querySelectorAll('pre')) {
    const p = document.createElement('p');
    p.innerHTML = el.innerHTML.replace(/\n/g, '<br>');
    p.style.fontFamily = 'Consolas';
    p.style.fontSize = '10pt';
    el.replaceWith(p);
  }
  for (const el of clone.querySelectorAll('colgroup')) el.remove();
  for (const el of clone.querySelectorAll('td,th')) {
    el.style.width = '';
    setDefault(el, 'border', el.closest('table.no-borders') ? 'none' : '1px solid #8e8e8e');
    if (el.tagName === 'TH') setDefault(el, 'backgroundColor', '#f2f2f2');
  }
  for (const el of clone.querySelectorAll('[style]')) {
    if (/transparent|rgba\(0, 0, 0, 0\)/.test(el.style.backgroundColor)) el.style.backgroundColor = '';
    if (el.tagName !== 'TD' && el.tagName !== 'TH' && el.tagName !== 'IMG') el.style.width = '';
  }
  // the .docx converter understands tags better than CSS for character styles
  for (const el of clone.querySelectorAll('span[style],font[style]')) {
    const st = el.style;
    const deco = `${st.textDecoration} ${st.textDecorationLine}`;
    const wraps = [];
    if (/^(bold|[6-9]00)$/.test(st.fontWeight)) { wraps.push('b'); st.fontWeight = ''; }
    if (st.fontStyle === 'italic') { wraps.push('i'); st.fontStyle = ''; }
    if (/underline/.test(deco)) wraps.push('u');
    if (/line-through/.test(deco)) wraps.push('s');
    if (/underline|line-through/.test(deco)) { st.textDecoration = ''; st.textDecorationLine = ''; }
    if (st.verticalAlign === 'sub' || st.verticalAlign === 'super') { wraps.push(st.verticalAlign === 'sub' ? 'sub' : 'sup'); st.verticalAlign = ''; }
    if (wraps.length) el.innerHTML = wraps.map(t => `<${t}>`).join('') + el.innerHTML + wraps.reverse().map(t => `</${t}>`).join('');
  }
  for (const el of clone.querySelectorAll('s,strike,del')) {
    const s = document.createElement('span');
    s.style.fontFamily = STRIKE_FONT;
    s.innerHTML = el.innerHTML;
    el.replaceWith(s);
  }
  for (const el of clone.querySelectorAll('.page-break')) {
    el.removeAttribute('contenteditable');
    el.setAttribute('style', 'page-break-after: always;');
  }
  // The converter drops a tag's formatting when it directly wraps another tag
  // (<b><i>x</i></b> loses the bold), but keeps it through a span: <b><span><i>…
  for (const el of clone.querySelectorAll('b,strong,i,em,u,sub,sup,a')) {
    const s = document.createElement('span');
    while (el.firstChild) s.appendChild(el.firstChild);
    el.appendChild(s);
  }
  return clone.innerHTML;
}
function buildHtmlDoc(raw) {
  const m = state.page.margins;
  const css = `@page{size:${state.pw}in ${state.ph}in;margin:${m.top}in ${m.right}in ${m.bottom}in ${m.left}in}
body{font-family:Calibri,'Segoe UI',sans-serif;font-size:11pt;line-height:1.15;max-width:${(state.pw - m.left - m.right)}in;margin:0.6in auto;color:#000}
p{margin:0 0 8pt}h1,h2,h3{font-family:'Calibri Light',Calibri,sans-serif;font-weight:normal;color:#2f5496}h1{font-size:16pt}h2{font-size:13pt}h3{font-size:12pt;color:#1f3763}
h1.doc-title{font-size:28pt;color:#000}blockquote{margin:10pt .6in;padding-left:10pt;font-style:italic;color:#404040;border-left:3px solid #c8c8c8}
pre{font-family:Consolas,monospace;font-size:10pt;background:#f3f3f3;padding:6pt 8pt;white-space:pre-wrap}
table{border-collapse:collapse;width:100%;margin:0 0 8pt}td,th{border:1px solid #8e8e8e;padding:3px 6px;vertical-align:top}th{background:#f2f2f2}
img{max-width:100%}.page-break{break-after:page}a{color:#0563c1}`;
  return `<!DOCTYPE html>\n<html><head><meta charset="utf-8"><title>${esc(baseName(state.name))}</title><meta name="generator" content="Nisar Office Studio"><style>${css}</style></head><body>${raw}</body></html>`;
}
function buildSaveData() {
  for (const h of hooks.beforeSave) h();
  const raw = cleanClone().innerHTML;
  // docxExtras: comments / tracked changes / footnotes that main.js writes as real Word parts
  const docxExtras = { comments: [], changes: [], footnotes: [] };
  const extra = {};
  for (const h of hooks.extra) Object.assign(extra, h());
  return {
    path: state.path,
    name: baseName(state.name),
    raw,
    exportHtml: buildExportHtml(docxExtras),
    docxExtras,
    extra,
    htmlDoc: buildHtmlDoc(raw),
    text: editor.innerText.replace(/​/g, ''),
    page: { ...state.page, w: state.pw, h: state.ph },
  };
}
let saving = false;
async function saveDoc(saveAs) {
  if (saving) return false;
  saving = true;
  hideBackstage();
  try {
    const r = await api.save(buildSaveData(), saveAs);
    if (!r || r.canceled) return false;
    if (r.error) { await alertBox("Couldn't save", r.error); return false; }
    state.path = r.path;
    state.name = r.name;
    state.dirty = false;
    state.savedAt = new Date();
    updateTitle();
    api.recoveryClear();
    $('#st-saved').textContent = `Saved ${state.savedAt.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}`;
    if (r.format === 'txt') toast('Saved as plain text. Formatting and pictures are not kept in .txt files.');
    else toast(`Saved to ${r.path}`, () => api.showInFolder(r.path), 'Show in folder');
    return true;
  } catch (err) {
    await alertBox("Couldn't save", String(err && err.message || err));
    return false;
  } finally {
    saving = false;
  }
}
async function openDoc() {
  const p = await api.openDialog();
  if (p) await openPath(p);
}
async function openPath(p) {
  hideBackstage();
  if (/\.pdf$/i.test(p)) {
    // PDFs open in the PDF editor: here if this window is empty, otherwise in a new window
    if (isPristine()) api.switchEditor('pdf', p); else api.newWindow({ file: p });
    return;
  }
  if (!isPristine()) { api.newWindow({ file: p }); return; }
  const doc = await api.openPath(p);
  if (doc.error) { await alertBox("Couldn't open", doc.error); refreshRecent(); return; }
  loadDoc(doc);
}
async function exportPdf() {
  hideBackstage();
  selectImage(null);
  if (window.CSS && CSS.highlights) CSS.highlights.clear();
  const dir = state.path ? state.path.replace(/[\\/][^\\/]*$/, '') : null;
  const r = await api.exportPdf({ name: baseName(state.name), dir, header: state.page.header, footer: state.page.footer, pageNumbers: state.page.pageNumbers });
  paintFind();
  if (!r || r.canceled) return;
  if (r.error) { await alertBox("Couldn't export PDF", r.error); return; }
  toast('PDF saved.', () => api.openLastPdf(), 'Open PDF');
}
async function printPreview() {
  hideBackstage();
  selectImage(null);
  if (window.CSS && CSS.highlights) CSS.highlights.clear();
  const p = state.page;
  const r = await api.printPreview({ name: baseName(state.name), header: p.header, footer: p.footer, pageNumbers: p.pageNumbers });
  paintFind();
  if (r && r.error) await alertBox('Print Preview', r.error);
}
async function printDoc() {
  hideBackstage();
  selectImage(null);
  if (window.CSS && CSS.highlights) CSS.highlights.clear();
  const r = await api.print();
  paintFind();
  if (r && !r.success && r.reason && r.reason !== 'cancelled' && r.reason !== 'Print job canceled') toast(`Printing failed: ${r.reason}`);
}
async function confirmDiscard() {
  if (!state.dirty) return true;
  const c = await api.confirmUnsaved(state.name);
  if (c === 0) return saveDoc(false);
  return c === 1;
}
async function pastePlain() {
  const t = await api.clipboardText();
  if (t) { restoreSelection(); document.execCommand('insertText', false, t); }
}

// ---------------------------------------------------------------- templates
function today() { return new Date().toLocaleDateString(undefined, { dateStyle: 'long' }); }
const TEMPLATES = {
  blank: { name: 'Blank document', thumb: '', html: () => '<p><br></p>' },
  letter: {
    name: 'Letter',
    thumb: '<div class="ln" style="width:40%"></div><div class="ln" style="width:35%"></div><div class="ln" style="width:30%;margin-bottom:12px"></div><div class="ln" style="width:28%;margin-bottom:12px"></div><div class="ln"></div><div class="ln"></div><div class="ln" style="width:70%"></div>',
    html: () => `<p>Your Name<br>Street Address<br>City, State ZIP<br>you@example.com</p><p>${esc(today())}</p><p>Recipient Name<br>Title<br>Company Name<br>Street Address<br>City, State ZIP</p><p>Dear Recipient Name,</p><p>Start your letter here. Say why you are writing in the first sentence or two.</p><p>Use the next paragraph to give the details. Keep each paragraph focused on one idea.</p><p>Close with what you would like to happen next, and thank the reader for their time.</p><p>Sincerely,</p><p><br></p><p>Your Name</p>`,
  },
  resume: {
    name: 'Resume',
    thumb: '<div class="t1">Your Name</div><div class="ln" style="width:60%"></div><div class="t2">EXPERIENCE</div><div class="ln"></div><div class="ln" style="width:80%"></div><div class="t2">EDUCATION</div><div class="ln" style="width:70%"></div><div class="t2">SKILLS</div><div class="ln" style="width:50%"></div>',
    html: () => '<h1 class="doc-title">Your Name</h1><p>City, State · (555) 555-0100 · you@example.com · linkedin.com/in/yourname</p><h1>Summary</h1><p>Two or three sentences about who you are, what you do best and the kind of role you are looking for.</p><h1>Experience</h1><h2>Job Title · Company Name</h2><p><i>Month Year – Present</i></p><ul><li>Describe an achievement, with a number if you can (for example, "cut costs by 15%").</li><li>Describe another responsibility or result.</li></ul><h2>Job Title · Company Name</h2><p><i>Month Year – Month Year</i></p><ul><li>Describe an achievement.</li><li>Describe another responsibility or result.</li></ul><h1>Education</h1><h2>Degree · School Name</h2><p><i>Year</i></p><h1>Skills</h1><ul><li>Skill one</li><li>Skill two</li><li>Skill three</li></ul>',
  },
  report: {
    name: 'Report',
    thumb: '<div class="t1">Report Title</div><div class="ln" style="width:50%;margin-bottom:10px"></div><div class="t2">Introduction</div><div class="ln"></div><div class="ln"></div><div class="t2">Findings</div><div class="ln"></div><div style="height:22px;border:1px solid #ccc;margin:5px 0"></div><div class="ln" style="width:80%"></div>',
    html: () => `<h1 class="doc-title">Report Title</h1><p class="doc-subtitle">Subtitle · ${esc(today())}</p><h1>Introduction</h1><p>Explain the purpose of this report and what the reader will learn.</p><h1>Findings</h1><p>Summarize what you found. Use the table below to compare results.</p><table><tbody><tr><th>Item</th><th>Value</th><th>Notes</th></tr><tr><td>First</td><td>100</td><td><br></td></tr><tr><td>Second</td><td>250</td><td><br></td></tr></tbody></table><h2>Details</h2><p>Add supporting detail here.</p><h1>Conclusion</h1><p>Wrap up with your main recommendation and next steps.</p>`,
  },
  notes: {
    name: 'Meeting notes',
    thumb: '<div class="t1">Meeting Notes</div><div class="ln" style="width:45%"></div><div class="t2">Agenda</div><div class="ln" style="width:70%"></div><div class="ln" style="width:60%"></div><div class="t2">Action items</div><div style="height:30px;border:1px solid #ccc;margin:5px 0"></div>',
    html: () => `<h1 class="doc-title">Meeting Notes</h1><p><b>Date:</b> ${esc(today())}<br><b>Attendees:</b> <br><b>Location:</b> </p><h1>Agenda</h1><ol><li>Topic one</li><li>Topic two</li><li>Topic three</li></ol><h1>Notes</h1><p><br></p><h1>Action items</h1><table><tbody><tr><th>Task</th><th>Owner</th><th>Due</th></tr><tr><td><br></td><td><br></td><td><br></td></tr><tr><td><br></td><td><br></td><td><br></td></tr></tbody></table>`,
  },
};
function templateCards() {
  const pdfCard = '<button class="tpl" data-cmd="openPdf"><div class="thumb pdf-thumb"><span>PDF</span></div>Edit a PDF</button>';
  return Object.entries(TEMPLATES).map(([k, t]) => `<button class="tpl" data-tpl="${k}"><div class="thumb">${t.thumb}</div>${esc(t.name)}</button>`).join('') + pdfCard;
}
async function openPdf() {
  const p = await api.openDialog('pdf');
  if (p) await openPath(p);
}
function newFromTemplate(key) {
  hideBackstage();
  if (!TEMPLATES[key]) return;
  if (isPristine()) {
    if (key !== 'blank') {
      editor.innerHTML = TEMPLATES[key].html();
      state.dirty = false;
      applyPage();
      placeCaret(editor.firstElementChild, true);
      resetHistory();
      updateTitle();
      scheduleStats();
    }
    return;
  }
  api.newWindow(key === 'blank' ? {} : { template: key });
}

// ---------------------------------------------------------------- backstage
function showBackstage(pane = 'home') {
  closeMenu();
  const h = new Date().getHours();
  $('#bs-greeting').textContent = h < 12 ? 'Good morning' : h < 18 ? 'Good afternoon' : 'Good evening';
  $('#backstage').hidden = false;
  switchPane(pane);
  refreshRecent();
}
function switchPane(pane) {
  for (const b of $$('.bs-item[data-bs]')) b.classList.toggle('active', b.dataset.bs === pane);
  for (const s of $$('.bs-pane')) s.classList.toggle('active', s.dataset.pane === pane);
}
function hideBackstage() {
  if ($('#backstage').hidden) return;
  $('#backstage').hidden = true;
  restoreSelection();
}
async function refreshRecent() {
  const list = await api.getRecent();
  const ul = $('#recent');
  if (!list.length) { ul.innerHTML = '<li class="empty">Documents you open or save will appear here.</li>'; return; }
  ul.innerHTML = list.map((p, i) => {
    const name = p.split(/[\\/]/).pop();
    return `<li data-i="${i}">${svg('doc')}<div class="meta"><div class="nm">${esc(name)}</div><div class="pt">${esc(p)}</div></div><button class="x" data-rm="${i}" title="Remove from list">${svg('close')}</button></li>`;
  }).join('');
  ul.onclick = e => {
    const rm = e.target.closest('[data-rm]');
    if (rm) { api.removeRecent(list[+rm.dataset.rm]); refreshRecent(); return; }
    const li = e.target.closest('li[data-i]');
    if (li) openPath(list[+li.dataset.i]);
  };
}
const SHORTCUTS = [
  ['New window', 'Ctrl+N'], ['Open', 'Ctrl+O'], ['Save', 'Ctrl+S'], ['Save As', 'F12 or Ctrl+Shift+S'], ['Print', 'Ctrl+P'], ['Close', 'Ctrl+W'],
  ['Undo / Redo', 'Ctrl+Z / Ctrl+Y'], ['Cut / Copy / Paste', 'Ctrl+X / C / V'], ['Paste as plain text', 'Ctrl+Shift+V'], ['Select all', 'Ctrl+A'],
  ['Find / Replace', 'Ctrl+F / Ctrl+H'], ['Bold / Italic / Underline', 'Ctrl+B / I / U'], ['Subscript / Superscript', 'Ctrl+= / Ctrl+Shift+='],
  ['Bigger / smaller text', 'Ctrl+] / Ctrl+['], ['Clear character formatting', 'Ctrl+Space'], ['Align left / center / right / justify', 'Ctrl+L / E / R / J'],
  ['Line spacing 1 / 1.5 / 2', 'Ctrl+1 / 5 / 2'], ['Heading 1 / 2 / 3', 'Ctrl+Alt+1 / 2 / 3'], ['Normal style', 'Ctrl+Shift+N'], ['Bullets', 'Ctrl+Shift+L'],
  ['Indent / outdent', 'Ctrl+M / Ctrl+Shift+M'], ['Insert link', 'Ctrl+K'], ['Page break', 'Ctrl+Enter'], ['Zoom', 'Ctrl+mouse wheel'], ['Collapse ribbon', 'Ctrl+F1'],
  ['Exit focus mode, close panels', 'Esc'],
];

// ---------------------------------------------------------------- events
document.addEventListener('mousedown', e => {
  // keep the caret in the document when clicking toolbar buttons
  if (e.target.closest('#ribbon button, #statusbar button, .menu button, #imgbar button, #findbar button')) e.preventDefault();
  if (openMenuEl && !e.target.closest('.menu') && !e.target.closest('[data-menu]')) closeMenu();
});
document.addEventListener('click', e => {
  const m = e.target.closest('[data-menu]');
  if (m) { openMenu(m.dataset.menu, m); return; }
  const c = e.target.closest('[data-cmd]');
  if (c) { if (c.dataset.cmd !== 'backstage') closeMenu(); runCommand(c.dataset.cmd, c); return; }
  const tab = e.target.closest('[data-tab]');
  if (tab) {
    for (const t of $$('.tab[data-tab]')) t.classList.toggle('active', t === tab);
    for (const p of $$('.panel')) p.classList.toggle('active', p.dataset.panel === tab.dataset.tab);
    if (document.body.classList.contains('ribbon-collapsed')) document.body.classList.remove('ribbon-collapsed');
    return;
  }
  const bs = e.target.closest('[data-bs]');
  if (bs) { switchPane(bs.dataset.bs); return; }
  const tpl = e.target.closest('[data-tpl]');
  if (tpl) { newFromTemplate(tpl.dataset.tpl); return; }
  const f = e.target.closest('[data-find]');
  if (f) { ({ prev: () => findStep(-1), next: () => findStep(1), close: closeFind, replace: () => withHistory(replaceOne), all: () => withHistory(replaceAll) })[f.dataset.find](); return; }
  const im = e.target.closest('[data-img]');
  if (im) withHistory(() => imageAction(im.dataset.img));
});

document.addEventListener('selectionchange', () => {
  const sel = getSelection();
  if (sel.rangeCount && inEditor(sel.anchorNode)) {
    state.savedRange = sel.getRangeAt(0).cloneRange();
    scheduleUi();
    scheduleStats();
  }
});

editor.addEventListener('input', e => {
  if (/^(delete|insertParagraph|insertFromPaste|insertFromDrop|formatOutdent|formatIndent)/.test(e.inputType || '')) {
    const block = closestBlock(getSelection().anchorNode);
    cleanRedundantSpans(e.inputType.startsWith('insertFrom') ? editor : (block && block.parentElement) || editor);
  }
  // keep text inside paragraphs
  for (const n of [...editor.childNodes]) {
    if (n.nodeType === 3 && n.data.trim()) { document.execCommand('formatBlock', false, 'p'); break; }
  }
  if (!editor.firstChild || editor.innerHTML === '<br>') editor.innerHTML = '<p><br></p>';
  const typing = /^(insertText|insertCompositionText|insertReplacementText|deleteContent|deleteWord)/.test(e.inputType || '');
  if (typing) dropCaretMarker();
  markDirty(typing);
  if (!$('#findbar').hidden && findInput.value) { state.find.index = -1; const keep = state.savedRange; refreshFind(false); state.savedRange = keep; }
});

editor.addEventListener('beforeinput', e => {
  if (e.inputType === 'historyUndo') { e.preventDefault(); undo(); }
  else if (e.inputType === 'historyRedo') { e.preventDefault(); redo(); }
});

editor.addEventListener('click', e => {
  const img = e.target.closest('img');
  selectImage(img && inEditor(img) ? img : null);
  const a = e.target.closest('a[href]');
  if (a && (e.ctrlKey || e.metaKey)) {
    e.preventDefault();
    const href = a.getAttribute('href') || '';
    if (href.startsWith('#')) for (const h of hooks.internalLink) h(href); else api.openExternal(a.href);
  }
});
editor.addEventListener('mouseup', () => {
  if (!state.painter) return;
  setTimeout(() => { const s = getSelection(); if (s.rangeCount && !s.isCollapsed && inEditor(s.anchorNode)) applyPainter(); }, 0);
});
editor.addEventListener('mouseover', e => {
  const a = e.target.closest('a[href]');
  if (a && !a.title) a.title = `${a.getAttribute('href')}\nCtrl+Click to open`;
});

editor.addEventListener('paste', e => {
  const cd = e.clipboardData;
  if (!cd) return;
  const html = cd.getData('text/html');
  const files = [...cd.files].filter(f => f.type.startsWith('image/'));
  if (files.length && !html) {
    e.preventDefault();
    for (const f of files) { const fr = new FileReader(); fr.onload = () => insertImage(fr.result); fr.readAsDataURL(f); }
    return;
  }
  if (html) {
    e.preventDefault();
    let clean = sanitize(html);
    if (!clean.trim()) clean = esc(cd.getData('text/plain')).replace(/\r?\n/g, '<br>');
    for (const h of hooks.paste) clean = h(clean); // e.g. mark pasted text as a tracked insertion
    if (clean.trim()) document.execCommand('insertHTML', false, clean);
  }
});

// drag & drop: documents open, pictures are inserted
document.addEventListener('dragover', e => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) e.preventDefault(); });
document.addEventListener('drop', e => {
  const files = e.dataTransfer ? [...e.dataTransfer.files] : [];
  if (!files.length) return;
  e.preventDefault();
  const docs = files.filter(f => /\.(docx|html?|txt|doc|pdf)$/i.test(f.name));
  for (const f of docs) { const p = api.pathForFile(f); if (p) openPath(p); }
  const imgs = files.filter(f => f.type.startsWith('image/'));
  if (imgs.length) {
    const pos = document.caretRangeFromPoint(e.clientX, e.clientY);
    if (pos && inEditor(pos.startContainer)) { editor.focus(); setRange(pos); }
    for (const f of imgs) { const fr = new FileReader(); fr.onload = () => insertImage(fr.result); fr.readAsDataURL(f); }
  }
});

document.addEventListener('keydown', e => {
  const k = e.key;
  const ctrl = e.ctrlKey || e.metaKey;
  if (k === 'Escape') {
    if (openMenuEl) { closeMenu(); return; }
    if ($('#modal').open) return;
    if (!$('#backstage').hidden) { hideBackstage(); return; }
    if (!$('#findbar').hidden) { closeFind(); return; }
    if (state.painter) { toggleFormatPainter(); return; }
    if (document.body.classList.contains('focus-mode')) { setFocusMode(false); return; }
    if (state.img) { selectImage(null); return; }
    return;
  }
  if ($('#modal').open) return;
  if (k === 'F12') { e.preventDefault(); saveDoc(true); return; }
  if (ctrl && k === 'F1') { e.preventDefault(); runCommand('collapseRibbon'); return; }

  if (e.target === findInput && k === 'Enter') { e.preventDefault(); findStep(e.shiftKey ? -1 : 1); return; }
  if (e.target === replaceInput && k === 'Enter') { e.preventDefault(); withHistory(replaceOne); return; }
  if (e.target.id === 'font-size' && k === 'Enter') { e.preventDefault(); withHistory(() => applyFontSize(e.target.value)); editor.focus(); return; }

  if (!ctrl) {
    if (k === 'Tab' && e.target === editor) { e.preventDefault(); withHistory(() => handleTab(e.shiftKey)); }
    if (k === 'Backspace' && e.target === editor && !e.shiftKey && !e.altKey) {
      // Backspace at the very start of a list item removes the bullet (or outdents), like Word
      const li = closestEl(getSelection().anchorNode, 'li');
      if (li && caretAtStartOf(li)) { e.preventDefault(); withHistory(() => listIndent(-1)); }
    }
    return;
  }
  const key = k.toLowerCase();
  const S = e.shiftKey, A = e.altKey;
  const inDoc = e.target === editor;
  const map = [
    [key === 'b' && !S && !A && inDoc, () => exec('bold')],
    [key === 'i' && !S && !A && inDoc, () => exec('italic')],
    [key === 'u' && !S && !A && inDoc, () => exec('underline')],
    [key === 'z' && !S && inDoc, undo],
    [(key === 'y' || (S && key === 'z')) && inDoc, redo],
    [!S && key === 's', () => saveDoc(false)],
    [S && key === 's', () => saveDoc(true)],
    [key === 'o', openDoc],
    [key === 'n' && !S, () => api.newWindow()],
    [key === 'p', printDoc],
    [key === 'w', () => window.close()],
    [key === 'f', () => openFind(false)],
    [key === 'h', () => openFind(true)],
    [key === 'k', insertLink],
    [key === 'e' && inDoc, () => exec('justifyCenter')],
    [key === 'l' && !S && inDoc, () => exec('justifyLeft')],
    [key === 'r' && inDoc, () => exec('justifyRight')],
    [key === 'j' && inDoc, () => exec('justifyFull')],
    [S && key === 'l' && inDoc, () => toggleList('ul')],
    [S && key === 'n' && inDoc, () => applyStyle('normal')],
    [S && key === 'v' && inDoc, pastePlain],
    [(k === '=' || k === '+') && S && inDoc, () => exec('superscript')],
    [k === '=' && !S && inDoc, () => exec('subscript')],
    [(k === ']' || (S && k === '>')) && inDoc, () => stepFont(1)],
    [(k === '[' || (S && k === '<')) && inDoc, () => stepFont(-1)],
    [k === ' ' && inDoc, () => exec('removeFormat')],
    [key === 'm' && inDoc, () => indent(S ? -1 : 1)],
    [k === 'Enter' && inDoc, insertPageBreak],
    [A && k === '1' && inDoc, () => applyStyle('h1')],
    [A && k === '2' && inDoc, () => applyStyle('h2')],
    [A && k === '3' && inDoc, () => applyStyle('h3')],
    [!A && k === '1' && inDoc, () => setLineSpacing('1')],
    [!A && k === '2' && inDoc, () => setLineSpacing('2.0')],
    [!A && k === '5' && inDoc, () => setLineSpacing('1.5')],
  ];
  for (const [when, fn] of map) {
    if (when) { e.preventDefault(); const r = fn === undo || fn === redo ? fn() : withHistory(fn); if (r && r.catch) r.catch(showError); return; }
  }
});

function handleTab(shift) {
  const sel = getSelection();
  const cell = closestEl(sel.anchorNode, 'td,th');
  if (cell) {
    const table = cell.closest('table');
    let cells = $$('td,th', table).filter(c => c.closest('table') === table);
    let i = cells.indexOf(cell) + (shift ? -1 : 1);
    if (i >= cells.length) { tableOp('rowBelow'); cells = $$('td,th', table).filter(c => c.closest('table') === table); }
    if (i < 0) return;
    const r = document.createRange();
    r.selectNodeContents(cells[i]);
    setRange(r);
    return;
  }
  if (closestEl(sel.anchorNode, 'li')) { listIndent(shift ? -1 : 1); return; }
  if (!shift) document.execCommand('insertText', false, '  ');
}

$('#modal').addEventListener('close', () => setTimeout(() => {
  if ($('#modal').open || !$('#backstage').hidden) return;
  restoreSelection();
  updateUi();
}, 0));
$('#font-family').addEventListener('change', e => withHistory(() => applyFontName(e.target.value)));
$('#font-family').addEventListener('mousedown', () => loadSystemFonts(true));
$('#font-size').addEventListener('change', e => withHistory(() => applyFontSize(e.target.value)));
$('#font-size').addEventListener('focus', e => e.target.select());
$('#zoom-range').addEventListener('input', e => setZoom(e.target.value / 100));
findInput.addEventListener('input', () => { state.find.index = -1; refreshFind(false); });
$('#find-case').addEventListener('change', () => refreshFind(false));
$('#find-word').addEventListener('change', () => refreshFind(false));

scroller.addEventListener('wheel', e => {
  if (!e.ctrlKey) return;
  e.preventDefault();
  setZoom(state.zoom + (e.deltaY < 0 ? 0.1 : -0.1));
}, { passive: false });
scroller.addEventListener('scroll', () => positionImgBar(), { passive: true });
// on narrow windows the ribbon scrolls sideways with the mouse wheel
$('#panels').addEventListener('wheel', e => {
  const p = e.currentTarget;
  if (p.scrollWidth <= p.clientWidth || Math.abs(e.deltaX) > Math.abs(e.deltaY)) return;
  e.preventDefault();
  p.scrollLeft += e.deltaY;
}, { passive: false });
scroller.addEventListener('mousedown', e => {
  // clicking the grey area or the page margin puts the caret in the document
  if (e.target === scroller || e.target === zoomer || e.target === page) {
    e.preventDefault();
    editor.focus({ preventScroll: true });
    const last = editor.lastElementChild;
    if (last && e.target !== page) placeCaret(last, false);
    else if (state.savedRange) setRange(state.savedRange);
  }
});
new ResizeObserver(() => scheduleLayout()).observe(editor);
window.addEventListener('resize', () => { closeMenu(); positionImgBar(); });

api.on('request-close', async () => {
  stopSpeak();
  if (await confirmDiscard()) { api.recoveryClear(); api.closeNow(); }
});
api.on('cmd', name => {
  if (name === 'pastePlain') pastePlain();
  else runCommand(name);
});

// autosave a recovery copy every 30 seconds while there are unsaved changes
setInterval(() => {
  if (!state.dirty) return;
  const extra = {};
  for (const h of hooks.extra) Object.assign(extra, h());
  api.recoverySave({ html: cleanClone().innerHTML, page: state.page, extra, name: state.name, path: state.path, time: Date.now() });
}, 30000);

// ---------------------------------------------------------------- fonts
function addFontOption(name, prepend = false) {
  const ff = $('#font-family');
  const o = document.createElement('option');
  o.value = name;
  o.textContent = name;
  o.style.fontFamily = `"${name}"`;
  if (prepend) ff.prepend(o); else ff.appendChild(o);
}
let systemFontsLoaded = false;
async function loadSystemFonts(fromGesture) {
  if (systemFontsLoaded || !window.queryLocalFonts) return;
  try {
    const fonts = await window.queryLocalFonts();
    const fams = [...new Set(fonts.map(f => f.family))].filter(f => f && !f.startsWith('@')).sort((a, b) => a.localeCompare(b));
    if (fams.length < 5) return;
    systemFontsLoaded = true;
    const ff = $('#font-family');
    const cur = ff.value;
    ff.innerHTML = '';
    for (const f of fams) addFontOption(f);
    if (![...ff.options].some(o => o.value === cur)) addFontOption(cur, true);
    ff.value = cur;
  } catch { /* fall back to the built-in list */ }
}

// ---------------------------------------------------------------- start up
(async function init() {
  for (const el of $$('[data-icon]')) el.insertAdjacentHTML('afterbegin', svg(el.dataset.icon));
  for (const f of FALLBACK_FONTS) addFontOption(f);
  $('#font-family').value = 'Calibri';
  $('#tpl-home').innerHTML = templateCards();
  $('#tpl-all').innerHTML = templateCards();
  $('#shortcuts').innerHTML = SHORTCUTS.map(([a, b]) => `<div><span>${esc(a)}</span><kbd>${esc(b)}</kbd></div>`).join('');

  document.execCommand('defaultParagraphSeparator', false, 'p');
  document.execCommand('styleWithCSS', false, true);
  editor.innerHTML = '<p><br></p>';
  applyPage();
  setZoom(1);

  let init = {};
  try { init = await api.init(); } catch (err) { console.error(err); }
  state.name = `Document${init.untitled || 1}`;
  state.userName = init.userName || 'Author';
  state.initials = init.initials || '';
  if (init.dark) document.body.classList.add('dark');

  if (init.doc) {
    if (init.doc.error) { updateTitle(); await alertBox("Couldn't open", init.doc.error); } else loadDoc(init.doc);
  } else if (init.template && TEMPLATES[init.template]) {
    editor.innerHTML = TEMPLATES[init.template].html();
    applyPage();
  }
  updateTitle();
  updateToggles();
  editor.focus();
  if (!init.doc) { placeCaret(editor.firstElementChild || editor, true); resetHistory(); }
  scheduleStats();
  updateUi();
  loadSystemFonts(false);

  if (init.recovery) {
    const when = new Date(init.recovery.time || Date.now()).toLocaleString();
    const ok = await confirmBox('Recover unsaved document?', `Nisar Office Studio closed before "${init.recovery.name || 'a document'}" was saved. A copy from ${when} was kept.`, 'Recover', 'Discard');
    if (ok) {
      if (isPristine()) {
        loadDoc({ html: init.recovery.html, page: init.recovery.page, extra: init.recovery.extra, path: init.recovery.path, name: init.recovery.name });
        state.dirty = true;
        updateTitle();
        toast('Recovered. Save it to keep it.');
      } else {
        // this window already holds a document; keep both by placing the recovered one in a new window
        toast('Recovered text was not placed: this window already has a document open.');
      }
    }
  }
})();
