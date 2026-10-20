'use strict';
/* Nisar Office Studio: document editor features from Word's References and Review tabs —
 * comments, Track Changes, table of contents, footnotes, navigation pane, Font and
 * Paragraph dialogs, watermark, links to headings.
 *
 * Loaded after app.js and uses its globals (state, editor, hooks, COMMANDS, MENUS…).
 * Comments, tracked changes and footnotes are exported to .docx as real Word parts:
 * the export marks them with private-use characters that main.js turns into
 * <w:comment>, <w:ins>/<w:del> and <w:footnote> elements.
 */

const uidShort = () => Math.random().toString(36).slice(2, 9);
const MARK = {
  cmtStart: '', cmtStartEnd: '', cmtEnd: '', cmtEndEnd: '',
  insStart: '', insStartEnd: '', insEnd: '', insEndEnd: '',
  delStart: '', delStartEnd: '', delEnd: '', delEndEnd: '',
  fn: '', fnEnd: '',
};
for (const c of ['toc', 'toc-entry', 'toc-l1', 'toc-l2', 'toc-l3', 'toc-text', 'toc-dots', 'toc-page', 'toc-title', 'toc-empty', 'fn-ref', 'footnotes', 'fn-sep', 'cmt-anchor', 'tc']) KEEP_CLASSES.add(c);

Object.assign(ICONS, {
  comment: '<path d="M4 4h16v12H9l-5 4z"/><path d="M8 9h8M8 12h5"/>',
  track: '<path d="M4 20h4L19 9l-4-4L4 16z"/><path d="M13 7l4 4"/><path d="M14 20h6" stroke-dasharray="2 2"/>',
  markup: '<path d="M4 6h10M4 10h16M4 14h8"/><path d="M14 18l2 2 4-5"/>',
  accept: '<path d="M5 12l5 5L20 7"/>',
  reject: '<path d="M6 6l12 12M18 6L6 18"/>',
  toc: '<path d="M4 5h11M4 9h14M7 13h11M7 17h9M4 21h12"/><path d="M19 5h1M20 9h0M20 13h0"/>',
  refresh: '<path d="M20 11a8 8 0 10-2.3 5.7"/><path d="M20 4v7h-7"/>',
  footnote: '<path d="M4 6h10M4 10h10M4 14h7"/><path d="M17 4v6M15 5l2-1"/><path d="M4 19h16" stroke-dasharray="2 2"/>',
  navpane: '<rect x="3" y="4" width="18" height="16" rx="1.5"/><path d="M9 4v16M5 8h2M5 11h2M5 14h2"/>',
  watermark: '<rect x="5" y="3" width="14" height="18" rx="1"/><path d="M8 16l8-8" opacity=".5"/><path d="M8 12l4-4" opacity=".5"/>',
  gear: '<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M4.9 19.1L7 17M17 7l2.1-2.1"/>',
});

// ---------------------------------------------------------------- shared helpers
function authorName() { return state.userName || 'Author'; }
function authorInitials() {
  if (state.initials) return state.initials;
  return authorName().split(/\s+/).filter(Boolean).map(w => w[0]).join('').slice(0, 3).toUpperCase() || 'A';
}
// stable colour per author, as Word does
function authorColor(name) {
  const palette = ['#b0217f', '#1f5fbf', '#1e8e3e', '#c55a11', '#7030a0', '#008080', '#c00000'];
  let h = 0;
  for (const ch of String(name)) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
  return palette[h % palette.length];
}
// wrap every piece of text inside `range` with an element made by `make()`
function wrapRangeText(range, make) {
  const segs = textSegments(range);
  const made = [];
  for (const seg of segs) {
    let node = seg.node;
    if (seg.e < node.length) node.splitText(seg.e);
    if (seg.s > 0) node = node.splitText(seg.s);
    const el = make();
    node.parentNode.insertBefore(el, node);
    el.appendChild(node);
    made.push(el);
  }
  return made;
}
function headingBlocks() {
  return $$('h1, h2, h3', editor).filter(h => !h.closest('.toc, .footnotes') && h.textContent.trim());
}
function headingId(h) {
  if (!h.dataset.hid) h.dataset.hid = uidShort();
  return h.dataset.hid;
}
function headingLevel(h) { return h.classList.contains('doc-title') ? 0 : +h.tagName[1]; }
// page on which an element sits (1-based), using the editor's pagination
function pageOfElement(el) {
  const er = editor.getBoundingClientRect();
  const r = el.getBoundingClientRect();
  const scale = er.height / (editor.offsetHeight || 1) || 1;
  const g = geometry();
  const y = (r.top - er.top) / scale;
  return clamp(Math.floor(y / (g.paged ? g.S : g.C)) + 1, 1, state.pageCount || 1);
}
function scrollToElement(el, select = true) {
  const r = el.getBoundingClientRect();
  const sr = scroller.getBoundingClientRect();
  scroller.scrollTop += r.top - sr.top - 80;
  if (select) { editor.focus({ preventScroll: true }); placeCaret(el, true); }
}

// ---------------------------------------------------------------- navigation pane
function toggleNavPane(force) {
  const on = force ?? $('#navpane').hidden;
  $('#navpane').hidden = !on;
  document.body.classList.toggle('show-nav', on);
  for (const b of $$('[data-cmd="navPane"]')) b.classList.toggle('on', on);
  if (on) renderNav();
  scheduleLayout();
}
let navTimer = null;
function scheduleNav() {
  if ($('#navpane').hidden) return;
  clearTimeout(navTimer);
  navTimer = setTimeout(renderNav, 400);
}
function renderNav() {
  const list = $('#nav-list');
  const q = $('#nav-search').value.trim();
  if (q) {
    // search results with a little context, like Word's Results tab
    const hits = [];
    const needle = q.toLocaleLowerCase();
    const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = walker.nextNode()) && hits.length < 200) {
      const hay = n.data.toLocaleLowerCase();
      let i = 0;
      while ((i = hay.indexOf(needle, i)) !== -1 && hits.length < 200) { hits.push({ node: n, i }); i += needle.length; }
    }
    list.innerHTML = hits.length
      ? `<div class="nav-empty">${hits.length}${hits.length === 200 ? '+' : ''} result${hits.length === 1 ? '' : 's'}</div>` + hits.map((h, k) => {
        const t = h.node.data;
        const a = Math.max(0, h.i - 30), b = Math.min(t.length, h.i + q.length + 40);
        return `<button class="nav-hit" data-hit="${k}">${a ? '…' : ''}${esc(t.slice(a, h.i))}<mark>${esc(t.slice(h.i, h.i + q.length))}</mark>${esc(t.slice(h.i + q.length, b))}${b < t.length ? '…' : ''}</button>`;
      }).join('')
      : '<div class="nav-empty">No results.</div>';
    list.onclick = ev => {
      const b = ev.target.closest('[data-hit]');
      if (!b) return;
      const h = hits[+b.dataset.hit];
      if (!h || !editor.contains(h.node)) return;
      const r = document.createRange();
      r.setStart(h.node, h.i);
      r.setEnd(h.node, h.i + q.length);
      editor.focus({ preventScroll: true });
      setRange(r);
      const rect = r.getBoundingClientRect(), sr = scroller.getBoundingClientRect();
      scroller.scrollTop += rect.top - sr.top - sr.height / 2;
    };
    return;
  }
  const hs = headingBlocks();
  list.innerHTML = hs.length
    ? hs.map((h, k) => { const lv = headingLevel(h); return `<button class="nav-item ${lv === 0 ? 'title' : `l${lv}`}" data-h="${k}" title="${esc(h.textContent.trim())}">${esc(h.textContent.trim())}</button>`; }).join('')
    : '<div class="nav-empty">Your headings appear here. Use the Heading 1, 2 and 3 styles on the Home tab to build an outline you can click to move around.</div>';
  list.onclick = ev => { const b = ev.target.closest('[data-h]'); if (b && hs[+b.dataset.h]) scrollToElement(hs[+b.dataset.h]); };
  markCurrentHeading();
}
function markCurrentHeading() {
  if ($('#navpane').hidden || $('#nav-search').value.trim()) return;
  const hs = headingBlocks();
  const sr = scroller.getBoundingClientRect();
  let cur = -1;
  hs.forEach((h, k) => { if (h.getBoundingClientRect().top < sr.top + sr.height * 0.35) cur = k; });
  $$('.nav-item').forEach((b, k) => b.classList.toggle('on', k === cur));
}
$('#nav-search').addEventListener('input', () => { clearTimeout(navTimer); navTimer = setTimeout(renderNav, 200); });
scroller.addEventListener('scroll', () => markCurrentHeading(), { passive: true });

// ---------------------------------------------------------------- table of contents
function tocHtml() {
  const hs = headingBlocks().filter(h => headingLevel(h) > 0);
  const entries = hs.map(h => {
    const lv = headingLevel(h);
    return `<p class="toc-entry toc-l${lv}" data-target="${headingId(h)}"><span class="toc-text">${esc(h.textContent.trim())}</span><span class="toc-dots"></span><span class="toc-page">${pageOfElement(h)}</span></p>`;
  }).join('');
  return `<p class="toc-title">Contents</p>${entries || '<p class="toc-empty">No headings yet. Apply Heading 1, 2 or 3 to your section titles, then click Update Table.</p>'}`;
}
function insertToc() {
  restoreSelection();
  if ($('.toc', editor)) { updateToc(); toast('The table of contents was updated.'); return; }
  const block = closestBlock(getSelection().anchorNode);
  const toc = document.createElement('div');
  toc.className = 'toc';
  toc.contentEditable = 'false';
  if (block && block.parentElement === editor) block.before(toc); else editor.prepend(toc);
  toc.innerHTML = tocHtml();
  layoutPages(); // headings may have moved down a page
  toc.innerHTML = tocHtml();
  markDirty();
}
function updateToc(silent = false) {
  const tocs = $$('.toc', editor);
  if (!tocs.length) { if (!silent) toast('There is no table of contents yet. Use References › Table of Contents.'); return; }
  for (const t of tocs) { t.contentEditable = 'false'; t.innerHTML = tocHtml(); }
  layoutPages();
  for (const t of tocs) t.innerHTML = tocHtml();
  if (!silent) markDirty();
}
function removeToc() { for (const t of $$('.toc', editor)) t.remove(); markDirty(); }
editor.addEventListener('click', ev => {
  const e = ev.target.closest('.toc-entry');
  if (e) {
    const h = $$('[data-hid]', editor).find(x => x.dataset.hid === e.dataset.target);
    if (h) scrollToElement(h);
  }
  const ref = ev.target.closest('sup.fn-ref');
  if (ref) { const li = $(`.footnotes li[data-fn="${ref.dataset.fn}"]`, editor); if (li) scrollToElement(li); }
});

// ---------------------------------------------------------------- footnotes
function footnotesBlock(create) {
  let fb = $('.footnotes', editor);
  if (!fb && create) {
    fb = document.createElement('div');
    fb.className = 'footnotes';
    fb.innerHTML = '<hr class="fn-sep" contenteditable="false"><ol></ol>';
    editor.appendChild(fb);
  }
  return fb;
}
function renumberFootnotes() {
  const refs = $$('sup.fn-ref', editor);
  const fb = footnotesBlock(false);
  if (!fb) { refs.forEach(r => r.remove()); return; }
  const ol = $('ol', fb) || fb.appendChild(document.createElement('ol'));
  const items = new Map($$('li[data-fn]', ol).map(li => [li.dataset.fn, li]));
  refs.forEach((r, i) => {
    r.textContent = String(i + 1);
    r.contentEditable = 'false';
    let li = items.get(r.dataset.fn);
    if (!li) { li = document.createElement('li'); li.dataset.fn = r.dataset.fn; li.innerHTML = '<br>'; }
    ol.appendChild(li); // keeps notes in the order of their references
    items.delete(r.dataset.fn);
  });
  for (const li of items.values()) li.remove(); // reference was deleted
  if (!refs.length) fb.remove();
}
function insertFootnote() {
  restoreSelection();
  const id = uidShort();
  const sup = document.createElement('sup');
  sup.className = 'fn-ref';
  sup.dataset.fn = id;
  sup.contentEditable = 'false';
  const sel = getSelection();
  const r = sel.getRangeAt(0);
  r.collapse(false);
  r.insertNode(sup);
  footnotesBlock(true);
  renumberFootnotes();
  const li = $(`.footnotes li[data-fn="${id}"]`, editor);
  markDirty();
  if (li) { scrollToElement(li, false); editor.focus({ preventScroll: true }); placeCaret(li, false); }
  toast('Type the footnote text. Footnotes become real Word footnotes at the bottom of the page in .docx files.');
}
function nextFootnote() {
  const refs = $$('sup.fn-ref', editor);
  if (!refs.length) { toast('There are no footnotes.'); return; }
  const sel = getSelection();
  const after = refs.find(rf => sel.rangeCount && sel.getRangeAt(0).comparePoint(rf, 0) > 0) || refs[0];
  scrollToElement(after, false);
  const r = document.createRange();
  r.setStartAfter(after);
  r.collapse(true);
  editor.focus({ preventScroll: true });
  setRange(r);
}
function showNotes() {
  const sel = getSelection();
  const li = closestEl(sel.anchorNode, '.footnotes li');
  if (li) { const ref = $(`sup.fn-ref[data-fn="${li.dataset.fn}"]`, editor); if (ref) { scrollToElement(ref, false); const r = document.createRange(); r.setStartAfter(ref); r.collapse(true); setRange(r); } return; }
  const fb = footnotesBlock(false);
  if (fb) scrollToElement(fb, false); else toast('There are no footnotes.');
}

// ---------------------------------------------------------------- comments
state.comments = [];
function commentById(id) { return state.comments.find(c => c.id === id); }
function anchorsOf(id) { return $$(`.cmt-anchor[data-cmt="${id}"]`, editor); }
function newComment() {
  restoreSelection();
  const sel = getSelection();
  if (!sel.rangeCount) return;
  if (sel.isCollapsed) { sel.modify('move', 'backward', 'word'); sel.modify('extend', 'forward', 'word'); }
  if (sel.isCollapsed || !sel.toString().trim()) { toast('Select the text you want to comment on.'); return; }
  const id = uidShort();
  const made = wrapRangeText(sel.getRangeAt(0), () => { const s = document.createElement('span'); s.className = 'cmt-anchor'; s.dataset.cmt = id; return s; });
  if (!made.length) return;
  state.comments.push({ id, author: authorName(), initials: authorInitials(), time: new Date().toISOString(), text: '', resolved: false });
  document.body.classList.remove('hide-comments');
  markDirty();
  renderComments();
  focusComment(id, true);
}
function deleteComment(id) {
  for (const a of anchorsOf(id)) unwrap(a);
  state.comments = state.comments.filter(c => c.id !== id);
  markDirty();
  renderComments();
}
function currentCommentId() {
  const active = $('.cmt-card.active');
  if (active) return active.dataset.cmt;
  const a = closestEl(getSelection().anchorNode, '.cmt-anchor');
  return a ? a.dataset.cmt : null;
}
function focusComment(id, edit = false) {
  for (const c of $$('.cmt-card')) c.classList.toggle('active', c.dataset.cmt === id);
  for (const a of $$('.cmt-anchor', editor)) a.classList.toggle('active', a.dataset.cmt === id);
  const card = $(`.cmt-card[data-cmt="${id}"]`);
  if (card && edit) { const b = $('.body', card); b.focus(); const r = document.createRange(); r.selectNodeContents(b); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); }
}
function stepComment(dir) {
  const ids = [...new Set($$('.cmt-anchor', editor).map(a => a.dataset.cmt))];
  if (!ids.length) { toast('There are no comments.'); return; }
  const cur = currentCommentId();
  const i = ids.indexOf(cur);
  const next = ids[(i + dir + ids.length) % ids.length] || ids[0];
  const a = anchorsOf(next)[0];
  if (a) scrollToElement(a, false);
  focusComment(next);
}
function renderComments() {
  const rail = $('#comment-rail');
  // keep only comments whose text is still in the document
  const live = state.comments.filter(c => anchorsOf(c.id).length);
  document.body.classList.toggle('has-comments', live.length > 0);
  for (const b of $$('[data-cmd="showComments"]')) b.classList.toggle('on', !document.body.classList.contains('hide-comments'));
  if (!live.length) { rail.innerHTML = ''; return; }
  const activeId = currentCommentId();
  const focused = document.activeElement && document.activeElement.closest && document.activeElement.closest('.cmt-card');
  if (focused) { positionComments(); return; } // don't rebuild under the user's caret
  rail.innerHTML = live.map(c => `<div class="cmt-card${c.resolved ? ' resolved' : ''}${c.id === activeId ? ' active' : ''}" data-cmt="${c.id}">
    <div class="who"><span class="av" style="background:${authorColor(c.author)}">${esc(c.initials || '?')}</span>${esc(c.author)}<span class="when">${esc(new Date(c.time).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' }))}</span></div>
    <div class="body" contenteditable="plaintext-only" spellcheck="true">${esc(c.text)}</div>
    <div class="acts"><button data-ca="resolve">${c.resolved ? 'Reopen' : 'Resolve'}</button><button data-ca="delete">Delete</button></div></div>`).join('');
  positionComments();
}
function positionComments() {
  const rail = $('#comment-rail');
  const pr = page.getBoundingClientRect();
  const scale = pr.height / (page.offsetHeight || 1) || 1;
  let lastBottom = 0;
  const cards = $$('.cmt-card', rail).map(card => {
    const a = anchorsOf(card.dataset.cmt)[0];
    return { card, top: a ? (a.getBoundingClientRect().top - pr.top) / scale : 0 };
  }).sort((x, y) => x.top - y.top);
  for (const { card, top } of cards) {
    const t = Math.max(top - 6, lastBottom + 8);
    card.style.top = `${t}px`;
    lastBottom = t + card.offsetHeight;
  }
}
$('#comment-rail').addEventListener('input', ev => {
  const card = ev.target.closest('.cmt-card');
  const c = card && commentById(card.dataset.cmt);
  if (!c) return;
  c.text = $('.body', card).innerText;
  if (!state.dirty) markDirty(true); else scheduleHistory();
  positionComments();
});
$('#comment-rail').addEventListener('focusout', ev => { if (ev.target.classList.contains('body')) { commitHistory(); setTimeout(renderComments, 0); } });
$('#comment-rail').addEventListener('click', ev => {
  const card = ev.target.closest('.cmt-card');
  if (!card) return;
  const id = card.dataset.cmt;
  const act = ev.target.closest('[data-ca]');
  if (act && act.dataset.ca === 'delete') { deleteComment(id); return; }
  if (act && act.dataset.ca === 'resolve') { const c = commentById(id); c.resolved = !c.resolved; markDirty(); renderComments(); return; }
  focusComment(id);
});
editor.addEventListener('click', ev => {
  const a = ev.target.closest('.cmt-anchor');
  if (a) focusComment(a.dataset.cmt); else if (!ev.target.closest('.cmt-card')) for (const c of $$('.cmt-card.active, .cmt-anchor.active')) c.classList.remove('active');
});

// ---------------------------------------------------------------- track changes
const TC = { on: false };
function setTracking(on) {
  TC.on = on;
  for (const b of $$('[data-cmd="trackChanges"]')) b.classList.toggle('on', on);
  toast(on ? 'Track Changes is on. Insertions and deletions are marked.' : 'Track Changes is off.');
}
function makeChange(tag) {
  const el = document.createElement(tag);
  el.className = 'tc';
  el.dataset.author = authorName();
  el.dataset.time = new Date().toISOString();
  el.style.setProperty('--tc', authorColor(el.dataset.author));
  return el;
}
const ownChange = (el, tag) => el && el.tagName === tag && el.dataset.author === authorName();
// mark a range as deleted (own insertions are simply removed); returns the created <del> elements
function markDeleted(range) {
  const dels = [];
  // pictures inside the range
  for (const img of $$('img', editor)) {
    if (!range.intersectsNode(img) || img.closest('del.tc')) continue;
    if (ownChange(img.closest('ins.tc'), 'INS')) { img.remove(); continue; }
    const d = makeChange('del');
    img.before(d);
    d.appendChild(img);
    dels.push(d);
  }
  for (const seg of textSegments(range).reverse()) {
    let node = seg.node;
    const host = node.parentElement && node.parentElement.closest('ins.tc, del.tc');
    if (host && host.tagName === 'DEL') continue;
    if (ownChange(host, 'INS')) {
      node.data = node.data.slice(0, seg.s) + node.data.slice(seg.e);
      if (!host.textContent && !host.querySelector('img')) host.remove();
      continue;
    }
    if (seg.e < node.length) node.splitText(seg.e);
    if (seg.s > 0) node = node.splitText(seg.s);
    // merge with a neighbouring deletion by the same author
    const prev = node.previousSibling, next = node.nextSibling;
    if (ownChange(next, 'DEL')) { next.prepend(node); dels.unshift(next); continue; }
    if (ownChange(prev, 'DEL')) { prev.appendChild(node); dels.unshift(prev); continue; }
    const d = makeChange('del');
    node.before(d);
    d.appendChild(node);
    dels.unshift(d);
  }
  return dels;
}
function insertTracked(text) {
  const sel = getSelection();
  let r = sel.getRangeAt(0);
  // never type inside deleted text
  const inDel = closestEl(r.startContainer, 'del.tc');
  if (inDel) { r = document.createRange(); r.setStartAfter(inDel); r.collapse(true); }
  const host = closestEl(r.startContainer, 'ins.tc');
  if (ownChange(host, 'INS') && r.startContainer.nodeType === 3) {
    const n = r.startContainer, off = r.startOffset;
    n.data = n.data.slice(0, off) + text + n.data.slice(off);
    r = document.createRange();
    r.setStart(n, off + text.length);
  } else {
    // extend an insertion we just made right before the caret
    const before = r.startContainer.nodeType === 3 && r.startOffset === 0 ? r.startContainer.previousSibling : r.startContainer.nodeType === 1 ? r.startContainer.childNodes[r.startOffset - 1] : null;
    if (ownChange(before, 'INS') && before.lastChild && before.lastChild.nodeType === 3) {
      before.lastChild.data += text;
      r = document.createRange();
      r.setStart(before.lastChild, before.lastChild.length);
    } else {
      const ins = makeChange('ins');
      ins.textContent = text;
      r.insertNode(ins);
      r = document.createRange();
      r.setStart(ins.firstChild, text.length);
    }
  }
  r.collapse(true);
  setRange(r);
}
editor.addEventListener('beforeinput', ev => {
  if (!TC.on) return;
  const t = ev.inputType;
  const ranges = ev.getTargetRanges ? ev.getTargetRanges() : [];
  const target = ranges[0];
  const toRange = sr => { const r = document.createRange(); r.setStart(sr.startContainer, sr.startOffset); r.setEnd(sr.endContainer, sr.endOffset); return r; };
  if (t === 'insertText' || t === 'insertReplacementText' || t === 'insertFromPaste' || t === 'insertFromDrop') {
    const text = ev.data != null ? ev.data : ev.dataTransfer ? ev.dataTransfer.getData('text/plain') : '';
    if (!text) return;
    ev.preventDefault();
    const sel = getSelection();
    let r = target ? toRange(target) : sel.getRangeAt(0);
    if (!r.collapsed) {
      const dels = markDeleted(r);
      r = document.createRange();
      if (dels.length) r.setStartAfter(dels[dels.length - 1]); else r.setStart(sel.anchorNode, sel.anchorOffset);
      r.collapse(true);
      setRange(r);
    } else setRange(r);
    insertTracked(text);
    markDirty(true);
    return;
  }
  if (/^delete(Content|Word|SoftLine|HardLine|ByCut)/.test(t)) {
    if (!target) return;
    ev.preventDefault();
    const backward = /Backward/.test(t);
    let r = toRange(target);
    if (r.collapsed) return;
    // only already-deleted text in the way: just move the caret over it
    const segs = textSegments(r);
    if (segs.length && segs.every(s => s.node.parentElement.closest('del.tc'))) {
      const d = segs[backward ? 0 : segs.length - 1].node.parentElement.closest('del.tc');
      const c = document.createRange();
      if (backward) c.setStartBefore(d); else c.setStartAfter(d);
      c.collapse(true);
      setRange(c);
      return;
    }
    const dels = markDeleted(r);
    const c = document.createRange();
    if (dels.length) { if (backward) c.setStartBefore(dels[0]); else c.setStartAfter(dels[dels.length - 1]); }
    else { const s = getSelection(); c.setStart(s.anchorNode, s.anchorOffset); }
    c.collapse(true);
    setRange(c);
    markDirty(true);
  }
}, true);
hooks.paste.push(html => (TC.on ? `<ins class="tc" data-author="${esc(authorName())}" data-time="${new Date().toISOString()}" style="--tc: ${authorColor(authorName())}">${html}</ins>` : html));

function allChanges() { return $$('ins.tc, del.tc', editor); }
function acceptChange(el) { if (el.tagName === 'INS') unwrap(el); else el.remove(); }
function rejectChange(el) { if (el.tagName === 'INS') el.remove(); else unwrap(el); }
function changeAtCaret() {
  const sel = getSelection();
  const here = sel.rangeCount && closestEl(sel.anchorNode, 'ins.tc, del.tc');
  if (here) return here;
  const cur = $('.tc-current', editor);
  return cur && editor.contains(cur) ? cur : null;
}
function stepChange(dir) {
  const all = allChanges();
  if (!all.length) { toast('There are no tracked changes.'); return null; }
  const cur = changeAtCaret();
  let i = cur ? all.indexOf(cur) + dir : (dir > 0 ? 0 : all.length - 1);
  if (!cur) {
    const sel = getSelection();
    if (sel.rangeCount) {
      const r = sel.getRangeAt(0);
      i = dir > 0 ? all.findIndex(el => r.comparePoint(el, 0) > 0) : [...all].reverse().findIndex(el => r.comparePoint(el, 0) < 0);
      if (dir < 0 && i >= 0) i = all.length - 1 - i;
      if (i < 0) i = dir > 0 ? 0 : all.length - 1;
    }
  }
  const el = all[(i + all.length) % all.length];
  for (const x of $$('.tc-current', editor)) x.classList.remove('tc-current');
  el.classList.add('tc-current');
  scrollToElement(el, false);
  const r = document.createRange();
  r.selectNodeContents(el);
  editor.focus({ preventScroll: true });
  setRange(r);
  return el;
}
function resolveChanges(kind, scope) {
  const apply = kind === 'accept' ? acceptChange : rejectChange;
  if (scope === 'all') {
    const all = allChanges();
    if (!all.length) { toast('There are no tracked changes.'); return; }
    for (const el of all) apply(el);
    toast(`${kind === 'accept' ? 'Accepted' : 'Rejected'} ${all.length} change${all.length === 1 ? '' : 's'}.`);
  } else {
    const el = changeAtCaret() || stepChange(1);
    if (!el) return;
    apply(el);
    if (scope === 'next') stepChange(1);
  }
  markDirty();
}
function setMarkup(mode) {
  document.body.classList.toggle('markup-simple', mode === 'simple');
  document.body.classList.toggle('markup-none', mode === 'none');
  state.markup = mode;
}

// ---------------------------------------------------------------- Font and Paragraph dialogs
function fontOptions(current) {
  const names = [...$('#font-family').options].map(o => o.value);
  if (current && !names.includes(current)) names.unshift(current);
  return names.map(n => ({ value: n, label: n }));
}
async function fontDialog() {
  restoreSelection();
  const sel = getSelection();
  if (!sel.rangeCount) return;
  const range = sel.getRangeAt(0).cloneRange();
  const el = styleAnchor();
  const cs = getComputedStyle(el);
  const q = c => { try { return document.queryCommandState(c); } catch { return false; } };
  const fam = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  const toHex = rgb => { const m = rgb.match(/\d+/g); return m ? `#${m.slice(0, 3).map(x => (+x).toString(16).padStart(2, '0')).join('')}` : '#000000'; };
  const spacing = parseFloat(cs.letterSpacing) || 0;
  const r = await askForm('Font', [
    { id: 'family', label: 'Font', type: 'select', options: fontOptions(fam), value: fam },
    { id: 'style', label: 'Font style', type: 'select', options: ['Regular', 'Italic', 'Bold', 'Bold Italic'].map(v => ({ value: v, label: v })), value: (q('bold') ? 'Bold' : '') + (q('bold') && q('italic') ? ' ' : '') + (q('italic') ? 'Italic' : '') || 'Regular' },
    { id: 'size', label: 'Size (pt)', type: 'number', value: Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2, attrs: 'min="1" max="1638" step="0.5"' },
    { id: 'color', label: 'Font colour', type: 'color', value: toHex(cs.color) },
    { id: 'underline', label: 'Underline', type: 'checkbox', value: q('underline') },
    { id: 'strike', label: 'Strikethrough', type: 'checkbox', value: q('strikeThrough') },
    { id: 'sup', label: 'Superscript', type: 'checkbox', value: q('superscript') },
    { id: 'sub', label: 'Subscript', type: 'checkbox', value: q('subscript') },
    { id: 'smallcaps', label: 'Small caps', type: 'checkbox', value: cs.fontVariantCaps === 'small-caps' },
    { id: 'allcaps', label: 'All caps', type: 'checkbox', value: cs.textTransform === 'uppercase' },
    { id: 'spacing', label: 'Character spacing (pt, negative = condensed)', type: 'number', value: Math.round(spacing * 0.75 * 10) / 10, attrs: 'step="0.1" min="-10" max="50"' },
  ], 'OK', null, true);
  if (!r) return;
  editor.focus({ preventScroll: true });
  setRange(range);
  if (range.collapsed) { toast('Select some text first.'); return; }
  withHistory(() => {
    document.execCommand('styleWithCSS', false, false);
    const want = { bold: /Bold/.test(r.style), italic: /Italic/.test(r.style), underline: r.underline, strikeThrough: r.strike, superscript: r.sup, subscript: r.sub && !r.sup };
    for (const [cmd, on] of Object.entries(want)) if (q(cmd) !== on) document.execCommand(cmd);
    document.execCommand('styleWithCSS', false, true);
    if (r.family) document.execCommand('fontName', false, r.family);
    document.execCommand('foreColor', false, r.color);
    const sp = parseFloat(r.spacing) || 0;
    const cur = getSelection().getRangeAt(0);
    wrapRangeText(cur, () => {
      const s = document.createElement('span');
      s.style.fontVariantCaps = r.smallcaps ? 'small-caps' : 'normal';
      s.style.textTransform = r.allcaps ? 'uppercase' : 'none';
      s.style.letterSpacing = sp ? `${sp}pt` : 'normal';
      return s;
    });
    cleanRedundantSpans(editor);
    const sizeVal = parseFloat(r.size);
    if (sizeVal) { setRange(range); applyFontSize(sizeVal); }
    markDirty();
  });
}
async function paragraphDialog() {
  restoreSelection();
  const blocks = getSelectedBlocks();
  if (!blocks.length) return;
  const range = getSelection().getRangeAt(0).cloneRange();
  const b = blocks[0];
  const cs = getComputedStyle(b);
  const inch = px => Math.round((parseFloat(px) / 96) * 100) / 100;
  const pt = px => Math.round(parseFloat(px) * 0.75 * 10) / 10;
  const ti = parseFloat(cs.textIndent) || 0;
  const lh = b.style.lineHeight || '';
  let lineMode = 'multiple', lineVal = 1.15;
  if (/pt$/.test(lh)) { lineMode = 'exact'; lineVal = parseFloat(lh); } else if (lh) { lineVal = parseFloat(lh); lineMode = { 1: 'single', 1.5: '1.5', 2: 'double' }[lineVal] || 'multiple'; } else lineMode = 'multiple';
  const r = await askForm('Paragraph', [
    { id: 'align', label: 'Alignment', type: 'select', value: cs.textAlign === 'start' ? 'left' : cs.textAlign, options: [['left', 'Left'], ['center', 'Centered'], ['right', 'Right'], ['justify', 'Justified']].map(([value, label]) => ({ value, label })) },
    { id: 'outline', label: 'Outline level', type: 'select', value: /^H[123]$/.test(b.tagName) ? b.tagName : 'P', options: [['P', 'Body text'], ['H1', 'Level 1 (Heading 1)'], ['H2', 'Level 2 (Heading 2)'], ['H3', 'Level 3 (Heading 3)']].map(([value, label]) => ({ value, label })) },
    { id: 'left', label: 'Indent left (in)', type: 'number', value: inch(cs.marginLeft), attrs: 'step="0.05" min="0" max="6"' },
    { id: 'right', label: 'Indent right (in)', type: 'number', value: inch(cs.marginRight), attrs: 'step="0.05" min="0" max="6"' },
    { id: 'special', label: 'Special indent', type: 'select', value: ti > 0 ? 'first' : ti < 0 ? 'hanging' : 'none', options: [['none', '(none)'], ['first', 'First line'], ['hanging', 'Hanging']].map(([value, label]) => ({ value, label })) },
    { id: 'by', label: 'By (in)', type: 'number', value: inch(Math.abs(ti)) || 0.5, attrs: 'step="0.05" min="0" max="4"' },
    { id: 'before', label: 'Spacing before (pt)', type: 'number', value: pt(cs.marginTop), attrs: 'step="1" min="0" max="200"' },
    { id: 'after', label: 'Spacing after (pt)', type: 'number', value: pt(cs.marginBottom), attrs: 'step="1" min="0" max="200"' },
    { id: 'line', label: 'Line spacing', type: 'select', value: lineMode, options: [['single', 'Single'], ['1.5', '1.5 lines'], ['double', 'Double'], ['multiple', 'Multiple'], ['exact', 'Exactly (pt)']].map(([value, label]) => ({ value, label })) },
    { id: 'at', label: 'At', type: 'number', value: lineVal, attrs: 'step="0.05" min="0.5" max="200"' },
  ], 'OK', null, true);
  if (!r) return;
  editor.focus({ preventScroll: true });
  setRange(range);
  withHistory(() => {
    let targets = getSelectedBlocks();
    if (r.outline !== (targets[0] && /^H[123]$/.test(targets[0].tagName) ? targets[0].tagName : 'P')) {
      document.execCommand('formatBlock', false, r.outline.toLowerCase());
      targets = getSelectedBlocks();
    }
    const px = v => `${Math.round(clamp(parseFloat(v) || 0, -10, 10) * 96)}px`;
    for (const t of targets) {
      t.style.textAlign = r.align === 'left' ? '' : r.align;
      const by = Math.abs(parseFloat(r.by) || 0);
      const left = parseFloat(r.left) || 0;
      t.style.marginLeft = left || r.special === 'hanging' ? px(left + (r.special === 'hanging' ? by : 0)) : '';
      t.style.marginRight = parseFloat(r.right) ? px(r.right) : '';
      t.style.textIndent = r.special === 'first' ? px(by) : r.special === 'hanging' ? px(-by) : '';
      t.style.marginTop = `${parseFloat(r.before) || 0}pt`;
      t.style.marginBottom = `${parseFloat(r.after) || 0}pt`;
      const at = parseFloat(r.at) || 1;
      t.style.lineHeight = { single: '1', '1.5': '1.5', double: '2', multiple: String(at), exact: `${at}pt` }[r.line];
      if (!t.getAttribute('style')) t.removeAttribute('style');
    }
    markDirty();
  });
}

// ---------------------------------------------------------------- watermark
async function watermarkDialog() {
  const w = state.page.watermark || {};
  const r = await askForm('Watermark', [
    { id: 'text', label: 'Text (leave empty for no watermark)', value: w.text ?? 'CONFIDENTIAL' },
    { id: 'color', label: 'Colour', type: 'color', value: w.color || '#c0c0c0' },
    { id: 'layout', label: 'Layout', type: 'select', value: w.layout || 'diagonal', options: [{ value: 'diagonal', label: 'Diagonal' }, { value: 'horizontal', label: 'Horizontal' }] },
    { id: 'opacity', label: 'Transparency', type: 'select', value: String(w.opacity ?? 0.5), options: [['0.25', 'Very light'], ['0.5', 'Semi-transparent'], ['0.8', 'Strong']].map(([value, label]) => ({ value, label })) },
  ], 'Apply', 'The watermark appears on every page on screen, when printing and in PDF exports.', true);
  if (!r) return;
  state.page.watermark = r.text.trim() ? { text: r.text.trim().slice(0, 60), color: r.color, layout: r.layout, opacity: parseFloat(r.opacity) } : null;
  applyWatermark();
  markDirty();
}
function watermarkStyle(wm, pageWidthPx) {
  const len = Math.max(4, wm.text.length);
  const size = Math.min(150, (pageWidthPx * (wm.layout === 'diagonal' ? 1.05 : 0.8)) / (len * 0.62));
  return `font-size:${size}px;color:${wm.color};opacity:${wm.opacity};--wm-rot:${wm.layout === 'diagonal' ? '-45deg' : '0deg'}`;
}
function applyWatermark() {
  state.sheetKey = '';
  const wm = state.page.watermark;
  const pw = $('#print-wm');
  if (pw) {
    pw.textContent = wm ? wm.text : '';
    pw.setAttribute('style', wm ? watermarkStyle(wm, state.pw * 96 * 0.75) : '');
  }
  layoutPages();
}
hooks.sheetKey.push(() => JSON.stringify(state.page.watermark || null));
hooks.sheet.push(() => {
  const wm = state.page.watermark;
  return wm ? `<div class="wm" style="${watermarkStyle(wm, state.pw * 96)}">${esc(wm.text)}</div>` : '';
});

// ---------------------------------------------------------------- links to headings
hooks.internalLink.push(href => {
  const m = href.match(/^#h:(.+)$/);
  const h = m && $$('[data-hid]', editor).find(x => x.dataset.hid === m[1]);
  if (h) scrollToElement(h); else toast('That heading no longer exists.');
});
async function insertLinkWithPlaces() {
  restoreSelection();
  const sel = getSelection();
  const text = sel.toString();
  const existing = closestEl(sel.anchorNode, 'a');
  const range = sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
  const hs = headingBlocks();
  const curHref = existing ? existing.getAttribute('href') || '' : '';
  const r = await askForm(existing ? 'Edit Link' : 'Insert Link', [
    { id: 'text', label: 'Text to display', value: existing ? existing.textContent : text },
    { id: 'url', label: 'Web page or e-mail address', value: curHref.startsWith('#') ? '' : curHref, placeholder: 'https://example.com' },
    { id: 'place', label: 'Or a place in this document', type: 'select', value: curHref.startsWith('#h:') ? curHref.slice(3) : '', options: [{ value: '', label: '(none)' }, ...hs.map(h => ({ value: headingId(h), label: `${'  '.repeat(Math.max(0, headingLevel(h) - 1))}${h.textContent.trim().slice(0, 60)}` }))] },
  ], existing ? 'Update' : 'Insert');
  if (!r) return;
  let url = r.place ? `#h:${r.place}` : r.url.trim();
  if (!url) return;
  if (!r.place && !/^(https?:|mailto:|#)/i.test(url)) url = /^[^\s/@]+@[^\s/@]+\.[^\s/@]+$/.test(url) ? `mailto:${url}` : `https://${url}`;
  const label = r.text || (r.place ? (hs.find(h => headingId(h) === r.place) || {}).textContent : url);
  editor.focus({ preventScroll: true });
  if (range) setRange(range);
  withHistory(() => {
    if (existing) { existing.setAttribute('href', url); if (r.text && r.text !== existing.textContent) existing.textContent = r.text; }
    else if (!text || r.text !== text) document.execCommand('insertHTML', false, `<a href="${esc(url)}">${esc(label)}</a>&nbsp;`);
    else document.execCommand('createLink', false, url);
    markDirty();
  });
}

// ---------------------------------------------------------------- saving, loading, export
hooks.extra.push(() => ({ comments: state.comments.filter(c => anchorsOf(c.id).length) }));
hooks.load.push(extra => {
  state.comments = Array.isArray(extra.comments) ? extra.comments.filter(c => c && c.id) : [];
  for (const el of $$('.toc', editor)) el.contentEditable = 'false';
  for (const el of $$('sup.fn-ref', editor)) el.contentEditable = 'false';
  for (const el of $$('.fn-sep', editor)) el.contentEditable = 'false';
  setTimeout(() => { renderComments(); applyWatermark(); scheduleNav(); }, 0);
});
hooks.sanitize.push(body => {
  for (const el of body.querySelectorAll('.toc, sup.fn-ref, .fn-sep')) el.setAttribute('contenteditable', 'false');
});
hooks.beforeSave.push(() => { if ($('.toc', editor)) updateToc(true); renumberFootnotes(); });
hooks.layout.push(() => { positionComments(); markCurrentHeading(); });
hooks.export.push((clone, extras) => {
  // table of contents -> a borderless two-column table (title, page)
  for (const toc of clone.querySelectorAll('.toc')) {
    const title = toc.querySelector('.toc-title');
    const rows = [...toc.querySelectorAll('.toc-entry')].map(e => `<tr><td style="border:none;padding:1px 4px 1px ${(+((e.className.match(/toc-l(\d)/) || [])[1] || 1) - 1) * 18}px">${esc(e.querySelector('.toc-text').textContent)}</td><td style="border:none;text-align:right;width:40px">${esc(e.querySelector('.toc-page').textContent)}</td></tr>`).join('');
    const frag = document.createElement('div');
    frag.innerHTML = `<p style="font-size:16pt;color:#2f5496;font-family:'Calibri Light'">${esc(title ? title.textContent : 'Contents')}</p>${rows ? `<table class="no-borders"><tbody>${rows}</tbody></table>` : ''}`;
    toc.replaceWith(...frag.childNodes);
  }
  if (!extras) return;
  // footnotes
  const notes = new Map([...clone.querySelectorAll('.footnotes li[data-fn]')].map(li => [li.dataset.fn, li.innerText.replace(/​/g, '').trim()]));
  [...clone.querySelectorAll('sup.fn-ref')].forEach((ref, i) => {
    extras.footnotes.push({ n: i + 1, text: notes.get(ref.dataset.fn) || '' });
    ref.replaceWith(document.createTextNode(`${MARK.fn}${i + 1}${MARK.fnEnd}`));
  });
  for (const fb of clone.querySelectorAll('.footnotes')) fb.remove();
  // comments
  let cn = 0;
  for (const c of state.comments) {
    const anchors = [...clone.querySelectorAll(`.cmt-anchor[data-cmt="${c.id}"]`)];
    if (!anchors.length) continue;
    const n = cn++;
    anchors[0].before(document.createTextNode(`${MARK.cmtStart}${n}${MARK.cmtStartEnd}`));
    anchors[anchors.length - 1].after(document.createTextNode(`${MARK.cmtEnd}${n}${MARK.cmtEndEnd}`));
    extras.comments.push({ n, author: c.author, initials: c.initials, date: c.time, text: c.text });
  }
  for (const a of clone.querySelectorAll('.cmt-anchor')) a.replaceWith(...a.childNodes);
  // tracked changes
  [...clone.querySelectorAll('ins.tc, del.tc')].forEach((el, n) => {
    const ins = el.tagName === 'INS';
    extras.changes.push({ n, type: ins ? 'ins' : 'del', author: el.dataset.author || authorName(), date: el.dataset.time || new Date().toISOString() });
    el.before(document.createTextNode(`${ins ? MARK.insStart : MARK.delStart}${n}${ins ? MARK.insStartEnd : MARK.delStartEnd}`));
    el.after(document.createTextNode(`${ins ? MARK.insEnd : MARK.delEnd}${n}${ins ? MARK.insEndEnd : MARK.delEndEnd}`));
    el.replaceWith(...el.childNodes);
  });
  // links to headings are not kept in .docx (Word needs bookmarks for them)
  for (const a of clone.querySelectorAll('a[href^="#"]')) a.replaceWith(...a.childNodes);
});
editor.addEventListener('input', () => {
  if ($$('sup.fn-ref', editor).length !== $$('.footnotes li', editor).length || $('.footnotes', editor)) renumberFootnotesSoon();
  scheduleNav();
  if (state.comments.length) renderCommentsSoon();
});
let fnTimer = null, cmtTimer = null;
function renumberFootnotesSoon() { clearTimeout(fnTimer); fnTimer = setTimeout(renumberFootnotes, 300); }
function renderCommentsSoon() { clearTimeout(cmtTimer); cmtTimer = setTimeout(renderComments, 250); }
// keep comment cards next to their text after undo/redo, zoom or window changes
new MutationObserver(() => { if (state.comments.length) renderCommentsSoon(); }).observe(editor, { childList: true, subtree: false });

// ---------------------------------------------------------------- options pane
$('#opt-username').addEventListener('input', ev => { state.userName = ev.target.value.trim() || 'Author'; api.setSetting('userName', ev.target.value.trim()); });
$('#opt-initials').addEventListener('input', ev => { state.initials = ev.target.value.trim(); api.setSetting('initials', ev.target.value.trim()); });
function fillOptions() { $('#opt-username').value = state.userName || ''; $('#opt-initials').value = state.initials || ''; }
document.addEventListener('click', ev => { if (ev.target.closest('[data-bs="options"]')) fillOptions(); });

// ---------------------------------------------------------------- ribbon wiring
Object.assign(COMMANDS, {
  navPane: () => toggleNavPane(),
  updateToc: () => updateToc(false),
  footnote: insertFootnote,
  nextFootnote,
  showNotes,
  newComment,
  deleteComment: () => { const id = currentCommentId(); if (id) deleteComment(id); else toast('Click a comment first.'); },
  deleteAllComments: async () => { if (!state.comments.length) { toast('There are no comments.'); return; } if (await confirmBox('Delete All Comments', `Delete all ${state.comments.length} comments?`, 'Delete', 'Cancel')) { for (const c of [...state.comments]) deleteComment(c.id); } },
  prevComment: () => stepComment(-1),
  nextComment: () => stepComment(1),
  showComments: () => { document.body.classList.toggle('hide-comments'); renderComments(); },
  trackChanges: () => setTracking(!TC.on),
  prevChange: () => stepChange(-1),
  nextChange: () => stepChange(1),
  fontDialog,
  paragraphDialog,
  watermark: watermarkDialog,
  link: insertLinkWithPlaces,
});
Object.assign(MENUS, {
  toc: () => [item('Automatic Table', insertToc), item('Update Table', () => updateToc(false)), '-', item('Remove Table of Contents', removeToc)],
  markup: () => [['all', 'All Markup'], ['simple', 'Simple Markup'], ['none', 'No Markup']].map(([m, l]) => item(l, () => setMarkup(m), { check: (state.markup || 'all') === m })),
  accept: () => [item('Accept This Change', () => resolveChanges('accept', 'one')), item('Accept and Move to Next', () => resolveChanges('accept', 'next')), '-', item('Accept All Changes', () => resolveChanges('accept', 'all'))],
  reject: () => [item('Reject This Change', () => resolveChanges('reject', 'one')), item('Reject and Move to Next', () => resolveChanges('reject', 'next')), '-', item('Reject All Changes', () => resolveChanges('reject', 'all'))],
});
// spacing menu: add the full dialog
const spacingMenu = MENUS.spacing;
MENUS.spacing = () => [...spacingMenu(), '-', item('Paragraph Settings…', paragraphDialog)];
document.addEventListener('keydown', ev => {
  if ($('#modal').open) return;
  const ctrl = ev.ctrlKey || ev.metaKey;
  if (!ctrl) return;
  const k = ev.key.toLowerCase();
  const run = fn => { ev.preventDefault(); ev.stopImmediatePropagation(); withHistory(fn); };
  if (ev.altKey && k === 'm') run(newComment);
  else if (ev.altKey && k === 'f') run(insertFootnote);
  else if (ev.shiftKey && k === 'e') { ev.preventDefault(); ev.stopImmediatePropagation(); setTracking(!TC.on); }
  else if (!ev.shiftKey && !ev.altKey && k === 'd' && ev.target === editor) { ev.preventDefault(); ev.stopImmediatePropagation(); fontDialog(); }
}, true);
for (const el of $$('#ribbon [data-icon]')) if (!el.querySelector('svg')) el.insertAdjacentHTML('afterbegin', svg(el.dataset.icon));
for (const el of $$('#bs-nav [data-icon]')) if (!el.querySelector('svg')) el.insertAdjacentHTML('afterbegin', svg(el.dataset.icon));
SHORTCUTS.push(['New comment', 'Ctrl+Alt+M'], ['Insert footnote', 'Ctrl+Alt+F'], ['Track Changes on/off', 'Ctrl+Shift+E'], ['Font dialog', 'Ctrl+D']);
$('#shortcuts').innerHTML = SHORTCUTS.map(([a, b]) => `<div><span>${esc(a)}</span><kbd>${esc(b)}</kbd></div>`).join('');

window.__doc = { TC, insertToc, updateToc, insertFootnote, renumberFootnotes, newComment, deleteComment, renderComments, setTracking, resolveChanges, stepChange, fontDialog, paragraphDialog, applyWatermark, toggleNavPane, markDeleted, insertTracked, MARK };
