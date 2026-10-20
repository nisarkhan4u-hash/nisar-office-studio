'use strict';
// Converts a .docx (Office Open XML) into Nisar Office Studio's editor HTML, keeping the
// direct formatting that general-purpose converters drop: fonts, sizes, colours,
// highlights, alignment, indents, spacing, lists, tables, pictures, links,
// page breaks, page setup, headers and footers.
const { DOMParser } = require('@xmldom/xmldom');

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  v: 'urn:schemas-microsoft-com:vml',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
};
const EDITOR_FONT = 'Calibri';
const EDITOR_SIZE = 11;
const HIGHLIGHT = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000',
  darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};
const LIST_TYPES = { decimal: 'decimal', lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha', lowerRoman: 'lower-roman', upperRoman: 'upper-roman', decimalZero: 'decimal-leading-zero' };
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };

const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
// font names go inside style="…", so quote them with single quotes
const cssFont = f => { const n = String(f).replace(/['"\\<>;]/g, ''); return /[\s,\d]/.test(n) ? `'${n}'` : n; };
const parse = s => new DOMParser({ onError: () => {} }).parseFromString(s, 'text/xml');

// --- small XML helpers (WordprocessingML namespace unless stated)
const isEl = (n, local, ns = NS.w) => n && n.nodeType === 1 && n.localName === local && n.namespaceURI === ns;
const kids = (el, local, ns = NS.w) => { const out = []; if (!el) return out; for (let n = el.firstChild; n; n = n.nextSibling) if (!local ? n.nodeType === 1 : isEl(n, local, ns)) out.push(n); return out; };
const kid = (el, local, ns = NS.w) => kids(el, local, ns)[0] || null;
const attr = (el, name, ns = NS.w) => { if (!el) return null; const v = el.getAttributeNS(ns, name); return v === '' || v == null ? null : v; };
const val = el => attr(el, 'val');
const onOff = el => { if (!el) return undefined; const v = val(el); return !(v === '0' || v === 'false' || v === 'off' || v === 'none'); };
const descendants = (el, local, ns = NS.w) => (el ? Array.from(el.getElementsByTagNameNS(ns, local)) : []);
const twipToPx = t => Math.round((+t / 15) * 100) / 100;
const twipToPt = t => Math.round((+t / 20) * 100) / 100;
const hex = c => (c && /^[0-9a-f]{6}$/i.test(c) ? `#${c.toLowerCase()}` : null);

function parseRels(xml) {
  const map = {};
  if (!xml) return map;
  for (const r of Array.from(parse(xml).getElementsByTagName('Relationship'))) {
    map[r.getAttribute('Id')] = { target: r.getAttribute('Target'), type: r.getAttribute('Type') || '', external: r.getAttribute('TargetMode') === 'External' };
  }
  return map;
}

// --- run & paragraph properties
function readRPr(rPr, theme) {
  const p = {};
  if (!rPr) return p;
  const set = (k, v) => { if (v !== undefined && v !== null) p[k] = v; };
  set('b', onOff(kid(rPr, 'b')));
  set('i', onOff(kid(rPr, 'i')));
  const u = kid(rPr, 'u');
  if (u) p.u = val(u) !== 'none';
  set('strike', onOff(kid(rPr, 'strike')) || onOff(kid(rPr, 'dstrike')));
  set('caps', onOff(kid(rPr, 'caps')));
  set('smallCaps', onOff(kid(rPr, 'smallCaps')));
  set('hidden', onOff(kid(rPr, 'vanish')));
  const va = val(kid(rPr, 'vertAlign'));
  if (va) p.va = va === 'superscript' ? 'sup' : va === 'subscript' ? 'sub' : '';
  const color = kid(rPr, 'color');
  if (color) { const c = val(color); p.color = c === 'auto' ? '#000000' : hex(c) || p.color; }
  const sz = val(kid(rPr, 'sz'));
  if (sz) p.size = +sz / 2;
  const f = kid(rPr, 'rFonts');
  if (f) {
    const themeName = attr(f, 'asciiTheme') || attr(f, 'hAnsiTheme');
    const name = attr(f, 'ascii') || attr(f, 'hAnsi') || (themeName ? (themeName.startsWith('major') ? theme.major : theme.minor) : null) || attr(f, 'cs') || attr(f, 'eastAsia');
    if (name) p.font = name;
  }
  const hl = val(kid(rPr, 'highlight'));
  if (hl) p.bg = hl === 'none' ? '' : HIGHLIGHT[hl] || p.bg;
  const shd = kid(rPr, 'shd');
  if (shd && !p.bg) { const fill = hex(attr(shd, 'fill')); if (fill && fill !== '#ffffff') p.bg = fill; }
  return p;
}
function readPPr(pPr) {
  const p = {};
  if (!pPr) return p;
  const jc = val(kid(pPr, 'jc'));
  if (jc) p.align = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[jc] || 'left';
  const ind = kid(pPr, 'ind');
  if (ind) {
    const left = attr(ind, 'left') ?? attr(ind, 'start');
    if (left != null) p.indLeft = twipToPx(left);
    const right = attr(ind, 'right') ?? attr(ind, 'end');
    if (right != null) p.indRight = twipToPx(right);
    if (attr(ind, 'firstLine') != null) p.firstLine = twipToPx(attr(ind, 'firstLine'));
    if (attr(ind, 'hanging') != null) p.firstLine = -twipToPx(attr(ind, 'hanging'));
  }
  const sp = kid(pPr, 'spacing');
  if (sp) {
    if (attr(sp, 'before') != null && attr(sp, 'beforeAutospacing') !== '1') p.before = twipToPt(attr(sp, 'before'));
    if (attr(sp, 'after') != null && attr(sp, 'afterAutospacing') !== '1') p.after = twipToPt(attr(sp, 'after'));
    const line = attr(sp, 'line');
    const rule = attr(sp, 'lineRule') || 'auto';
    // "at least" spacing is a minimum, so normal text height already satisfies it
    if (line != null && rule !== 'atLeast') p.line = rule === 'auto' ? `${Math.round((+line / 240) * 100) / 100}` : `${twipToPt(line)}pt`;
  }
  const shd = kid(pPr, 'shd');
  if (shd) { const fill = hex(attr(shd, 'fill')); if (fill && fill !== '#ffffff') p.shade = fill; }
  if (onOff(kid(pPr, 'pageBreakBefore'))) p.pageBreakBefore = true;
  const numPr = kid(pPr, 'numPr');
  if (numPr) {
    const id = val(kid(numPr, 'numId'));
    if (id != null) p.numId = id;
    const lvl = val(kid(numPr, 'ilvl'));
    if (lvl != null) p.ilvl = +lvl;
  }
  if (kid(pPr, 'pBdr')) p.hasBorder = true;
  return p;
}

function parseStyles(xml, theme) {
  const out = { byId: {}, defaultPara: null, docR: {}, docP: {} };
  if (!xml) return out;
  const d = parse(xml).documentElement;
  const dd = kid(d, 'docDefaults');
  if (dd) {
    out.docR = readRPr(kid(kid(dd, 'rPrDefault'), 'rPr'), theme);
    out.docP = readPPr(kid(kid(dd, 'pPrDefault'), 'pPr'));
  }
  for (const s of kids(d, 'style')) {
    const id = attr(s, 'styleId');
    if (!id) continue;
    out.byId[id] = {
      id,
      type: attr(s, 'type'),
      name: (val(kid(s, 'name')) || id).toLowerCase(),
      basedOn: val(kid(s, 'basedOn')),
      r: readRPr(kid(s, 'rPr'), theme),
      p: readPPr(kid(s, 'pPr')),
    };
    if (attr(s, 'type') === 'paragraph' && attr(s, 'default') === '1') out.defaultPara = id;
  }
  return out;
}
function resolveStyle(styles, id, part, seen = new Set()) {
  const s = id && styles.byId[id];
  if (!s || seen.has(id)) return {};
  seen.add(id);
  return { ...resolveStyle(styles, s.basedOn, part, seen), ...s[part] };
}
function styleName(styles, id, seen = new Set()) {
  // the first recognised name up the basedOn chain (so "My Heading" based on Heading 1 is a heading)
  const s = id && styles.byId[id];
  if (!s || seen.has(id)) return '';
  seen.add(id);
  if (/^(heading [1-6]|title|subtitle|quote|intense quote)$/.test(s.name)) return s.name;
  // localised Word versions keep English-like style ids (Heading1, Title…)
  const id2 = s.id.match(/^heading\s?([1-6])$/i);
  if (id2) return `heading ${id2[1]}`;
  if (/^(title|subtitle|quote|intensequote)$/i.test(s.id)) return s.id.toLowerCase().replace('intensequote', 'intense quote');
  return styleName(styles, s.basedOn, seen);
}

function parseNumbering(xml) {
  const abstract = {};
  const nums = {};
  if (!xml) return { abstract, nums };
  const d = parse(xml).documentElement;
  for (const a of kids(d, 'abstractNum')) {
    const levels = {};
    for (const l of kids(a, 'lvl')) {
      levels[+attr(l, 'ilvl')] = { fmt: val(kid(l, 'numFmt')) || 'decimal', text: val(kid(l, 'lvlText')) || '', start: +(val(kid(l, 'start')) || 1) };
    }
    abstract[attr(a, 'abstractNumId')] = levels;
  }
  for (const n of kids(d, 'num')) {
    const overrides = {};
    for (const o of kids(n, 'lvlOverride')) {
      const l = kid(o, 'lvl');
      if (l) overrides[+attr(o, 'ilvl')] = { fmt: val(kid(l, 'numFmt')) || 'decimal', text: val(kid(l, 'lvlText')) || '' };
    }
    nums[attr(n, 'numId')] = { abstractId: val(kid(n, 'abstractNumId')), overrides };
  }
  return { abstract, nums };
}
function listInfo(numbering, numId, ilvl) {
  const n = numbering.nums[numId];
  if (!n || numId === '0') return null;
  const lvl = n.overrides[ilvl] || (numbering.abstract[n.abstractId] || {})[ilvl] || { fmt: 'decimal', text: '' };
  if (lvl.fmt === 'none') return null;
  if (lvl.fmt === 'bullet') {
    const ch = lvl.text;
    return { tag: 'ul', type: ch === 'o' || ch === '◦' ? 'circle' : ch === '' || ch === '▪' || ch === '■' ? 'square' : '' };
  }
  return { tag: 'ol', type: LIST_TYPES[lvl.fmt] || 'decimal', start: lvl.start };
}

// --- converter
class Converter {
  constructor(zip, ctx) {
    this.zip = zip;
    Object.assign(this, ctx);
    this.images = 0;
  }

  async imageHtml(relId, cx, cy, partRels) {
    const rel = (partRels || this.rels)[relId];
    if (!rel || rel.external) return '';
    const path = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`.replace(/\/[^/]+\/\.\.\//g, '/');
    const file = this.zip.file(path);
    if (!file) return '';
    const ext = path.split('.').pop().toLowerCase();
    const mime = MIME[ext];
    if (!mime) return ''; // EMF/WMF drawings can't be shown in the editor
    const b64 = await file.async('base64');
    const w = cx ? Math.round(cx / 9525) : 0;
    const h = cy ? Math.round(cy / 9525) : 0;
    this.images++;
    void h;
    return `<img src="data:${mime};base64,${b64}" alt=""${w ? ` style="width: ${w}px; height: auto;"` : ''}>`;
  }

  async drawing(el) {
    const blip = descendants(el, 'blip', NS.a)[0];
    const extent = descendants(el, 'extent', NS.wp)[0];
    if (blip) return this.imageHtml(attr(blip, 'embed', NS.r), extent ? +extent.getAttribute('cx') : 0, extent ? +extent.getAttribute('cy') : 0);
    const imgData = descendants(el, 'imagedata', NS.v)[0];
    if (imgData) return this.imageHtml(attr(imgData, 'id', NS.r), 0, 0);
    return '';
  }

  runCss(eff, base) {
    const css = [];
    if (eff.font && eff.font !== base.font) css.push(`font-family: ${cssFont(eff.font)}`);
    if (eff.size && eff.size !== base.size) css.push(`font-size: ${eff.size}pt`);
    if (eff.color && eff.color !== (base.color || '#000000')) css.push(`color: ${eff.color}`);
    if (eff.bg) css.push(`background-color: ${eff.bg}`);
    if (eff.caps) css.push('text-transform: uppercase');
    if (eff.smallCaps) css.push('font-variant: small-caps');
    return css.join('; ');
  }

  // returns an array of { key, html } segments for the inline content of a paragraph
  // comment ranges -> <span class="cmt-anchor">, continued across paragraphs when needed
  commentMarks(html) {
    const stack = this.openComments;
    const openTag = id => `<span class="cmt-anchor" data-cmt="c${id}">`;
    let out = stack.map(openTag).join('');
    out += html.replace(/\u0000C([SE])(\d+)\u0000/g, (m, t, id) => {
      if (!this.comments.has(id)) return '';
      if (t === 'S') { stack.push(id); return openTag(id); }
      const i = stack.lastIndexOf(id);
      if (i < 0) return '';
      const inner = stack.slice(i + 1);
      stack.splice(i);
      stack.push(...inner);
      return '</span>'.repeat(inner.length + 1) + inner.map(openTag).join('');
    });
    return out + '</span>'.repeat(stack.length);
  }

  async inline(el, base, out, state) {
    for (let n = el.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const ln = n.localName;
      if (n.namespaceURI === NS.mc && ln === 'AlternateContent') {
        const choice = kid(n, 'Choice', NS.mc) || kid(n, 'Fallback', NS.mc);
        if (choice) await this.inline(choice, base, out, state);
        continue;
      }
      if (n.namespaceURI !== NS.w) continue;
      if (ln === 'r') await this.run(n, base, out, state);
      else if (ln === 'hyperlink') {
        const rel = this.rels[attr(n, 'id', NS.r)];
        const anchor = attr(n, 'anchor');
        const href = rel && rel.external ? rel.target : anchor ? `#${anchor}` : null;
        const inner = [];
        await this.inline(n, base, inner, state);
        const html = joinSegments(inner);
        if (href && /^(https?:|mailto:|#)/i.test(href)) out.push({ key: null, html: `<a href="${esc(href)}">${html}</a>` });
        else out.push({ key: null, html });
      } else if (ln === 'ins' || ln === 'del' || ln === 'moveFrom' || ln === 'moveTo') {
        // tracked changes stay tracked: <ins class="tc"> / <del class="tc">
        const tag = ln === 'ins' || ln === 'moveTo' ? 'ins' : 'del';
        const inner = [];
        const wasDel = state.deleted;
        state.deleted = tag === 'del';
        await this.inline(n, base, inner, state);
        state.deleted = wasDel;
        const html = joinSegments(inner);
        const author = attr(n, 'author') || 'Author';
        if (html) out.push({ key: null, html: `<${tag} class="tc" data-author="${esc(author)}" data-time="${esc(attr(n, 'date') || '')}">${html}</${tag}>` });
      } else if (ln === 'commentRangeStart') {
        out.push({ key: null, html: `\u0000CS${attr(n, 'id')}\u0000` });
      } else if (ln === 'commentRangeEnd') {
        out.push({ key: null, html: `\u0000CE${attr(n, 'id')}\u0000` });
      } else if (['smartTag', 'customXml', 'sdt', 'sdtContent', 'fldSimple', 'dir', 'bdo'].includes(ln)) {
        await this.inline(ln === 'sdt' ? kid(n, 'sdtContent') || n : n, base, out, state);
      }
      // bookmarks and proofing marks: skipped
    }
  }

  async run(r, base, out, state) {
    const rStyle = val(kid(kid(r, 'rPr'), 'rStyle'));
    const eff = { ...base, ...resolveStyle(this.styles, rStyle, 'r'), ...readRPr(kid(r, 'rPr'), this.theme) };
    if (eff.hidden) return;
    const tags = [];
    if (eff.b) tags.push('b');
    if (eff.i) tags.push('i');
    if (eff.u) tags.push('u');
    if (eff.strike) tags.push('s');
    if (eff.va) tags.push(eff.va);
    const css = this.runCss(eff, base);
    const key = `${tags.join(',')}|${css}`;
    const wrap = text => {
      let h = text;
      for (const t of [...tags].reverse()) h = `<${t}>${h}</${t}>`;
      return css ? `<span style="${css}">${h}</span>` : h;
    };
    for (let n = r.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const ln = n.localName;
      if (n.namespaceURI === NS.mc && ln === 'AlternateContent') {
        const choice = kid(n, 'Choice', NS.mc) || kid(n, 'Fallback', NS.mc);
        if (choice) { const h = await this.drawing(choice); if (h) out.push({ key: null, html: h }); this.collectTextBoxes(choice, state); }
        continue;
      }
      if (n.namespaceURI !== NS.w) continue;
      if (ln === 'fldChar') {
        const t = attr(n, 'fldCharType');
        if (t === 'begin') state.field = 'code';
        else if (t === 'separate') state.field = 'result';
        else if (t === 'end') state.field = null;
        continue;
      }
      if (state.field === 'code') { if (ln === 'instrText') state.instr = (state.instr || '') + n.textContent; continue; }
      const text = t => out.push({ key, text: t, wrap });
      if (ln === 'footnoteReference' || ln === 'endnoteReference') {
        const id = attr(n, 'id');
        const note = (ln === 'footnoteReference' ? this.footnotes : this.endnotes).get(id);
        if (note != null) {
          const fid = `${ln === 'footnoteReference' ? 'f' : 'e'}${id}`;
          this.noteOrder.push({ fid, text: note });
          out.push({ key: null, html: `<sup class="fn-ref" data-fn="${fid}" contenteditable="false">${this.noteOrder.length}</sup>` });
        }
        continue;
      }
      if (ln === 'delText' && state.deleted) { text(esc(n.textContent)); continue; }
      if (ln === 't') text(esc(n.textContent).replace(/ {2,}/g, m => ' ' + ' '.repeat(m.length - 1)));
      else if (ln === 'tab' || ln === 'ptab') text('  ');
      else if (ln === 'br') {
        if (attr(n, 'type') === 'page') out.push({ key: null, html: '\u0000PAGEBREAK\u0000' });
        else out.push({ key: null, html: '<br>' });
      } else if (ln === 'cr') out.push({ key: null, html: '<br>' });
      else if (ln === 'noBreakHyphen') text('‑');
      else if (ln === 'sym') {
        const code = parseInt(attr(n, 'char') || '', 16);
        if (code) text(esc(String.fromCodePoint(code >= 0xf000 ? code - 0xf000 : code)));
      } else if (ln === 'drawing' || ln === 'pict' || ln === 'object') {
        const h = await this.drawing(n);
        if (h) out.push({ key: null, html: h });
        this.collectTextBoxes(n, state);
      }
    }
  }

  collectTextBoxes(el, state) {
    for (const tb of descendants(el, 'txbxContent')) state.textBoxes.push(tb);
  }

  paraBase(pStyle) {
    return { ...this.styles.docR, ...resolveStyle(this.styles, pStyle, 'r') };
  }

  // one <w:p> -> { html, list }
  async paragraph(p) {
    const pPrEl = kid(p, 'pPr');
    const pStyle = val(kid(pPrEl, 'pStyle')) || this.styles.defaultPara;
    const ppr = { ...this.styles.docP, ...resolveStyle(this.styles, pStyle, 'p'), ...readPPr(pPrEl) };
    const name = styleName(this.styles, pStyle);
    const base = this.paraBase(pStyle);
    const state = { field: null, textBoxes: [] };
    const segs = [];
    await this.inline(p, base, segs, state);
    let inner = this.commentMarks(joinSegments(segs));

    let tag = 'p', cls = '';
    const hm = name.match(/^heading ([1-6])$/);
    if (hm) tag = `h${Math.min(4, +hm[1])}`;
    else if (name === 'title') { tag = 'h1'; cls = 'doc-title'; }
    else if (name === 'subtitle') cls = 'doc-subtitle';
    else if (name === 'quote' || name === 'intense quote') tag = 'blockquote';

    const css = [];
    if (ppr.align && ppr.align !== 'left') css.push(`text-align: ${ppr.align}`);
    const list = ppr.numId != null ? listInfo(this.numbering, ppr.numId, ppr.ilvl || 0) : null;
    if (!list) {
      if (ppr.indLeft) css.push(`margin-left: ${ppr.indLeft}px`);
      if (ppr.firstLine) css.push(`text-indent: ${ppr.firstLine}px`);
    }
    if (ppr.indRight) css.push(`margin-right: ${ppr.indRight}px`);
    if (tag === 'p' && !cls) {
      if (ppr.before) css.push(`margin-top: ${ppr.before}pt`);
      if (ppr.after != null && Math.abs(ppr.after - 8) > 0.01) css.push(`margin-bottom: ${ppr.after}pt`);
      if (ppr.line && ppr.line !== '1.15') css.push(`line-height: ${ppr.line}`);
      // the paragraph's own font, if it isn't the editor default
      if (base.font && base.font !== EDITOR_FONT) css.push(`font-family: ${cssFont(base.font)}`);
      if (base.size && base.size !== EDITOR_SIZE) css.push(`font-size: ${base.size}pt`);
      if (base.color && base.color !== '#000000') css.push(`color: ${base.color}`);
    }
    if (ppr.shade) css.push(`background-color: ${ppr.shade}`);

    // a page break inside the paragraph splits it into two around a page-break block
    const parts = inner.split('\u0000PAGEBREAK\u0000');
    const attrs = `${cls ? ` class="${cls}"` : ''}${css.length ? ` style="${css.join('; ')}"` : ''}`;
    const blocks = parts.map((part, i) => {
      const hasText = part.replace(/<[^>]+>/g, '').trim() || /<img|<br>/.test(part);
      if (!hasText && parts.length > 1) return ''; // a paragraph holding only a page break
      return `<${list ? 'li' : tag}${list ? (css.length ? ` style="${css.join('; ')}"` : '') : attrs}>${hasText ? part : '<br>'}</${list ? 'li' : tag}>`;
    });
    const PB = '<div class="page-break" contenteditable="false"></div>';
    let html = blocks.map((b, i) => (i > 0 ? PB : '') + b).join('');
    if (ppr.pageBreakBefore && !list) html = PB + html;
    const after = [];
    for (const tb of state.textBoxes) after.push(await this.blocks(tb));
    return { html, list: list ? { ...list, level: ppr.ilvl || 0, numId: ppr.numId } : null, after: after.join('') };
  }

  async table(tbl) {
    const tblPr = kid(tbl, 'tblPr');
    const tStyle = val(kid(tblPr, 'tblStyle')) || '';
    const borders = kid(tblPr, 'tblBorders');
    const styleHasGrid = /grid|table ?grid|light|medium|dark|list/i.test(tStyle) || /grid/i.test((this.styles.byId[tStyle] || {}).name || '');
    const bordersNone = borders && kids(borders).every(b => ['none', 'nil'].includes(val(b)));
    const noBorders = bordersNone || (!borders && !styleHasGrid);
    const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map(g => +attr(g, 'w') || 0);
    const total = grid.reduce((a, b) => a + b, 0);
    const rows = kids(tbl, 'tr');
    // map cells to grid columns to work out vertical merges
    const matrix = rows.map(tr => {
      let col = 0;
      return kids(tr, 'tc').map(tc => {
        const tcPr = kid(tc, 'tcPr');
        const span = +(val(kid(tcPr, 'gridSpan')) || 1);
        const vm = kid(tcPr, 'vMerge');
        const cell = { tc, col, span, vMerge: vm ? (val(vm) === 'restart' ? 'restart' : 'continue') : null };
        col += span;
        return cell;
      });
    });
    let html = `<table${noBorders ? ' class="no-borders"' : ''}>`;
    if (total > 0 && grid.length) html += `<colgroup>${grid.map(w => `<col style="width: ${Math.round((w / total) * 1000) / 10}%">`).join('')}</colgroup>`;
    html += '<tbody>';
    for (let ri = 0; ri < matrix.length; ri++) {
      html += '<tr>';
      for (const cell of matrix[ri]) {
        if (cell.vMerge === 'continue') continue;
        let rowspan = 1;
        if (cell.vMerge === 'restart') {
          for (let rj = ri + 1; rj < matrix.length; rj++) {
            const below = matrix[rj].find(c => c.col === cell.col);
            if (below && below.vMerge === 'continue') rowspan++; else break;
          }
        }
        const tcPr = kid(cell.tc, 'tcPr');
        const css = [];
        const shd = kid(tcPr, 'shd');
        const fill = shd && hex(attr(shd, 'fill'));
        if (fill && fill !== '#ffffff') css.push(`background-color: ${fill}`);
        const vAlign = val(kid(tcPr, 'vAlign'));
        if (vAlign === 'center') css.push('vertical-align: middle');
        else if (vAlign === 'bottom') css.push('vertical-align: bottom');
        const content = await this.blocks(cell.tc);
        html += `<td${cell.span > 1 ? ` colspan="${cell.span}"` : ''}${rowspan > 1 ? ` rowspan="${rowspan}"` : ''}${css.length ? ` style="${css.join('; ')}"` : ''}>${content || '<br>'}</td>`;
      }
      html += '</tr>';
    }
    return `${html}</tbody></table>`;
  }

  // block content of body / table cell / text box, with list grouping
  async blocks(container) {
    let html = '';
    const stack = []; // open lists, innermost last: { tag, level, numId }
    const closeTop = () => { const t = stack.pop(); html += `</li></${t.tag}>`; };
    const closeAll = () => { while (stack.length) closeTop(); };
    for (let n = container.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      if (isEl(n, 'p')) {
        const para = await this.paragraph(n);
        if (para.list) {
          const L = para.list;
          const item = para.html.replace(/<\/li>$/, ''); // left open for nested lists
          while (stack.length && stack[stack.length - 1].level > L.level) closeTop();
          const top = stack[stack.length - 1];
          if (top && top.level === L.level && top.tag === L.tag && top.numId === L.numId) {
            html += `</li>${item}`;
          } else {
            if (top && top.level === L.level) closeTop();
            const start = L.tag === 'ol' && L.start > 1 ? ` start="${L.start}"` : '';
            const type = L.type && L.type !== 'decimal' ? ` style="list-style-type: ${L.type}"` : '';
            html += `<${L.tag}${start}${type}>${item}`; // nested inside the open item when deeper
            stack.push({ tag: L.tag, level: L.level, numId: L.numId });
          }
          if (para.after) { closeAll(); html += para.after; }
        } else {
          closeAll();
          html += para.html + para.after;
        }
      } else if (isEl(n, 'tbl')) {
        closeAll();
        html += await this.table(n);
      } else if (isEl(n, 'sdt')) {
        closeAll();
        const c = kid(n, 'sdtContent');
        if (c) html += await this.blocks(c);
      } else if (isEl(n, 'customXml') || isEl(n, 'ins')) {
        closeAll();
        html += await this.blocks(n);
      } else if (n.namespaceURI === NS.mc && n.localName === 'AlternateContent') {
        const c = kid(n, 'Choice', NS.mc) || kid(n, 'Fallback', NS.mc);
        if (c) { closeAll(); html += await this.blocks(c); }
      }
    }
    closeAll();
    return html;
  }
}

function joinSegments(segs) {
  // merge neighbouring runs that share the same formatting: <b>ab</b> instead of <b>a</b><b>b</b>
  const out = [];
  for (const s of segs) {
    const prev = out[out.length - 1];
    if (prev && s.key != null && prev.key === s.key) { prev.text += s.text; continue; }
    out.push({ ...s });
  }
  return out.map(s => (s.key != null ? s.wrap(s.text) : s.html)).join('');
}

async function partText(zip, relTarget) {
  const f = zip.file(`word/${relTarget}`);
  if (!f) return { text: '', page: false };
  const d = parse(await f.async('string'));
  const instr = descendants(d.documentElement, 'instrText').map(n => n.textContent).join(' ') + ' ' + descendants(d.documentElement, 'fldSimple').map(n => attr(n, 'instr') || '').join(' ');
  const page = /\bPAGE\b/.test(instr);
  // visible text, without field results (page numbers etc.)
  const paras = descendants(d.documentElement, 'p').map(p => {
    let inField = null;
    let s = '';
    for (const r of descendants(p, 'r')) {
      for (let n = r.firstChild; n; n = n.nextSibling) {
        if (isEl(n, 'fldChar')) { const t = attr(n, 'fldCharType'); inField = t === 'begin' ? 'code' : t === 'separate' ? 'result' : null; }
        else if (isEl(n, 't') && !inField && !(n.parentNode && n.parentNode.parentNode && isEl(n.parentNode.parentNode, 'fldSimple'))) s += n.textContent;
        else if (isEl(n, 'tab')) s += ' ';
      }
    }
    return s.replace(/\s+/g, ' ').trim();
  }).filter(Boolean);
  let text = paras.join(' · ');
  if (page) text = text.replace(/\b(Page|Pg\.?)\s*(of)?\s*$/i, '').replace(/\s*[·|-]\s*$/, '').trim();
  return { text, page };
}

async function convertDocx(zip, parsePageSetup) {
  const read = async p => { const f = zip.file(p); return f ? f.async('string') : null; };
  const docXml = await read('word/document.xml');
  if (!docXml) throw new Error('This file is not a Word document (word/document.xml is missing).');
  const themeXml = await read('word/theme/theme1.xml');
  const theme = { minor: 'Calibri', major: 'Calibri Light' };
  if (themeXml) {
    const t = parse(themeXml).documentElement;
    const minor = descendants(t, 'minorFont', NS.a)[0];
    const major = descendants(t, 'majorFont', NS.a)[0];
    const latin = el => el && kid(el, 'latin', NS.a) && kid(el, 'latin', NS.a).getAttribute('typeface');
    theme.minor = latin(minor) || theme.minor;
    theme.major = latin(major) || theme.major;
  }
  const ctx = {
    theme,
    rels: parseRels(await read('word/_rels/document.xml.rels')),
    styles: parseStyles(await read('word/styles.xml'), theme),
    numbering: parseNumbering(await read('word/numbering.xml')),
  };
  // comments, footnotes and endnotes live in their own parts
  const partFor = type => { const r = Object.values(ctx.rels).find(x => x.type.endsWith(`/${type}`)); return r ? read(`word/${r.target.replace(/^\/?word\//, '')}`) : null; };
  const paraText = el => descendants(el, 'p').map(p => descendants(p, 't').map(t => t.textContent).join('')).join('\n').trim();
  const comments = new Map();
  const cxml = await partFor('comments');
  if (cxml) for (const c of descendants(parse(cxml).documentElement, 'comment')) {
    comments.set(attr(c, 'id'), { id: `c${attr(c, 'id')}`, author: attr(c, 'author') || 'Author', initials: attr(c, 'initials') || '', time: attr(c, 'date') || new Date().toISOString(), text: paraText(c), resolved: false });
  }
  const notesOf = async (type, tag) => {
    const m = new Map();
    const x = await partFor(type);
    if (x) for (const f of descendants(parse(x).documentElement, tag)) if (!attr(f, 'type') || attr(f, 'type') === 'normal') m.set(attr(f, 'id'), paraText(f));
    return m;
  };
  const conv = new Converter(zip, ctx);
  conv.comments = comments;
  conv.openComments = [];
  conv.footnotes = await notesOf('footnotes', 'footnote');
  conv.endnotes = await notesOf('endnotes', 'endnote');
  conv.noteOrder = [];
  const doc = parse(docXml);
  const body = descendants(doc.documentElement, 'body')[0];
  let html = await conv.blocks(body);
  if (conv.noteOrder.length) {
    html += `<div class="footnotes"><hr class="fn-sep" contenteditable="false"><ol>${conv.noteOrder.map(n => `<li data-fn="${n.fid}">${esc(n.text) || '<br>'}</li>`).join('')}</ol></div>`;
  }
  const usedComments = [...comments.values()].filter(c => html.includes(`data-cmt="${c.id}"`));

  const page = (parsePageSetup && parsePageSetup(docXml)) || {};
  const sect = descendants(body, 'sectPr').pop();
  if (sect) {
    for (const ref of kids(sect, 'headerReference')) if ((attr(ref, 'type') || 'default') === 'default') {
      const rel = ctx.rels[attr(ref, 'id', NS.r)];
      if (rel) { const h = await partText(zip, rel.target); page.header = h.text; if (h.page) page.pageNumbers = true; }
    }
    for (const ref of kids(sect, 'footerReference')) if ((attr(ref, 'type') || 'default') === 'default') {
      const rel = ctx.rels[attr(ref, 'id', NS.r)];
      if (rel) { const f = await partText(zip, rel.target); page.footer = f.text; if (f.page) page.pageNumbers = true; }
    }
  }
  return { html, page, extra: { comments: usedComments } };
}

module.exports = { convertDocx };
