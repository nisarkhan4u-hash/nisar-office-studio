'use strict';
// Turns the private-use markers that the editor puts into exported text into real Word
// features: comments (word/comments.xml), tracked changes (<w:ins>/<w:del>) and
// footnotes (word/footnotes.xml). html-to-docx can't produce these itself.

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MARKER = /([])([^]*)([])/;
const MARKER_G = new RegExp(MARKER.source, 'g');
const ANY_MARK = /[-]/;

const xmlEsc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const isoDate = d => { const t = new Date(d); return Number.isNaN(t.getTime()) ? new Date().toISOString().replace(/\.\d+Z$/, 'Z') : t.toISOString().replace(/\.\d+Z$/, 'Z'); };

// split one <w:r> whose text contains markers into plain runs and marker tokens
function splitRun(run) {
  const tm = run.match(/<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/);
  if (!tm || !ANY_MARK.test(tm[2])) return [{ run }];
  const rPr = (run.match(/<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>/) || [''])[0];
  const out = [];
  let rest = tm[2];
  let m;
  while ((m = rest.match(MARKER))) {
    const before = rest.slice(0, m.index);
    if (before) out.push({ run: `<w:r>${rPr}<w:t xml:space="preserve">${before}</w:t></w:r>` });
    out.push({ marker: m[1], value: m[2] });
    rest = rest.slice(m.index + m[0].length);
  }
  if (rest) out.push({ run: `<w:r>${rPr}<w:t xml:space="preserve">${rest}</w:t></w:r>` });
  return out;
}
const asDeleted = run => run.replace(/<w:t(\s[^>]*)?>/g, '<w:delText xml:space="preserve">').replace(/<\/w:t>/g, '</w:delText>');

function rewriteBody(xml, extras) {
  const changes = new Map(extras.changes.map(c => [String(c.n), c]));
  let revId = 1000;
  const open = []; // wrappers still open across paragraphs: { type, n }
  const wrapOpen = w => { const c = changes.get(w.n) || {}; return `<w:${w.type} w:id="${revId++}" w:author="${xmlEsc(c.author || 'Author')}" w:date="${isoDate(c.date)}">`; };
  return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>|<w:p\b[^>]*\/>/g, para => {
    if (!ANY_MARK.test(para) && !open.length) return para;
    if (para.endsWith('/>')) return para;
    const startTag = para.match(/^<w:p\b[^>]*>/)[0];
    let inner = para.slice(startTag.length, -'</w:p>'.length);
    let pPr = '';
    const pm = inner.match(/^\s*<w:pPr>[\s\S]*?<\/w:pPr>|^\s*<w:pPr\/>/);
    if (pm) { pPr = pm[0]; inner = inner.slice(pm[0].length); }
    let out = open.map(wrapOpen).join(''); // continue changes that started in an earlier paragraph
    const inDel = () => open.some(w => w.type === 'del');
    const parts = inner.split(/(<w:r>[\s\S]*?<\/w:r>|<w:r\s[^>]*>[\s\S]*?<\/w:r>)/);
    for (const part of parts) {
      if (!part) continue;
      if (!/^<w:r[\s>]/.test(part)) { out += part; continue; }
      for (const tok of splitRun(part)) {
        if (tok.run) { out += inDel() ? asDeleted(tok.run) : tok.run; continue; }
        switch (tok.marker) {
          case '': out += `<w:commentRangeStart w:id="${tok.value}"/>`; break;
          case '': out += `<w:commentRangeEnd w:id="${tok.value}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${tok.value}"/></w:r>`; break;
          case '': { const w = { type: 'ins', n: tok.value }; open.push(w); out += wrapOpen(w); break; }
          case '': { const w = { type: 'del', n: tok.value }; open.push(w); out += wrapOpen(w); break; }
          case '':
          case '': {
            const type = tok.marker === '' ? 'ins' : 'del';
            const i = open.map(w => w.type).lastIndexOf(type);
            if (i >= 0) { for (let k = open.length - 1; k >= i; k--) out += `</w:${open[k].type}>`; const reopen = open.splice(i); reopen.shift(); for (const w of reopen) { open.push(w); out += wrapOpen(w); } }
            break;
          }
          case '': out += `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${tok.value}"/></w:r>`; break;
        }
      }
    }
    out += [...open].reverse().map(w => `</w:${w.type}>`).join(''); // close at the paragraph end
    return `${startTag}${pPr}${out}</w:p>`;
  });
}

function textParagraphs(text, firstRun = '') {
  const lines = String(text || '').split(/\r?\n/);
  return lines.map((l, i) => `<w:p>${i === 0 ? firstRun : ''}<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${xmlEsc(i === 0 && firstRun ? ` ${l}` : l)}</w:t></w:r></w:p>`).join('');
}
function commentsXml(list) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${W_NS}">${list.map(c => `<w:comment w:id="${c.n}" w:author="${xmlEsc(c.author || 'Author')}" w:date="${isoDate(c.date)}" w:initials="${xmlEsc(c.initials || '')}">${textParagraphs(c.text || ' ')}</w:comment>`).join('')}</w:comments>`;
}
function footnotesXml(list) {
  const sep = (type, id, inner) => `<w:footnote w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r>${inner}</w:r></w:p></w:footnote>`;
  const notes = list.map(f => `<w:footnote w:id="${f.n}">${textParagraphs(f.text || ' ', '<w:r><w:rPr><w:vertAlign w:val="superscript"/><w:sz w:val="18"/></w:rPr><w:footnoteRef/></w:r>')}</w:footnote>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="${W_NS}">${sep('separator', -1, '<w:separator/>')}${sep('continuationSeparator', 0, '<w:continuationSeparator/>')}${notes}</w:footnotes>`;
}
async function addPart(zip, name, xml, relType, contentType, relId) {
  zip.file(`word/${name}`, xml);
  let rels = await zip.file('word/_rels/document.xml.rels').async('string');
  if (!rels.includes(`Target="${name}"`)) rels = rels.replace('</Relationships>', `<Relationship Id="${relId}" Type="${REL}/${relType}" Target="${name}"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', rels);
  let types = await zip.file('[Content_Types].xml').async('string');
  if (!types.includes(`PartName="/word/${name}"`)) types = types.replace('</Types>', `<Override PartName="/word/${name}" ContentType="${contentType}"/></Types>`);
  zip.file('[Content_Types].xml', types);
}

// returns the rewritten document.xml; adds comments/footnotes parts to the zip
async function applyDocxExtras(zip, xml, extras) {
  if (!extras || (!extras.comments.length && !extras.changes.length && !extras.footnotes.length)) return xml.replace(MARKER_G, '');
  let out = rewriteBody(xml, extras).replace(/[-]/g, '');
  if (extras.comments.length) await addPart(zip, 'comments.xml', commentsXml(extras.comments), 'comments', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml', 'rIdNisarComments');
  if (extras.footnotes.length) {
    await addPart(zip, 'footnotes.xml', footnotesXml(extras.footnotes), 'footnotes', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml', 'rIdNisarFootnotes');
    const s = zip.file('word/settings.xml');
    if (s) {
      let settings = await s.async('string');
      if (!settings.includes('<w:footnotePr>')) settings = settings.replace(/(<w:settings\b[^>]*>)/, '$1<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>');
      zip.file('word/settings.xml', settings);
    }
  }
  return out;
}

module.exports = { applyDocxExtras };
