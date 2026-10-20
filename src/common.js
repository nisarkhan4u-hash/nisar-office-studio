'use strict';
/* Nisar Office Studio: helpers shared by the document editor and the PDF editor */

const $ =(s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const clamp = (v, a, b) => Math.min(b, Math.max(a, v));

// ---------------------------------------------------------------- icons
const ICONS = {
  save: '<path d="M5 3h11l3 3v15H5z"/><path d="M8 3v5h8V3M8 21v-7h8v7"/>',
  saveas: '<path d="M5 3h11l3 3v7"/><path d="M5 3v18h7"/><path d="M8 3v5h8V3"/><path d="M14 21l1-3 5-5 2 2-5 5z"/>',
  undo: '<path d="M9 14L4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 010 11H11"/>',
  redo: '<path d="M15 14l5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 000 11H13"/>',
  paste: '<rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4.5V3h6v1.5M9 10h6M9 14h6M9 18h3"/>',
  cut: '<circle cx="6" cy="18" r="3"/><circle cx="18" cy="18" r="3"/><path d="M8.2 15.8L19 3M15.8 15.8L5 3"/>',
  copy: '<rect x="8" y="8" width="12" height="13" rx="2"/><path d="M16 8V5a2 2 0 00-2-2H6a2 2 0 00-2 2v11a2 2 0 002 2h2"/>',
  brush: '<rect x="4" y="3" width="13" height="6" rx="1"/><path d="M17 6h3v6h-8v3"/><rect x="10.5" y="15" width="3" height="6" rx="1"/>',
  clear: '<path d="M6 5h11M11.5 5L8 19"/><path d="M14 14l6 6M20 14l-6 6"/>',
  highlight: '<path d="M15 3l5 5-8 8H7v-5z"/><path d="M5 18l2-2"/>',
  shade: '<path d="M5 12l7-7 6 6-7 7z"/><path d="M18 11c1 1.5 2 2.7 2 3.7a2 2 0 01-4 0c0-1 1-2.2 2-3.7z"/>',
  alignLeft: '<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>',
  alignCenter: '<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>',
  alignRight: '<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>',
  justify: '<path d="M4 6h16M4 10h16M4 14h16M4 18h16"/>',
  bullets: '<circle cx="5" cy="6.5" r="1.4" fill="currentColor"/><circle cx="5" cy="12" r="1.4" fill="currentColor"/><circle cx="5" cy="17.5" r="1.4" fill="currentColor"/><path d="M9.5 6.5H20M9.5 12H20M9.5 17.5H20"/>',
  numbers: '<path d="M4 4.5h1.5V9M4 9h3"/><path d="M4 13c.5-.8 3-.9 3 .5 0 1.2-3 2.2-3 3.5h3"/><path d="M10 6.5h10M10 12h10M10 17.5h10"/>',
  indent: '<path d="M4 5h16M11 10h9M11 14h9M4 19h16M4 9l3 3-3 3"/>',
  outdent: '<path d="M4 5h16M11 10h9M11 14h9M4 19h16M7 9l-3 3 3 3"/>',
  spacing: '<path d="M11 6h9M11 12h9M11 18h9M5.5 4v16M3 6.5L5.5 4 8 6.5M3 17.5L5.5 20 8 17.5"/>',
  find: '<circle cx="10.5" cy="10.5" r="6"/><path d="M15 15l5.5 5.5"/>',
  replace: '<path d="M4 8h12l-3-3M20 16H8l3 3"/>',
  selectAll: '<path d="M4 8V4h4M16 4h4v4M20 16v4h-4M8 20H4v-4"/><path d="M8 9h8M8 12h8M8 15h5"/>',
  pagebreak: '<path d="M6 3h12v6H6zM6 21h12v-6H6z"/><path d="M3 12h2M8 12h2M14 12h2M19 12h2"/>',
  table: '<rect x="3" y="4" width="18" height="16" rx="1.5"/><path d="M3 9.5h18M3 14.8h18M9 4v16M15 4v16"/>',
  image: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="8.5" cy="9.5" r="1.8"/><path d="M21 16l-5-5-9 9"/>',
  link: '<path d="M10 14a4 4 0 005.66 0l3-3a4 4 0 00-5.66-5.66l-1 1"/><path d="M14 10a4 4 0 00-5.66 0l-3 3a4 4 0 005.66 5.66l1-1"/>',
  hr: '<path d="M3 12h18"/><path d="M6 7h12M6 17h12" opacity=".35"/>',
  date: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4M7.5 14h2M11 14h2M14.5 14h2M7.5 17.5h2M11 17.5h2"/>',
  headerfooter: '<rect x="4" y="3" width="16" height="18" rx="1"/><path d="M4 7.5h16M4 16.5h16" stroke-dasharray="2 1.5"/><path d="M8 11h8M8 13.5h5"/>',
  pagenum: '<path d="M6 3h9l4 4v14H6z"/><path d="M15 3v4h4"/><path d="M11 13.5c.4-.8 3-.9 3 .6 0 1.2-3 2-3 3.4h3"/>',
  margins: '<rect x="4" y="3" width="16" height="18"/><path d="M8 3v18M16 3v18M4 7h16M4 17h16" stroke-dasharray="1.5 1.5"/>',
  orientation: '<rect x="3" y="8" width="11" height="13" rx="1"/><rect x="9" y="3" width="12" height="9" rx="1"/>',
  size: '<path d="M6 3h9l4 4v14H6z"/><path d="M15 3v4h4"/><path d="M9 17l6-6M15 14v-3h-3"/>',
  columns: '<path d="M4 5h7M4 9h7M4 13h7M4 17h7M13 5h7M13 9h7M13 13h7M13 17h7"/>',
  rowAbove: '<rect x="3" y="12" width="18" height="8" rx="1"/><path d="M12 3v6M9 6h6"/>',
  rowBelow: '<rect x="3" y="4" width="18" height="8" rx="1"/><path d="M12 15v6M9 18h6"/>',
  colLeft: '<rect x="12" y="3" width="8" height="18" rx="1"/><path d="M3 12h6M6 9v6"/>',
  colRight: '<rect x="4" y="3" width="8" height="18" rx="1"/><path d="M15 12h6M18 9v6"/>',
  trash: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>',
  spell: '<path d="M3.5 15L7 5l3.5 10M4.8 11.5h4.4"/><path d="M13 5v10h3a2.5 2.5 0 000-5h-3m3 0a2.5 2.5 0 000-5h-3"/><path d="M13 19.5l2.5 2 5-5.5"/>',
  count: '<path d="M8 4L6 20M16 4l-2 16M4 9h16M3.5 15h16"/>',
  speak: '<path d="M4 9h4l5-4v14l-5-4H4z"/><path d="M16 9a4 4 0 010 6M18.5 6.5a7.5 7.5 0 010 11"/>',
  stop: '<rect x="6" y="6" width="12" height="12" rx="1.5"/>',
  zoom100: '<circle cx="10.5" cy="10.5" r="6"/><path d="M15 15l5.5 5.5M8 9l1.5-1v5"/><path d="M12 8.5h1.2v4.5"/>',
  onepage: '<rect x="6" y="3" width="12" height="18" rx="1"/><path d="M9 7h6M9 10h6M9 13h6M9 16h4"/>',
  pagewidth: '<path d="M3 12h18M6.5 8.5L3 12l3.5 3.5M17.5 8.5L21 12l-3.5 3.5"/><path d="M3 4v4M21 4v4M3 16v4M21 16v4"/>',
  moon: '<path d="M20 14.5A8 8 0 019.5 4 8 8 0 1020 14.5z"/>',
  focus: '<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/><circle cx="12" cy="12" r="2.5"/>',
  web: '<circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3c3 3 3 15 0 18M12 3c-3 3-3 15 0 18"/>',
  printlayout: '<path d="M6 3h12v18H6z"/><path d="M9 7h6M9 10h6M9 13h6M9 16h4"/>',
  boundaries: '<rect x="4" y="3" width="16" height="18" rx="1"/><path d="M4 12h16" stroke-dasharray="2 2"/>',
  print: '<path d="M7 9V3h10v6"/><rect x="3" y="9" width="18" height="8" rx="2"/><path d="M7 14h10v7H7z"/>',
  preview: '<path d="M6 3h9l4 4v6"/><path d="M6 3v18h6"/><path d="M15 3v4h4"/><circle cx="16.5" cy="16.5" r="3"/><path d="M18.7 18.7L21 21"/>',
  pdf: '<path d="M6 3h9l4 4v14H6z"/><path d="M15 3v4h4"/><path d="M9 12h6M9 15h6M9 18h4"/>',
  new: '<path d="M6 3h9l4 4v14H6z"/><path d="M15 3v4h4M12.5 11v6M9.5 14h6"/>',
  open: '<path d="M3 7a2 2 0 012-2h4l2 2h8a2 2 0 012 2v9a2 2 0 01-2 2H5a2 2 0 01-2-2z"/>',
  doc: '<path d="M6 3h9l4 4v14H6z"/><path d="M15 3v4h4M9 11h6M9 14h6M9 17h4"/>',
  home: '<path d="M3 11l9-7 9 7"/><path d="M5 10v10h5v-6h4v6h5V10"/>',
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7.5v.5"/>',
  close: '<path d="M6 6l12 12M18 6L6 18"/>',
  back: '<path d="M19 12H5M11 6l-6 6 6 6"/>',
  chevron: '<path d="M6 9l6 6 6-6"/>',
  chevronUp: '<path d="M6 15l6-6 6 6"/>',
  up: '<path d="M6 15l6-6 6 6"/>',
  down: '<path d="M6 9l6 6 6-6"/>',
  wrapLeft: '<rect x="3" y="6" width="8" height="8" rx="1"/><path d="M14 6h7M14 10h7M14 14h7M3 18h18"/>',
  wrapRight: '<rect x="13" y="6" width="8" height="8" rx="1"/><path d="M3 6h7M3 10h7M3 14h7M3 18h18"/>',
};
const svg = name => (ICONS[name] ? `<svg class="ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name]}</svg>` : '');

// ---------------------------------------------------------------- menus
// Each page defines MENUS (name -> items or { build }) and may set runMenuAction
// to wrap menu actions (the document editor makes each one an undo step).
let runMenuAction = fn => fn();
let reportError = err => { console.error(err); toast(`Something went wrong: ${err && err.message ? err.message : err}`); };
const COLORS = [
  '#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47',
  '#f2f2f2', '#7f7f7f', '#d0cece', '#d6dce4', '#d9e2f3', '#fbe5d5', '#ededed', '#fff2cc', '#deebf6', '#e2efd9',
  '#d8d8d8', '#595959', '#aeabab', '#adb9ca', '#b4c6e7', '#f7cbac', '#dbdbdb', '#fee599', '#bdd7ee', '#c5e0b3',
  '#bfbfbf', '#3f3f3f', '#757070', '#8496b0', '#8eaadb', '#f4b183', '#c9c9c9', '#ffd965', '#9cc3e5', '#a8d08d',
  '#a5a5a5', '#262626', '#3a3838', '#323f4f', '#2f5496', '#c55a11', '#7b7b7b', '#bf9000', '#2e75b5', '#538135',
  '#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0',
];

let openMenuEl = null;
function closeMenu() {
  if (openMenuEl) { openMenuEl.remove(); openMenuEl = null; }
}
function item(label, fn, opts = {}) {
  return { label, fn, ...opts };
}
function menuHtml(items) {
  return items.map((it, i) => {
    if (it === '-') return '<hr>';
    if (it.head) return `<div class="mhead">${esc(it.head)}</div>`;
    if (it.html) return it.html;
    return `<button class="item" data-i="${i}">${it.check !== undefined ? `<span class="ck">${it.check ? '✓' : ''}</span>` : ''}${it.pre || ''}<span>${esc(it.label)}</span>${it.hint ? `<small>${esc(it.hint)}</small>` : ''}</button>`;
  }).join('');
}
function colorMenu(apply, autoLabel, list = COLORS, hl = false) {
  return {
    build: el => {
      el.innerHTML = `${autoLabel ? `<button class="item" data-auto>${esc(autoLabel)}</button><hr>` : ''}<div class="swatches${hl ? ' hl' : ''}">${list.map(c => `<button class="sw" data-c="${c}" style="background:${c}" title="${c}"></button>`).join('')}</div>${hl ? '' : '<hr><button class="item" data-more>More Colors…</button>'}`;
      el.addEventListener('click', e => {
        const sw = e.target.closest('[data-c]');
        if (sw) { closeMenu(); runMenuAction(() => apply(sw.dataset.c)); return; }
        if (e.target.closest('[data-auto]')) { closeMenu(); runMenuAction(() => apply(null)); return; }
        if (e.target.closest('[data-more]')) {
          closeMenu();
          const inp = document.createElement('input');
          inp.type = 'color';
          inp.addEventListener('change', () => runMenuAction(() => apply(inp.value)));
          inp.click();
        }
      });
    },
  };
}

function openMenu(name, anchor) {
  const wasOpen = openMenuEl && openMenuEl.dataset.name === name;
  closeMenu();
  if (wasOpen) return;
  const spec = MENUS[name] && MENUS[name]();
  if (!spec) return;
  const el = document.createElement('div');
  el.className = 'menu';
  el.dataset.name = name;
  if (Array.isArray(spec)) {
    el.innerHTML = menuHtml(spec);
    el.addEventListener('click', e => {
      const b = e.target.closest('[data-i]');
      if (!b) return;
      closeMenu();
      try { const r = runMenuAction(spec[+b.dataset.i].fn); if (r && r.catch) r.catch(reportError); } catch (err) { reportError(err); }
    });
  } else spec.build(el);
  $('#menu-layer').appendChild(el);
  const r = anchor.closest('.split, .combo, button').getBoundingClientRect();
  const left = clamp(r.left, 4, innerWidth - el.offsetWidth - 4);
  let top = r.bottom + 2;
  if (top + el.offsetHeight > innerHeight - 4) top = Math.max(4, r.top - el.offsetHeight - 2);
  el.style.left = `${left}px`;
  el.style.top = `${top}px`;
  openMenuEl = el;
}

// ---------------------------------------------------------------- dialogs & toast
function askForm(title, fields, okLabel = 'OK', note = null, twoCol = false) {
  return new Promise(resolve => {
    const d = $('#modal');
    const field = f => {
      if (f.type === 'checkbox') return `<label class="chk"><input type="checkbox" name="${f.id}" ${f.value ? 'checked' : ''}> ${esc(f.label)}</label>`;
      if (f.type === 'select') return `<label class="field">${esc(f.label)}<select name="${f.id}">${(f.options || []).map(o => `<option value="${esc(o.value)}"${String(o.value) === String(f.value ?? '') ? ' selected' : ''}>${esc(o.label)}</option>`).join('')}</select></label>`;
      return `<label class="field">${esc(f.label)}<input name="${f.id}" type="${f.type || 'text'}" value="${esc(f.value ?? '')}" placeholder="${esc(f.placeholder || '')}" ${f.attrs || ''} spellcheck="false"></label>`;
    };
    d.innerHTML = `<form><h2>${esc(title)}</h2>${note ? `<p>${esc(note)}</p>` : ''}<div class="${twoCol ? 'grid2' : 'col'}" style="gap:12px">${fields.map(field).join('')}</div><div class="actions"><button type="button" class="btn" data-cancel>Cancel</button><button type="submit" class="btn primary">${esc(okLabel)}</button></div></form>`;
    const form = $('form', d);
    let done = false;
    form.addEventListener('submit', e => {
      e.preventDefault();
      done = true;
      const out = {};
      for (const f of fields) { const i = form.elements[f.id]; out[f.id] = f.type === 'checkbox' ? i.checked : i.value; }
      d.close();
      resolve(out);
    });
    $('[data-cancel]', d).addEventListener('click', () => d.close());
    d.addEventListener('close', () => { if (!done) resolve(null); }, { once: true });
    d.showModal();
    const first = $('input:not([type=checkbox])', d);
    if (first) { first.focus(); first.select(); }
  });
}
function alertBox(title, body, isHtml = false) {
  return new Promise(resolve => {
    const d = $('#modal');
    d.innerHTML = `<div class="dbody"><h2>${esc(title)}</h2>${isHtml ? body : `<p>${esc(body)}</p>`}<div class="actions"><button class="btn primary" data-ok>OK</button></div></div>`;
    $('[data-ok]', d).addEventListener('click', () => d.close());
    d.addEventListener('close', () => resolve(), { once: true });
    d.showModal();
    $('[data-ok]', d).focus();
  });
}
function confirmBox(title, body, okLabel, cancelLabel) {
  return new Promise(resolve => {
    const d = $('#modal');
    let ok = false;
    d.innerHTML = `<div class="dbody"><h2>${esc(title)}</h2><p>${esc(body)}</p><div class="actions"><button class="btn" data-no>${esc(cancelLabel)}</button><button class="btn primary" data-yes>${esc(okLabel)}</button></div></div>`;
    $('[data-yes]', d).addEventListener('click', () => { ok = true; d.close(); });
    $('[data-no]', d).addEventListener('click', () => d.close());
    d.addEventListener('close', () => resolve(ok), { once: true });
    d.showModal();
    $('[data-yes]', d).focus();
  });
}
let toastTimer = null;
function toast(msg, action, actionLabel) {
  const t = $('#toast');
  t.innerHTML = `<span>${esc(msg)}</span>${action ? `<button>${esc(actionLabel)}</button>` : ''}`;
  if (action) $('button', t).addEventListener('click', () => { t.hidden = true; action(); });
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.hidden = true; }, action ? 6000 : 2800);
}
