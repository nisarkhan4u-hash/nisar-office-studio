/* Nisar Office Studio: PDF editor (renderer process, ES module)
 *
 * Pages are drawn by pdf.js. Everything the user adds (text, pictures, shapes,
 * highlights, notes, signatures…) lives in an editable layer above each page and
 * is written into the PDF with pdf-lib when saving.
 *
 * Coordinates: each added object has a local box (w × h, y pointing down, units =
 * PDF points) and a matrix `m` mapping that box into the page's PDF user space.
 * A page's viewport matrix maps PDF space onto the screen, so objects follow page
 * rotation and zoom exactly, and the same matrix is used when writing the PDF.
 */
import * as pdfjsLib from '../node_modules/pdfjs-dist/build/pdf.min.mjs';

const PDFJS = new URL('../node_modules/pdfjs-dist/', import.meta.url).href;
pdfjsLib.GlobalWorkerOptions.workerSrc = `${PDFJS}build/pdf.worker.min.mjs`;
const PDF_OPTS = {
  cMapUrl: `${PDFJS}cmaps/`, cMapPacked: true,
  standardFontDataUrl: `${PDFJS}standard_fonts/`,
  wasmUrl: `${PDFJS}wasm/`, iccUrl: `${PDFJS}iccs/`,
  isEvalSupported: false, enableXfa: false,
};
const PL = window.PDFLib;
const PT = 96 / 72;                 // screen pixels per PDF point at 100 %
const EDIT_FILE = 'nisarstudio-edits.json';
const CSS_FONT = { Helvetica: 'Arial, Helvetica, sans-serif', Times: '"Times New Roman", Times, serif', Courier: '"Courier New", Courier, monospace' };
const BASELINE = { Helvetica: 0.9465, Times: 0.9375, Courier: 0.8665 }; // first baseline / font size at line-height 1.2
const TEXT_PAD = 2;
const MARKUP = new Set(['highlight', 'underline', 'strike']);

Object.assign(ICONS, {
  cursor: '<path d="M6 3l12 9-5.5 1.2L16 20l-2.6 1.2-2.6-6.6L7 18z"/>',
  hand: '<path d="M8 13V6.5a1.5 1.5 0 013 0V12M11 11.5V5a1.5 1.5 0 013 0v6.5M14 11.5V6.5a1.5 1.5 0 013 0V14c0 4-2.5 7-6.5 7-2.7 0-4.2-1.2-5.6-3.4L3.4 14a1.5 1.5 0 012.5-1.6L8 15"/>',
  underlineTool: '<path d="M7 4v7a5 5 0 0010 0V4M5 20h14"/>',
  strikeTool: '<path d="M16 6.5C15 5 13.6 4 12 4c-2.5 0-4.5 1.5-4.5 3.6 0 1.6 1.1 2.6 3 3.2M4 12h16M8 17c.9 1.8 2.4 3 4.4 3 2.6 0 4.6-1.5 4.6-3.6 0-.6-.1-1.1-.4-1.6"/>',
  note: '<path d="M4 4h16v12H9l-5 4z"/><path d="M8 9h8M8 12h5"/>',
  pen: '<path d="M4 20c2-1 3.5-1.6 5-3.6L18.5 7a2.1 2.1 0 00-3-3L6 13.5C4 15 4.5 17 4 20z"/><path d="M14 5.5l4.5 4.5"/>',
  rectTool: '<rect x="4" y="6" width="16" height="12" rx="1"/>',
  ellipseTool: '<ellipse cx="12" cy="12" rx="8.5" ry="6"/>',
  lineTool: '<path d="M5 19L19 5"/>',
  arrowTool: '<path d="M5 19L19 5M10 5h9v9"/>',
  widthTool: '<path d="M4 6h16" stroke-width="1"/><path d="M4 11h16" stroke-width="2.2"/><path d="M4 17.5h16" stroke-width="3.6"/>',
  opacityTool: '<circle cx="9" cy="12" r="6"/><circle cx="15" cy="12" r="6" opacity=".45"/>',
  addText: '<path d="M5 6V4h14v2M12 4v16M9 20h6"/>',
  editText: '<path d="M4 6V4h10v2M9 4v14M6.5 18h5"/><path d="M14 20l1-3 5-5 2 2-5 5z"/>',
  sign: '<path d="M3 17c3-6 5-9 6.5-9 2 0-2 9 .5 9 1.5 0 2.5-4 4-4 1.2 0 .5 3 2 3 1 0 2-1 3-2"/><path d="M3 21h18"/>',
  stamp: '<path d="M9 13V9.5a3 3 0 116 0V13"/><rect x="4" y="13" width="16" height="4" rx="1"/><path d="M6 20h12"/>',
  whiteout: '<rect x="3" y="7" width="18" height="10" rx="1.5"/><path d="M7 12h10" opacity=".35"/>',
  redact: '<rect x="3" y="7" width="18" height="10" rx="1.5" fill="currentColor"/>',
  flatten: '<path d="M4 8l8-4 8 4-8 4z"/><path d="M4 12l8 4 8-4M4 16l8 4 8-4"/>',
  merge: '<path d="M7 3h7l4 4v8H7z"/><path d="M4 7v14h10"/><path d="M14 3v4h4"/>',
  rotateLeft: '<path d="M4 4v5h5"/><path d="M4.5 9A8 8 0 1112 20"/>',
  rotateRight: '<path d="M20 4v5h-5"/><path d="M19.5 9A8 8 0 1012 20"/>',
  extract: '<path d="M6 3h9l4 4v6"/><path d="M6 3v18h6"/><path d="M15 3v4h4"/><path d="M15 17h7M19 14l3 3-3 3"/>',
  goto: '<path d="M6 3h9l4 4v14H6z"/><path d="M15 3v4h4"/><path d="M9 14h6M12.5 11.5L15 14l-2.5 2.5"/>',
  thumbs: '<rect x="3" y="3" width="7" height="9" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><path d="M14 5h7M14 9h7M14 16h7M14 20h4"/>',
  headerfooter: '<rect x="4" y="3" width="16" height="18" rx="1"/><path d="M4 7.5h16M4 16.5h16" stroke-dasharray="2 1.5"/><path d="M8 11h8M8 13.5h5"/>',
  watermark: '<rect x="5" y="3" width="14" height="18" rx="1"/><path d="M8 16l8-8" opacity=".5"/><path d="M8 12l4-4" opacity=".5"/>',
  crop: '<path d="M6 2v14a2 2 0 002 2h14"/><path d="M2 6h14a2 2 0 012 2v14"/>',
  split: '<path d="M6 3h5v18H6zM13 3h5v18h-5z"/><path d="M12 1v22" stroke-dasharray="2 2"/>',
  fieldText: '<rect x="3" y="7" width="18" height="10" rx="1.5"/><path d="M7 10v4"/>',
  fieldCheck: '<rect x="5" y="5" width="14" height="14" rx="2"/><path d="M8.5 12l2.5 2.5 4.5-5"/>',
  fieldDropdown: '<rect x="3" y="7" width="18" height="10" rx="1.5"/><path d="M15 11l2 2 2-2"/>',
  ocr: '<path d="M4 8V4h4M16 4h4v4M20 16v4h-4M8 20H4v-4"/><path d="M8 9h8M8 12h8M8 15h5"/>',
  lock: '<rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 018 0v3"/>',
  unlock: '<rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 017.5-2"/>',
  findRedact: '<circle cx="10" cy="10" r="5.5"/><path d="M14 14l6 6"/><rect x="7" y="8.5" width="6" height="3" fill="currentColor"/>',
  sanitize: '<path d="M4 20l7-7M14 4l6 6-8 8-6-6z"/><path d="M3 21h8"/>',
  compress: '<path d="M8 3v6H2M16 3v6h6M8 21v-6H2M16 21v-6h6"/>',
});

// ---------------------------------------------------------------- matrices
// [a, b, c, d, e, f] means x' = a·x + c·y + e, y' = b·x + d·y + f
const mul = (A, B) => [A[0] * B[0] + A[2] * B[1], A[1] * B[0] + A[3] * B[1], A[0] * B[2] + A[2] * B[3], A[1] * B[2] + A[3] * B[3], A[0] * B[4] + A[2] * B[5] + A[4], A[1] * B[4] + A[3] * B[5] + A[5]];
const inv = A => { const d = A[0] * A[3] - A[1] * A[2]; return [A[3] / d, -A[1] / d, -A[2] / d, A[0] / d, (A[2] * A[5] - A[3] * A[4]) / d, (A[1] * A[4] - A[0] * A[5]) / d]; };
const apply = (A, x, y) => [A[0] * x + A[2] * y + A[4], A[1] * x + A[3] * y + A[5]];
const T = (x, y) => [1, 0, 0, 1, x, y];
const Sc = s => [s, 0, 0, s, 0, 0];

// same transform pdf.js uses for a page viewport (PDF space -> screen, y down)
function viewportMatrix(view, rotation, scale) {
  const [x0, y0, x1, y1] = view;
  const cx = (x0 + x1) / 2, cy = (y0 + y1) / 2;
  let a, b, c, d;
  switch (rotation) {
    case 90: a = 0; b = 1; c = 1; d = 0; break;
    case 180: a = -1; b = 0; c = 0; d = 1; break;
    case 270: a = 0; b = -1; c = -1; d = 0; break;
    default: a = 1; b = 0; c = 0; d = -1;
  }
  let ox, oy;
  if (a === 0) { ox = Math.abs(cy - y0) * scale; oy = Math.abs(cx - x0) * scale; }
  else { ox = Math.abs(cx - x0) * scale; oy = Math.abs(cy - y0) * scale; }
  return [a * scale, b * scale, c * scale, d * scale, ox - a * scale * cx - c * scale * cy, oy - b * scale * cx - d * scale * cy];
}

// ---------------------------------------------------------------- state
const S = {
  path: null, name: 'Untitled.pdf', dirty: false, open: false,
  sources: [],       // { bytes, doc, name, pages: [PDFPageProxy], password }
  entries: [],       // pages in document order: { id, src, index, rot, blank: {w,h} | null, objects: [] }
  form: {},          // form field values the user changed, by field name
  flattenForms: false,
  zoom: 1, fit: 'width',
  tool: 'select',
  style: { color: '#e8231a', hl: '#ffe14d', width: 2, opacity: 1, font: 'Helvetica', size: 12, bold: false, italic: false, textColor: '#000000', align: 'left' },
  sel: null,         // { eid, oid }
  editing: null,     // object id of the text box being typed in
  selPages: new Set(),
  current: 0,
  history: [], hIndex: -1,
  clip: null,
  signatures: [],
  untitled: 1,
};
const views = new Map();          // entry id -> page view
const textCache = new Map();      // "src:index" -> Promise<textContent>
const annotCache = new Map();     // "src:index" -> Promise<annotations>
const pagesEl = $('#pdf-pages');
const scroller = $('#pdf-scroller');
const uid = () => Math.random().toString(36).slice(2, 10);
const scale = () => S.zoom * PT;
const entryById = id => S.entries.find(e => e.id === id);
const pageProxy = ent => S.sources[ent.src].pages[ent.index];
const keyOf = ent => `${ent.src}:${ent.index}`;

// Page geometry. `crop` (margins in PDF units of the unrotated page) shrinks the visible
// box; V1 maps PDF space onto the cropped page at 100 % and `shift(s)` is the extra
// translation needed when pdf.js draws the full page.
function geom(ent) {
  let full, baseRot;
  if (ent.blank) { full = [0, 0, ent.blank.w, ent.blank.h]; baseRot = 0; }
  else { const p = pageProxy(ent); full = p.view; baseRot = p.rotate; }
  const c = ent.crop;
  const view = c ? [full[0] + c.l, full[1] + c.b, full[2] - c.r, full[3] - c.t] : full;
  const rotation = (((baseRot + ent.rot) % 360) + 360) % 360;
  const pw = view[2] - view[0], ph = view[3] - view[1];
  const swap = rotation % 180 !== 0;
  const V1 = viewportMatrix(view, rotation, 1);
  const fullV1 = viewportMatrix(full, rotation, 1);
  const d = mul(V1, inv(fullV1)); // pure translation
  return { view, full, rotation, baseRot, V1, w1: swap ? ph : pw, h1: swap ? pw : ph, cropped: !!c, shift: s => [1, 0, 0, 1, d[4] * s, d[5] * s] };
}
function textContent(ent) {
  const k = keyOf(ent);
  if (!textCache.has(k)) textCache.set(k, pageProxy(ent).getTextContent({ includeMarkedContent: true }));
  return textCache.get(k);
}
function annotations(ent) {
  const k = keyOf(ent);
  if (!annotCache.has(k)) annotCache.set(k, pageProxy(ent).getAnnotations({ intent: 'display' }).catch(() => []));
  return annotCache.get(k);
}

// ---------------------------------------------------------------- loading
async function loadPdfjs(bytes, name) {
  let password = null;
  const task = pdfjsLib.getDocument({ data: bytes.slice(), ...PDF_OPTS });
  task.onPassword = async (update, reason) => {
    const wrong = reason === pdfjsLib.PasswordResponses.INCORRECT_PASSWORD;
    const r = await askForm(wrong ? 'Wrong password' : 'Password required', [{ id: 'pw', label: `"${name}" is protected. Enter its password to open it.`, type: 'password' }], 'Open', wrong ? 'That password is not correct. Please try again.' : null);
    if (!r) { task.destroy(); return; }
    password = r.pw;
    update(r.pw);
  };
  const doc = await task.promise;
  const pages = [];
  for (let i = 1; i <= doc.numPages; i++) pages.push(await doc.getPage(i));
  return { doc, pages, password };
}

async function openBytes(bytes, name, path) {
  setBusy(`Opening ${name}…`);
  try {
    const first = await loadPdfjs(bytes, name);
    const previous = S.sources.slice();
    let restored = null;
    try { restored = await readEditLayer(first.doc); } catch (err) { console.warn('edit layer ignored', err); }
    textCache.clear();
    annotCache.clear();
    if (restored) {
      // reopen the document exactly as it was being edited
      S.sources = [];
      for (const src of restored.sources) {
        const b = Uint8Array.fromBase64(src.b64);
        S.sources.push({ bytes: b, name: src.name, ...(await loadPdfjs(b, src.name)) });
      }
      S.entries = restored.entries;
      S.form = restored.form || {};
      S.flattenForms = !!restored.flatten;
      S.protect = first.password ? { user: first.password, owner: first.password, allowPrint: true, allowCopy: true, allowModify: true } : null;
      first.doc.destroy();
    } else {
      const plain = await decryptWithPassword(bytes, first);
      S.sources = [{ bytes: plain, name, ...first }];
      S.entries = first.pages.map((p, i) => ({ id: uid(), src: 0, index: i, rot: 0, blank: null, objects: [] }));
      S.form = {};
      S.flattenForms = false;
      // keep a password-protected PDF protected when it is saved again
      S.protect = first.password ? { user: first.password, owner: first.password, allowPrint: true, allowCopy: true, allowModify: true } : null;
    }
    S.sanitize = false;
    S.path = path;
    S.name = name;
    S.open = true;
    S.dirty = false;
    S.sel = null;
    S.selPages.clear();
    S.current = 0;
    $('#empty-state').hidden = true;
    S.fit = 'width';
    buildPages();
    scroller.scrollTop = 0;
    fitWidth();
    updateStatus();
    resetHistory();
    // free the document this window showed before
    for (const src of previous) if (src.doc && !S.sources.includes(src)) src.doc.destroy().catch(() => {});
    updateTitle();
    updateToggles();
    for (const ent of S.entries) if (ent.ocr) injectOcrText(ent);
    if (restored) toast('Your earlier edits are still editable.');
    else if (first.password) toast('This PDF is password-protected. It stays protected when you save; use Protect › Remove Password to remove it.');
  } finally {
    setBusy(null);
  }
}
async function readEditLayer(doc) {
  const atts = await doc.getAttachments();
  if (!atts) return null;
  const a = Object.values(atts).find(x => x && x.filename === EDIT_FILE);
  if (!a) return null;
  const data = JSON.parse(new TextDecoder().decode(a.content));
  if (!data || data.v !== 1 || !Array.isArray(data.sources) || !Array.isArray(data.entries)) return null;
  // only trust it if no other program has saved the file since Nisar Office Studio did
  const meta = await doc.getMetadata();
  const mod = pdfjsLib.PDFDateString.toDateObject(meta && meta.info && meta.info.ModDate);
  if (!mod || Math.abs(mod.getTime() - Date.parse(data.mod)) > 2500) return null;
  return data;
}

async function newBlankPdf() {
  const doc = await PL.PDFDocument.create();
  const a4 = !/^en-(US|CA)|^es-(MX|US)/.test(navigator.language);
  doc.addPage(a4 ? [595.28, 841.89] : [612, 792]);
  doc.setProducer('Nisar Office Studio');
  const bytes = await doc.save();
  await openBytes(bytes, S.untitled > 1 ? `Untitled ${S.untitled}.pdf` : 'Untitled.pdf', null);
  S.dirty = false;
  updateTitle();
}

// ---------------------------------------------------------------- page views
const io = new IntersectionObserver(items => {
  for (const it of items) {
    const v = views.get(it.target.dataset.eid);
    if (!v) continue;
    v.visible = it.isIntersecting;
    if (v.visible) renderView(v); else releaseView(v);
  }
}, { root: scroller, rootMargin: '1200px 0px' });

function buildPages() {
  io.disconnect();
  for (const v of views.values()) if (v.task) v.task.cancel();
  views.clear();
  pagesEl.innerHTML = '';
  for (const ent of S.entries) {
    const el = document.createElement('div');
    el.className = 'pg';
    el.dataset.eid = ent.id;
    el.innerHTML = '<div class="layer search-layer"></div><div class="layer field-layer"></div><div class="layer obj-layer"></div>';
    const v = { ent, el, canvas: null, text: null, textLayer: null, done: 0, token: 0, task: null, visible: false };
    views.set(ent.id, v);
    sizeView(v);
    pagesEl.appendChild(el);
    renderObjects(v);
    io.observe(el);
  }
  renderThumbs();
  updateStatus();
}
function sizeView(v) {
  const g = geom(v.ent);
  const s = scale();
  v.el.style.width = `${g.w1 * s}px`;
  v.el.style.height = `${g.h1 * s}px`;
  v.el.style.setProperty('--total-scale-factor', s);
}
function releaseView(v) {
  if (v.task) { v.task.cancel(); v.task = null; }
  v.token++;
  if (v.canvas) { v.canvas.width = 0; v.canvas.remove(); v.canvas = null; }
  if (v.text) { v.text.remove(); v.text = null; v.textLayer = null; }
  v.done = 0;
}
async function renderView(v) {
  const s = scale();
  if (v.done === s) return;
  const token = ++v.token;
  if (v.task) { v.task.cancel(); v.task = null; }
  const ent = v.ent;
  const g = geom(ent);
  const cssW = g.w1 * s, cssH = g.h1 * s;
  let ratio = window.devicePixelRatio || 1;
  const maxPx = 16e6;
  if (cssW * cssH * ratio * ratio > maxPx) ratio = Math.sqrt(maxPx / (cssW * cssH));
  const canvas = document.createElement('canvas');
  canvas.className = 'render';
  canvas.width = Math.max(1, Math.floor(cssW * ratio));
  canvas.height = Math.max(1, Math.floor(cssH * ratio));
  const ctx = canvas.getContext('2d', { alpha: false });
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  if (!ent.blank) {
    const page = pageProxy(ent);
    const task = page.render({
      canvasContext: ctx,
      viewport: page.getViewport({ scale: s * ratio, rotation: g.rotation }),
      transform: g.cropped ? g.shift(s * ratio) : null,
      annotationMode: ent.src === 0 ? pdfjsLib.AnnotationMode.ENABLE_FORMS : pdfjsLib.AnnotationMode.ENABLE,
    });
    v.task = task;
    try { await task.promise; } catch (err) { if (token === v.token && !(err && err.name === 'RenderingCancelledException')) console.error(err); return; }
    if (token !== v.token) return;
    v.task = null;
  }
  if (v.canvas) { v.canvas.width = 0; v.canvas.replaceWith(canvas); } else v.el.prepend(canvas);
  v.canvas = canvas;
  v.done = s;
  if (!ent.blank) {
    try {
      await buildTextLayer(v, s, g, token);
      if (token !== v.token) return;
      await buildFieldLayer(v, s, g);
    } catch (err) { console.error(err); }
  }
  if (token !== v.token) return;
  renderObjects(v);
  if (find.query) paintFind(v);
}
async function buildTextLayer(v, s, g, token) {
  const tc = await textContent(v.ent);
  if (token !== v.token) return;
  const div = document.createElement('div');
  div.className = 'textLayer';
  const layer = new pdfjsLib.TextLayer({ textContentSource: tc, container: div, viewport: pageProxy(v.ent).getViewport({ scale: s, rotation: g.rotation }) });
  await layer.render();
  if (token !== v.token) return;
  if (g.cropped) { const t = g.shift(s); div.style.left = `${t[4]}px`; div.style.top = `${t[5]}px`; div.style.right = 'auto'; div.style.bottom = 'auto'; }
  v.el.classList.toggle('cropped', g.cropped);
  const end = document.createElement('div');
  end.className = 'endOfContent';
  div.appendChild(end);
  if (v.text) v.text.replaceWith(div); else v.el.querySelector('.search-layer').before(div);
  v.text = div;
  v.textLayer = layer;
}

// links, comments and fillable form fields that are part of the PDF
async function buildFieldLayer(v, s, g) {
  const layer = v.el.querySelector('.field-layer');
  const annots = await annotations(v.ent);
  layer.innerHTML = '';
  const Vs = mul(Sc(s), g.V1);
  for (const a of annots) {
    if (!a.rect) continue;
    const pts = [apply(Vs, a.rect[0], a.rect[1]), apply(Vs, a.rect[2], a.rect[3])];
    const box = { x: Math.min(pts[0][0], pts[1][0]), y: Math.min(pts[0][1], pts[1][1]) };
    box.w = Math.abs(pts[0][0] - pts[1][0]);
    box.h = Math.abs(pts[0][1] - pts[1][1]);
    const place = el => { Object.assign(el.style, { left: `${box.x}px`, top: `${box.y}px`, width: `${box.w}px`, height: `${box.h}px` }); layer.appendChild(el); return el; };
    if (a.subtype === 'Link' && (a.url || a.dest || a.unsafeUrl)) {
      const el = document.createElement('a');
      el.className = 'lnk';
      const url = a.url || a.unsafeUrl;
      el.title = url ? `${url}\nClick to open` : 'Go to the linked page';
      el.addEventListener('click', ev => { ev.preventDefault(); if (url) api.openExternal(url); else goToDest(a.dest); });
      place(el);
    } else if (a.subtype === 'Widget' && v.ent.src === 0 && a.fieldName && !a.hidden) {
      const el = fieldInput(a, box, s);
      if (el) place(el);
    } else if (a.subtype === 'Text' && a.contentsObj && a.contentsObj.str) {
      const el = document.createElement('div');
      el.className = 'cmt';
      el.title = `${a.titleObj && a.titleObj.str ? `${a.titleObj.str}: ` : ''}${a.contentsObj.str}`;
      Object.assign(el.style, { left: `${box.x}px`, top: `${box.y}px` });
      layer.appendChild(el);
    }
  }
}
function fieldValue(a) {
  if (a.fieldName in S.form) return S.form[a.fieldName];
  if (a.fieldType === 'Btn' && a.checkBox) return !!a.fieldValue && a.fieldValue !== 'Off';
  return a.fieldValue;
}
function fieldInput(a, box, s) {
  const name = a.fieldName;
  let el;
  const fontPt = a.defaultAppearanceData && a.defaultAppearanceData.fontSize ? a.defaultAppearanceData.fontSize : 0;
  const fontPx = fontPt ? fontPt * s : Math.max(8, Math.min(box.h * 0.68, 14 * s));
  const val = fieldValue(a);
  if (a.fieldType === 'Tx') {
    el = document.createElement(a.multiLine ? 'textarea' : 'input');
    if (!a.multiLine) el.type = a.password ? 'password' : 'text';
    el.value = val == null ? '' : String(val);
    if (a.maxLen) el.maxLength = a.maxLen;
    el.style.fontSize = `${fontPx}px`;
    el.style.textAlign = ['left', 'center', 'right'][a.textAlignment] || 'left';
    el.spellcheck = false;
    el.addEventListener('input', () => setField(name, el.value, false, el));
    el.addEventListener('change', () => setField(name, el.value, true, el));
  } else if (a.fieldType === 'Btn' && a.checkBox) {
    el = document.createElement('input');
    el.type = 'checkbox';
    el.checked = !!val;
    el.style.setProperty('--ck', `${Math.max(8, box.h * 0.8)}px`);
    el.addEventListener('change', () => setField(name, el.checked, true, el));
  } else if (a.fieldType === 'Btn' && a.radioButton) {
    el = document.createElement('input');
    el.type = 'radio';
    el.name = `radio:${name}`;
    el.value = a.buttonValue;
    el.checked = val === a.buttonValue;
    el.addEventListener('change', () => { if (el.checked) setField(name, a.buttonValue, true, el); });
  } else if (a.fieldType === 'Ch') {
    el = document.createElement('select');
    if (!a.combo) { el.size = Math.max(2, Math.floor(box.h / (fontPx * 1.3))); el.multiple = !!a.multiSelect; }
    for (const o of a.options || []) {
      const opt = document.createElement('option');
      opt.value = o.exportValue;
      opt.textContent = o.displayValue;
      opt.selected = Array.isArray(val) ? val.includes(o.exportValue) : val === o.exportValue || val === o.displayValue;
      el.appendChild(opt);
    }
    el.style.fontSize = `${fontPx}px`;
    el.addEventListener('change', () => setField(name, el.multiple ? [...el.selectedOptions].map(o => o.value) : el.value, true, el));
  } else return null;
  el.className = 'fld';
  el.dataset.name = name;
  el.disabled = !!a.readOnly;
  el.title = a.alternativeText || name;
  return el;
}
function setField(name, value, commitNow, from) {
  S.form[name] = value;
  // keep other widgets of the same field (on other pages) in step
  for (const el of $$('.fld')) {
    if (el === from || el.dataset.name !== name) continue;
    if (el.type === 'checkbox') el.checked = !!value;
    else if (el.type === 'radio') el.checked = el.value === value;
    else if (el.tagName !== 'SELECT') el.value = value;
  }
  if (commitNow) commit(); else markDirty();
}

// ---------------------------------------------------------------- objects on pages
function renderObjects(v) {
  const layer = v.el.querySelector('.obj-layer');
  layer.innerHTML = '';
  const s = scale();
  const Vs = mul(Sc(s), geom(v.ent).V1);
  for (const o of v.ent.objects) layer.appendChild(objElement(o, Vs, s));
  if (S.sel && S.sel.eid === v.ent.id) markSelection();
  syncTextSizes(v);
}
function colorAlpha(hex, a) {
  const n = parseInt(hex.slice(1), 16);
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${a})`;
}
function objElement(o, Vs, s) {
  const el = document.createElement('div');
  el.className = `obj obj-${o.type}`;
  el.dataset.oid = o.id;
  el.style.width = `${o.w}px`;
  el.style.height = `${o.h}px`;
  el.style.transform = `matrix(${mul(Vs, o.m).map(n => Math.round(n * 1e5) / 1e5).join(',')})`;
  el.style.setProperty('--hs', `${10 / s}px`);
  el.style.setProperty('--hw', `${1.4 / s}px`);
  const sw = o.width || 2;
  const W = o.w, H = o.h;
  const svgWrap = inner => `<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none">${inner}</svg>`;
  switch (o.type) {
    case 'text': {
      const t = document.createElement('div');
      t.className = 'txt';
      t.dir = 'auto';
      t.textContent = o.text;
      t.style.font = `${o.italic ? 'italic ' : ''}${o.bold ? '700 ' : '400 '}${o.size}px/1.2 ${CSS_FONT[o.font] || CSS_FONT.Helvetica}`;
      t.style.color = o.color;
      t.style.textAlign = o.align || 'left';
      if (o.opacity != null && o.opacity < 1) t.style.opacity = o.opacity;
      if (o.autoW) { el.style.width = 'max-content'; t.style.whiteSpace = 'pre'; }
      el.style.height = 'auto';
      el.appendChild(t);
      if (!o.text) el.classList.add('empty');
      break;
    }
    case 'image': el.innerHTML = `<img src="${o.src}" alt="" draggable="false">`; break;
    case 'rect': el.innerHTML = svgWrap(`<rect x="${sw / 2}" y="${sw / 2}" width="${Math.max(0, W - sw)}" height="${Math.max(0, H - sw)}" fill="${o.fill || 'none'}" stroke="${o.color}" stroke-width="${sw}" opacity="${o.opacity}"/>`); break;
    case 'ellipse': el.innerHTML = svgWrap(`<ellipse cx="${W / 2}" cy="${H / 2}" rx="${Math.max(0, (W - sw) / 2)}" ry="${Math.max(0, (H - sw) / 2)}" fill="${o.fill || 'none'}" stroke="${o.color}" stroke-width="${sw}" opacity="${o.opacity}"/>`); break;
    case 'line':
    case 'arrow': {
      const [[x1, y1], [x2, y2]] = o.pts.map(([x, y]) => [x * W, y * H]);
      let head = '';
      if (o.type === 'arrow') head = `<polygon points="${arrowHead(x1, y1, x2, y2, sw).map(p => p.join(',')).join(' ')}" fill="${o.color}"/>`;
      el.innerHTML = svgWrap(`<g opacity="${o.opacity}"><line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${o.color}" stroke-width="${sw}" stroke-linecap="round"/>${head}</g>`);
      break;
    }
    case 'ink': el.innerHTML = svgWrap(`<path d="${inkPath(o)}" fill="none" stroke="${o.color}" stroke-width="${sw}" stroke-linecap="round" stroke-linejoin="round" opacity="${o.opacity}"/>`); break;
    // multiply blending can't reach the page canvas through the object layer, so show it translucent
    case 'highlight': el.style.background = colorAlpha(o.color, 0.45); break;
    case 'underline':
    case 'strike': {
      const line = document.createElement('div');
      const th = Math.max(0.8, H * 0.075);
      Object.assign(line.style, { position: 'absolute', left: 0, right: 0, height: `${th}px`, background: o.color, [o.type === 'underline' ? 'bottom' : 'top']: o.type === 'underline' ? `${H * 0.04}px` : `${H * 0.52 - th / 2}px` });
      el.appendChild(line);
      break;
    }
    case 'whiteout': case 'redact': break;
    case 'note':
      el.style.background = o.color;
      el.title = o.text || 'Empty comment';
      el.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="#3b3000" stroke-width="2" stroke-linejoin="round"><path d="M4 4h16v12H9l-5 4z"/></svg>';
      break;
    case 'field':
      el.innerHTML = `<span>${esc(o.name || 'Field')}</span><span class="fk">${{ text: 'Text', check: '☐', dropdown: '▾' }[o.kind] || ''}</span>`;
      el.style.fontSize = `${Math.max(7, Math.min(12, H * 0.5))}px`;
      break;
    case 'stamp':
      el.style.borderColor = o.color;
      el.style.borderWidth = `${Math.max(1.5, H * 0.07)}px`;
      el.style.color = o.color;
      el.style.fontSize = `${H * 0.46}px`;
      el.textContent = o.text;
      break;
  }
  return el;
}
function arrowHead(x1, y1, x2, y2, sw) {
  const ang = Math.atan2(y2 - y1, x2 - x1);
  const len = Math.max(8, sw * 4.5), spread = Math.PI / 7;
  return [[x2, y2], [x2 - len * Math.cos(ang - spread), y2 - len * Math.sin(ang - spread)], [x2 - len * Math.cos(ang + spread), y2 - len * Math.sin(ang + spread)]];
}
function inkPath(o) {
  return o.strokes.map(st => st.map(([x, y], i) => `${i ? 'L' : 'M'}${(x * o.w).toFixed(2)} ${(y * o.h).toFixed(2)}`).join(' ')).join(' ');
}
// text boxes size themselves to their text; keep the stored box in step
function syncTextSizes(v) {
  for (const o of v.ent.objects) {
    if (o.type !== 'text') continue;
    const el = v.el.querySelector(`.obj[data-oid="${o.id}"]`);
    if (!el) continue;
    const h = el.offsetHeight, w = el.offsetWidth;
    if (h && Math.abs(h - o.h) > 0.5) o.h = h;
    if (o.autoW && w && Math.abs(w - o.w) > 0.5) o.w = w;
  }
}
function viewOfObject(oid) {
  for (const v of views.values()) if (v.ent.objects.some(o => o.id === oid)) return v;
  return null;
}
function objById(oid) {
  for (const ent of S.entries) { const o = ent.objects.find(x => x.id === oid); if (o) return { o, ent }; }
  return null;
}
function objectEl(oid) { return pagesEl.querySelector(`.obj[data-oid="${oid}"]`); }

function defaultsFor(type) {
  const st = S.style;
  switch (type) {
    case 'text': return { text: '', font: st.font, size: st.size, bold: st.bold, italic: st.italic, color: st.textColor, align: st.align, autoW: true };
    case 'highlight': return { color: st.hl };
    case 'underline': case 'strike': return { color: st.color === '#e8231a' ? '#e8231a' : st.color };
    case 'note': return { text: '', color: '#ffd34d' };
    case 'whiteout': case 'redact': return {};
    default: return { color: st.color, width: st.width, opacity: st.opacity };
  }
}
function createObject(v, type, x, y, w, h, props = {}) {
  const g = geom(v.ent);
  const o = { id: uid(), type, m: mul(inv(g.V1), T(x, y)), w: Math.max(1, w), h: Math.max(1, h), ...defaultsFor(type), ...props };
  v.ent.objects.push(o);
  renderObjects(v);
  select(v.ent.id, o.id);
  return o;
}
function deleteObject(oid) {
  const hit = objById(oid);
  if (!hit) return;
  hit.ent.objects = hit.ent.objects.filter(o => o.id !== oid);
  if (S.sel && S.sel.oid === oid) select(null);
  const v = views.get(hit.ent.id);
  if (v) renderObjects(v);
}

// ---- selection & the floating object toolbar
function select(eid, oid) {
  if (S.editing && (!oid || S.editing !== oid)) finishEditing();
  S.sel = eid && oid ? { eid, oid } : null;
  markSelection();
  showObjbar();
  closeNote(true);
  updateFontControls();
}
function markSelection() {
  for (const el of $$('.obj.selected', pagesEl)) { el.classList.remove('selected'); for (const h of $$('.hd', el)) h.remove(); }
  if (!S.sel) return;
  const el = objectEl(S.sel.oid);
  const hit = objById(S.sel.oid);
  if (!el || !hit) return;
  el.classList.add('selected');
  const t = hit.o.type;
  const handles = t === 'note' ? [] : t === 'text' ? ['w', 'e'] : ['nw', 'ne', 'sw', 'se'];
  for (const h of handles) { const d = document.createElement('span'); d.className = 'hd'; d.dataset.h = h; el.appendChild(d); }
}
function showObjbar() {
  const bar = $('#objbar');
  const hit = S.sel && objById(S.sel.oid);
  if (!hit) { bar.hidden = true; return; }
  const o = hit.o;
  const parts = [];
  if (o.type === 'text') {
    parts.push(`<select data-ob="font">${Object.keys(CSS_FONT).map(f => `<option value="${f}"${o.font === f ? ' selected' : ''}>${f === 'Times' ? 'Times New Roman' : f === 'Courier' ? 'Courier New' : 'Helvetica'}</option>`).join('')}</select>`);
    parts.push(`<input data-ob="size" value="${o.size}" title="Font size">`);
    parts.push(`<button class="ico${o.bold ? ' on' : ''}" data-ob="bold" title="Bold"><b class="glyph">B</b></button><button class="ico${o.italic ? ' on' : ''}" data-ob="italic" title="Italic"><i class="glyph serif">I</i></button>`);
    parts.push(`<button class="ico" data-menu="objColor" title="Text colour"><span class="swatch" style="background:${o.color}"></span></button>`);
    parts.push(`<button class="ico" data-ob="edit" title="Edit text" >${svg('editText')}</button>`);
  } else if (['rect', 'ellipse', 'line', 'arrow', 'ink'].includes(o.type)) {
    parts.push(`<button class="ico" data-menu="objColor" title="Colour"><span class="swatch" style="background:${o.color}"></span></button>`);
    if (o.type === 'rect' || o.type === 'ellipse') parts.push(`<button class="btn" data-menu="objFill" title="Fill colour">Fill</button>`);
    parts.push(`<button class="btn" data-menu="objWidth" title="Line thickness">${o.width} pt</button>`);
    parts.push(`<button class="btn" data-menu="objOpacity" title="Opacity">${Math.round(o.opacity * 100)}%</button>`);
  } else if (MARKUP.has(o.type) || o.type === 'stamp') {
    parts.push(`<button class="ico" data-menu="objColor" title="Colour"><span class="swatch" style="background:${o.color}"></span></button>`);
  } else if (o.type === 'field') {
    parts.push(`<input data-ob="fname" value="${esc(o.name || '')}" title="Field name" style="width:120px">`);
    if (o.kind === 'dropdown') parts.push('<button class="btn" data-ob="fopts">Options…</button>');
    if (o.kind === 'text') parts.push(`<button class="btn${o.multiline ? ' on' : ''}" data-ob="fmulti" title="Allow several lines">Multi-line</button>`);
  } else if (o.type === 'note') {
    parts.push('<button class="btn" data-ob="note">Edit comment</button>');
    parts.push(`<button class="ico" data-menu="objColor" title="Colour"><span class="swatch" style="background:${o.color}"></span></button>`);
  }
  parts.push(`<span class="sep"></span><button class="ico" data-ob="duplicate" title="Duplicate (Ctrl+D)">${svg('copy')}</button><button class="ico" data-ob="delete" title="Delete (Del)">${svg('trash')}</button>`);
  bar.innerHTML = parts.join('');
  bar.hidden = false;
  positionObjbar();
}
function positionObjbar() {
  const bar = $('#objbar');
  if (bar.hidden || !S.sel) return;
  const el = objectEl(S.sel.oid);
  if (!el) { bar.hidden = true; return; }
  const ws = $('#pdf-workspace').getBoundingClientRect();
  const r = el.getBoundingClientRect();
  let top = r.top - ws.top - bar.offsetHeight - 10;
  if (top < 6) top = r.bottom - ws.top + 10;
  bar.style.top = `${clamp(top, 6, ws.height - bar.offsetHeight - 6)}px`;
  bar.style.left = `${clamp(r.left - ws.left, 6, ws.width - bar.offsetWidth - 6)}px`;
}
function updateSelected(changes) {
  const hit = S.sel && objById(S.sel.oid);
  if (!hit) return false;
  Object.assign(hit.o, changes);
  const v = views.get(hit.ent.id);
  if (v) renderObjects(v);
  showObjbar();
  commit();
  return true;
}

// ---- text editing
function startEditing(oid, selectAll = false) {
  const hit = objById(oid);
  if (!hit || hit.o.type !== 'text') return;
  const el = objectEl(oid);
  if (!el) return;
  if (S.editing !== oid) S.editStart = hit.o.text;
  S.editing = oid;
  const t = $('.txt', el);
  el.classList.add('editing');
  t.contentEditable = 'plaintext-only';
  t.spellcheck = true;
  t.focus();
  const r = document.createRange();
  r.selectNodeContents(t);
  if (!selectAll) r.collapse(false);
  const sel = getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
  t.addEventListener('input', () => {
    hit.o.text = t.innerText.replace(/\n$/, '');
    el.classList.toggle('empty', !hit.o.text);
    const v = views.get(hit.ent.id);
    if (v) syncTextSizes(v);
    positionObjbar();
    markDirty();
  });
  t.addEventListener('blur', () => setTimeout(() => { if (S.editing === oid && document.activeElement !== t) finishEditing(); }, 0));
}
function finishEditing() {
  const oid = S.editing;
  if (!oid) return;
  S.editing = null;
  const hit = objById(oid);
  const el = objectEl(oid);
  if (el) { el.classList.remove('editing'); const t = $('.txt', el); if (t) t.contentEditable = 'false'; }
  if (!hit) return;
  hit.o.text = (hit.o.text || '').replace(/\s+$/, '');
  if (!hit.o.text.trim()) { deleteObject(oid); commit(); return; }
  const v = views.get(hit.ent.id);
  if (v) renderObjects(v);
  if (hit.o.text !== S.editStart) commit();
}
function updateFontControls() {
  const hit = S.sel && objById(S.sel.oid);
  const src = hit && hit.o.type === 'text' ? hit.o : { font: S.style.font, size: S.style.size, bold: S.style.bold, italic: S.style.italic, color: S.style.textColor, align: S.style.align };
  $('#font-select').value = src.font;
  if (document.activeElement !== $('#size-input')) $('#size-input').value = src.size;
  $('[data-cmd="bold"]').classList.toggle('on', !!src.bold);
  $('[data-cmd="italic"]').classList.toggle('on', !!src.italic);
  for (const a of ['Left', 'Center', 'Right']) $(`[data-cmd="align${a}"]`).classList.toggle('on', (src.align || 'left') === a.toLowerCase());
  $('#text-color-bar').style.background = src.color;
}
// font controls change the selected text box, or the style for new text
function setTextStyle(changes) {
  const hit = S.sel && objById(S.sel.oid);
  if (hit && hit.o.type === 'text') {
    updateSelected(changes);
    if (S.editing === hit.o.id) startEditing(hit.o.id);
  } else {
    const map = { color: 'textColor' };
    for (const [k, val] of Object.entries(changes)) S.style[map[k] || k] = val;
  }
  updateFontControls();
}

// ---------------------------------------------------------------- pointer tools
const toolHints = {
  select: 'Select text, form fields and added items. Double-click added text to edit it.',
  hand: 'Drag to move around the document.',
  text: 'Click where the new text should go.',
  edittext: 'Click a line of text in the PDF to change it.',
  highlight: 'Select text to highlight it, or drag a box.',
  underline: 'Select text to underline it.',
  strike: 'Select text to strike it through.',
  note: 'Click to place a comment.',
  ink: 'Drag to draw.',
  rect: 'Drag to draw a rectangle.', ellipse: 'Drag to draw an ellipse.', line: 'Drag to draw a line.', arrow: 'Drag to draw an arrow.',
  whiteout: 'Drag over content to cover it with white.',
  redact: 'Drag over content to remove it permanently when you save.',
  fieldText: 'Drag to draw a fillable text field.',
  fieldCheck: 'Click or drag to place a check box.',
  fieldDropdown: 'Drag to draw a drop-down list.',
};
const FIELD_KIND = { fieldText: 'text', fieldCheck: 'check', fieldDropdown: 'dropdown' };
function nextFieldName(kind) {
  const base = { text: 'Text', check: 'Check', dropdown: 'Dropdown' }[kind];
  const used = new Set(S.entries.flatMap(e => e.objects.filter(o => o.type === 'field').map(o => o.name)));
  let i = 1;
  while (used.has(`${base}${i}`)) i++;
  return `${base}${i}`;
}
function setTool(tool) {
  if (!toolHints[tool]) return;
  finishEditing();
  S.tool = tool;
  document.body.dataset.tool = tool;
  for (const b of $$('[data-tool]')) b.classList.toggle('on', b.dataset.tool === tool);
  $('#st-tool').textContent = tool === 'select' ? '' : toolHints[tool];
  if (tool !== 'select') select(null);
  getSelection().removeAllRanges();
}
function viewAt(target) { const el = target.closest && target.closest('.pg'); return el ? views.get(el.dataset.eid) : null; }
function pointIn(v, ev) {
  const r = v.el.getBoundingClientRect();
  const k = r.width / geom(v.ent).w1;
  return { x: (ev.clientX - r.left) / k, y: (ev.clientY - r.top) / k, k };
}
function dragLoop(ev, onMove, onUp) {
  const target = ev.target;
  try { target.setPointerCapture(ev.pointerId); } catch { /* ignore */ }
  const move = e => onMove(e);
  const up = e => {
    target.removeEventListener('pointermove', move);
    target.removeEventListener('pointerup', up);
    target.removeEventListener('pointercancel', up);
    onUp(e);
  };
  target.addEventListener('pointermove', move);
  target.addEventListener('pointerup', up);
  target.addEventListener('pointercancel', up);
}

pagesEl.addEventListener('pointerdown', ev => {
  if (ev.button !== 0 || !S.open) return;
  const v = viewAt(ev.target);
  const tool = S.tool;
  if (tool === 'hand') { startPan(ev); return; }
  if (!v) { if (tool === 'select') select(null); return; }
  const objEl = ev.target.closest('.obj');
  if (tool === 'select') {
    if (objEl) {
      const hit = objById(objEl.dataset.oid);
      if (!hit) return;
      if (S.editing === hit.o.id && ev.target.closest('.txt')) return; // typing: let the caret move
      ev.preventDefault();
      const h = ev.target.closest('.hd');
      if (h) { startResize(v, hit.o, h.dataset.h, ev); return; }
      select(v.ent.id, hit.o.id);
      if (hit.o.type === 'text' && ev.detail >= 2) { startEditing(hit.o.id); return; }
      if (hit.o.type === 'note' && ev.detail >= 2) { openNote(hit.o); return; }
      startMove(v, hit.o, ev);
      return;
    }
    if (ev.target.closest('.fld, .lnk, .cmt')) return;
    select(null);
    const tl = ev.target.closest('.textLayer');
    if (tl) tl.classList.add('selecting');
    return;
  }
  const p = pointIn(v, ev);
  switch (tool) {
    case 'text': {
      ev.preventDefault();
      const o = createObject(v, 'text', p.x - TEXT_PAD, p.y - (S.style.size * 1.2) / 2 - TEXT_PAD, 10, S.style.size * 1.2 + 2 * TEXT_PAD);
      startEditing(o.id);
      return;
    }
    case 'note': {
      ev.preventDefault();
      const o = createObject(v, 'note', p.x - 11, p.y - 11, 22, 22);
      commit();
      setTool('select');
      select(v.ent.id, o.id);
      openNote(o);
      return;
    }
    case 'edittext': {
      const span = ev.target.closest('.textLayer span');
      if (span && !span.classList.contains('endOfContent') && span.textContent.trim()) { ev.preventDefault(); editTextSpan(v, span); }
      return;
    }
    case 'highlight': case 'underline': case 'strike':
      if (ev.target.closest('.textLayer span:not(.endOfContent)')) { ev.target.closest('.textLayer').classList.add('selecting'); return; }
      ev.preventDefault();
      startBox(v, p, ev, tool);
      return;
    case 'ink': ev.preventDefault(); startInk(v, p, ev); return;
    case 'rect': case 'ellipse': case 'whiteout': case 'redact': case 'fieldText': case 'fieldCheck': case 'fieldDropdown': ev.preventDefault(); startBox(v, p, ev, tool); return;
    case 'line': case 'arrow': ev.preventDefault(); startLine(v, p, ev, tool); return;
  }
});
document.addEventListener('pointerup', () => {
  for (const tl of $$('.textLayer.selecting')) tl.classList.remove('selecting');
  if (MARKUP.has(S.tool)) setTimeout(() => markupFromSelection(S.tool), 0);
});

function startPan(ev) {
  ev.preventDefault();
  const sx = ev.clientX, sy = ev.clientY, sl = scroller.scrollLeft, st = scroller.scrollTop;
  document.body.classList.add('panning');
  dragLoop(ev, e => { scroller.scrollLeft = sl - (e.clientX - sx); scroller.scrollTop = st - (e.clientY - sy); }, () => document.body.classList.remove('panning'));
}
function localDelta(v, o, dxScreen, dyScreen, k) {
  // screen movement -> movement in the object's own (local) axes
  const L = mul(geom(v.ent).V1, o.m);
  const Li = inv([L[0], L[1], L[2], L[3], 0, 0]);
  return apply(Li, dxScreen / k, dyScreen / k);
}
function startMove(v, o, ev) {
  const start = { x: ev.clientX, y: ev.clientY, m: o.m };
  const k = pointIn(v, ev).k;
  let moved = false;
  dragLoop(ev, e => {
    const [lx, ly] = localDelta(v, { m: start.m }, e.clientX - start.x, e.clientY - start.y, k);
    if (!moved && Math.hypot(e.clientX - start.x, e.clientY - start.y) < 3) return;
    moved = true;
    o.m = mul(start.m, T(lx, ly));
    const el = objectEl(o.id);
    if (el) el.style.transform = `matrix(${mul(mul(Sc(scale()), geom(v.ent).V1), o.m).join(',')})`;
    positionObjbar();
  }, () => { if (moved) commit(); });
}
function startResize(v, o, handle, ev) {
  const start = { x: ev.clientX, y: ev.clientY, m: o.m, w: o.w, h: o.h, pts: o.pts };
  const k = pointIn(v, ev).k;
  const keepAspect = o.type === 'image' || o.type === 'stamp';
  if (o.type === 'text') o.autoW = false;
  dragLoop(ev, e => {
    let [dx, dy] = localDelta(v, { m: start.m }, e.clientX - start.x, e.clientY - start.y, k);
    let w = start.w, h = start.h, ox = 0, oy = 0;
    if (handle.includes('e')) w = start.w + dx;
    if (handle.includes('w')) { w = start.w - dx; ox = dx; }
    if (handle.includes('s')) h = start.h + dy;
    if (handle.includes('n')) { h = start.h - dy; oy = dy; }
    if (keepAspect && handle.length === 2) {
      const ratio = start.w / start.h;
      if (Math.abs(w / start.w) > Math.abs(h / start.h)) { const nh = w / ratio; if (handle.includes('n')) oy += h - nh; h = nh; }
      else { const nw = h * ratio; if (handle.includes('w')) ox += w - nw; w = nw; }
    }
    const min = 4;
    if (w < min) { if (handle.includes('w')) ox -= min - w; w = min; }
    if (h < min && o.type !== 'text') { if (handle.includes('n')) oy -= min - h; h = min; }
    o.m = mul(start.m, T(ox, oy));
    o.w = w;
    if (o.type !== 'text') o.h = h;
    renderObjects(v);
    positionObjbar();
  }, () => commit());
}
function dragBoxEl(v, cls) {
  const d = document.createElement('div');
  d.className = `drag-box ${cls || ''}`;
  v.el.appendChild(d);
  return d;
}
function startBox(v, p0, ev, type) {
  const box = dragBoxEl(v, type);
  let p1 = p0;
  const draw = () => {
    const x = Math.min(p0.x, p1.x) * p0.k, y = Math.min(p0.y, p1.y) * p0.k;
    Object.assign(box.style, { left: `${x}px`, top: `${y}px`, width: `${Math.abs(p1.x - p0.x) * p0.k}px`, height: `${Math.abs(p1.y - p0.y) * p0.k}px` });
  };
  draw();
  dragLoop(ev, e => { p1 = pointIn(v, e); draw(); }, () => {
    box.remove();
    let x = Math.min(p0.x, p1.x), y = Math.min(p0.y, p1.y), w = Math.abs(p1.x - p0.x), h = Math.abs(p1.y - p0.y);
    const kind = FIELD_KIND[type];
    if (w < 3 && h < 3) {
      if (type !== 'rect' && type !== 'ellipse' && !kind) return;
      // a click makes a default-sized shape or field
      [w, h] = kind === 'check' ? [14, 14] : kind ? [180, 22] : [120, 80];
      x = p0.x - (kind ? 0 : w / 2); y = p0.y - h / 2;
    }
    if (w < 2 || h < 2) return;
    if (kind) {
      if (kind === 'check') { const s = Math.min(w, h); w = h = Math.max(10, s); }
      createObject(v, 'field', x, y, w, h, { kind, name: nextFieldName(kind), options: kind === 'dropdown' ? ['Option 1', 'Option 2', 'Option 3'] : undefined });
      commit();
      return;
    }
    createObject(v, type, x, y, w, h);
    commit();
  });
}
function startLine(v, p0, ev, type) {
  const g = geom(v.ent);
  const pre = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  pre.setAttribute('class', 'layer');
  pre.style.cssText = 'pointer-events:none;z-index:9;overflow:visible';
  pre.setAttribute('viewBox', `0 0 ${g.w1} ${g.h1}`);
  v.el.appendChild(pre);
  let p1 = p0;
  const draw = () => { pre.innerHTML = `<line x1="${p0.x}" y1="${p0.y}" x2="${p1.x}" y2="${p1.y}" stroke="${S.style.color}" stroke-width="${S.style.width}" stroke-linecap="round" opacity="${S.style.opacity}"/>`; };
  dragLoop(ev, e => {
    p1 = pointIn(v, e);
    if (e.shiftKey) { // snap to 45°
      const ang = Math.round(Math.atan2(p1.y - p0.y, p1.x - p0.x) / (Math.PI / 4)) * (Math.PI / 4);
      const len = Math.hypot(p1.x - p0.x, p1.y - p0.y);
      p1 = { ...p1, x: p0.x + len * Math.cos(ang), y: p0.y + len * Math.sin(ang) };
    }
    draw();
  }, () => {
    pre.remove();
    if (Math.hypot(p1.x - p0.x, p1.y - p0.y) < 4) return;
    const pad = Math.max(2, S.style.width);
    const x = Math.min(p0.x, p1.x) - pad, y = Math.min(p0.y, p1.y) - pad;
    const w = Math.abs(p1.x - p0.x) + 2 * pad, h = Math.abs(p1.y - p0.y) + 2 * pad;
    const pts = [[(p0.x - x) / w, (p0.y - y) / h], [(p1.x - x) / w, (p1.y - y) / h]];
    createObject(v, type, x, y, w, h, { pts });
    commit();
  });
}
function startInk(v, p0, ev) {
  const g = geom(v.ent);
  const pre = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  pre.setAttribute('class', 'layer');
  pre.style.cssText = 'pointer-events:none;z-index:9;overflow:visible';
  pre.setAttribute('viewBox', `0 0 ${g.w1} ${g.h1}`);
  v.el.appendChild(pre);
  const pts = [[p0.x, p0.y]];
  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  Object.entries({ fill: 'none', stroke: S.style.color, 'stroke-width': S.style.width, 'stroke-linecap': 'round', 'stroke-linejoin': 'round', opacity: S.style.opacity }).forEach(([k, val]) => path.setAttribute(k, val));
  pre.appendChild(path);
  const draw = () => path.setAttribute('d', pts.map(([x, y], i) => `${i ? 'L' : 'M'}${x} ${y}`).join(' '));
  draw();
  dragLoop(ev, e => {
    const events = e.getCoalescedEvents ? e.getCoalescedEvents() : [e];
    for (const ce of events) { const p = pointIn(v, ce); const last = pts[pts.length - 1]; if (Math.hypot(p.x - last[0], p.y - last[1]) > 0.6) pts.push([p.x, p.y]); }
    draw();
  }, () => {
    pre.remove();
    if (pts.length === 1) pts.push([pts[0][0] + 0.5, pts[0][1] + 0.5]);
    const pad = S.style.width / 2 + 1;
    const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
    const x = Math.min(...xs) - pad, y = Math.min(...ys) - pad;
    const w = Math.max(...xs) - Math.min(...xs) + 2 * pad, h = Math.max(...ys) - Math.min(...ys) + 2 * pad;
    createObject(v, 'ink', x, y, w, h, { strokes: [pts.map(([px, py]) => [(px - x) / w, (py - y) / h])] });
    select(null); // keep drawing without selection boxes in the way
    commit();
  });
}
// highlight / underline / strike the text that is selected in the text layer
function markupFromSelection(type) {
  const sel = getSelection();
  if (!sel.rangeCount || sel.isCollapsed) return;
  const range = sel.getRangeAt(0);
  const rects = [...range.getClientRects()].filter(r => r.width > 0.5 && r.height > 0.5);
  if (!rects.length) return;
  const byView = new Map();
  for (const r of rects) {
    const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
    for (const v of views.values()) {
      const pr = v.el.getBoundingClientRect();
      if (cx >= pr.left && cx <= pr.right && cy >= pr.top && cy <= pr.bottom) {
        const g = geom(v.ent);
        const k = pr.width / g.w1;
        const b = { x: (r.left - pr.left) / k, y: (r.top - pr.top) / k, w: r.width / k, h: r.height / k };
        if (b.h < g.h1 * 0.15) { if (!byView.has(v)) byView.set(v, []); byView.get(v).push(b); }
        break;
      }
    }
  }
  let made = 0;
  for (const [v, boxes] of byView) {
    boxes.sort((a, b) => a.y - b.y || a.x - b.x);
    const lines = [];
    for (const b of boxes) {
      const L = lines.find(l => Math.abs(l.y + l.h / 2 - (b.y + b.h / 2)) < Math.max(l.h, b.h) * 0.5 && b.x <= l.x + l.w + Math.max(l.h, b.h) * 0.8);
      if (L) { const x2 = Math.max(L.x + L.w, b.x + b.w), y2 = Math.max(L.y + L.h, b.y + b.h); L.x = Math.min(L.x, b.x); L.y = Math.min(L.y, b.y); L.w = x2 - L.x; L.h = y2 - L.y; }
      else lines.push({ ...b });
    }
    for (const l of lines) { createObject(v, type, l.x, l.y, l.w, l.h); made++; }
  }
  sel.removeAllRanges();
  if (made) { select(null); commit(); }
}
// Edit Text: cover the original line with white and put an editable copy on top
function editTextSpan(v, span) {
  const pr = v.el.getBoundingClientRect();
  const g = geom(v.ent);
  const k = pr.width / g.w1;
  const r = span.getBoundingClientRect();
  const box = { x: (r.left - pr.left) / k, y: (r.top - pr.top) / k, w: r.width / k, h: r.height / k };
  const fam = (span.style.fontFamily || '').toLowerCase();
  const font = /mono|courier/.test(fam) ? 'Courier' : /serif/.test(fam) && !/sans/.test(fam) ? 'Times' : 'Helvetica';
  const size = Math.max(4, Math.round((parseFloat(getComputedStyle(span).fontSize) / k) * 10) / 10);
  createObject(v, 'whiteout', box.x - 1, box.y - 1, box.w + 2, box.h + 2);
  const lineH = size * 1.2;
  const o = createObject(v, 'text', box.x - TEXT_PAD, box.y + box.h / 2 - lineH / 2 - TEXT_PAD, box.w + 2 * TEXT_PAD, lineH + 2 * TEXT_PAD, { text: span.textContent, font, size, color: '#000000', bold: false, italic: false, align: 'left', autoW: true });
  commit();
  setTool('select');
  select(v.ent.id, o.id);
  startEditing(o.id, true);
}
// notes
function openNote(o) {
  const pop = $('#notepop');
  const el = objectEl(o.id);
  if (!el) return;
  const ws = $('#pdf-workspace').getBoundingClientRect();
  const r = el.getBoundingClientRect();
  pop.hidden = false;
  pop.style.left = `${clamp(r.right - ws.left + 8, 6, ws.width - pop.offsetWidth - 6)}px`;
  pop.style.top = `${clamp(r.top - ws.top, 6, ws.height - pop.offsetHeight - 6)}px`;
  const ta = $('#note-text');
  ta.value = o.text || '';
  pop.dataset.oid = o.id;
  ta.focus();
}
function closeNote(save = true) {
  const pop = $('#notepop');
  if (pop.hidden) return;
  const hit = objById(pop.dataset.oid);
  pop.hidden = true;
  if (!hit || !save) return;
  const text = $('#note-text').value;
  if (text !== (hit.o.text || '')) {
    hit.o.text = text;
    const v = views.get(hit.ent.id);
    if (v) renderObjects(v);
    commit();
  }
}

// ---------------------------------------------------------------- inserting pictures, signatures, stamps
function currentView() {
  const ent = S.entries[S.current];
  return ent ? views.get(ent.id) : null;
}
// a spot in the middle of the visible part of the current page
function visibleCenter(v) {
  const pr = v.el.getBoundingClientRect();
  const sr = scroller.getBoundingClientRect();
  const g = geom(v.ent);
  const k = pr.width / g.w1;
  const top = Math.max(pr.top, sr.top), bottom = Math.min(pr.bottom, sr.bottom);
  const left = Math.max(pr.left, sr.left), right = Math.min(pr.right, sr.right);
  return { x: ((left + right) / 2 - pr.left) / k, y: ((top + bottom) / 2 - pr.top) / k, g };
}
function imageSize(src) {
  return new Promise((resolve, reject) => { const im = new Image(); im.onload = () => resolve({ w: im.naturalWidth, h: im.naturalHeight }); im.onerror = () => reject(new Error('This picture could not be read.')); im.src = src; });
}
async function placeImage(src, maxFrac = 0.5, kind = 'image') {
  const v = currentView();
  if (!v) return;
  const { w: iw, h: ih } = await imageSize(src);
  const c = visibleCenter(v);
  let w = Math.min(c.g.w1 * maxFrac, iw * 0.75), h = (w * ih) / iw;
  if (h > c.g.h1 * 0.6) { h = c.g.h1 * 0.6; w = (h * iw) / ih; }
  setTool('select');
  createObject(v, 'image', c.x - w / 2, c.y - h / 2, w, h, { src, kind });
  commit();
}
async function insertImage() {
  const r = await api.pickImage();
  if (!r) return;
  if (r.error) { await alertBox('Image', r.error); return; }
  await placeImage(r.url);
}
const STAMPS = [
  ['APPROVED', '#1e8e3e'], ['NOT APPROVED', '#c5221f'], ['DRAFT', '#1f5fbf'], ['FINAL', '#1e8e3e'],
  ['CONFIDENTIAL', '#c5221f'], ['FOR COMMENT', '#1f5fbf'], ['COMPLETED', '#1e8e3e'], ['VOID', '#c5221f'],
];
function placeStamp(text, color) {
  const v = currentView();
  if (!v) return;
  const c = visibleCenter(v);
  const h = 36, w = Math.max(90, text.length * h * 0.36 + h * 0.9);
  setTool('select');
  createObject(v, 'stamp', c.x - w / 2, c.y - h / 2, w, h, { text, color });
  commit();
}

// signature dialog: draw, type or reuse a saved one
function signatureDialog() {
  const d = $('#modal');
  const fonts = ['Segoe Script', 'Ink Free', 'Lucida Handwriting', 'Brush Script MT', 'Gabriola'];
  d.classList.add('wide');
  d.innerHTML = `<div class="dbody"><h2>Add Signature</h2>
    <div class="sig-tabs"><button data-st="draw" class="on">Draw</button><button data-st="type">Type</button>${S.signatures.length ? '<button data-st="saved">Saved</button>' : ''}</div>
    <div class="sig-pane" data-sp="draw"><div class="sig-line"><canvas class="sig-pad" width="1120" height="400"></canvas></div>
      <div class="sig-row"><span style="color:var(--muted)">Sign above the line with your mouse, pen or finger.</span><span class="spacer"></span><button class="btn" data-sig="clear">Clear</button></div></div>
    <div class="sig-pane" data-sp="type" hidden><input class="sig-name" placeholder="Type your name" spellcheck="false">
      <div class="sig-type-preview"></div><div class="sig-fonts">${fonts.map((f, i) => `<button data-font="${esc(f)}" class="${i ? '' : 'on'}" style="font-family:'${esc(f)}'">Signature</button>`).join('')}</div></div>
    <div class="sig-pane" data-sp="saved" hidden><div class="sig-saved">${S.signatures.map((s, i) => `<div class="s" data-use="${i}"><img src="${s}" alt="Saved signature ${i + 1}"><button class="x" data-del="${i}" title="Remove">✕</button></div>`).join('')}</div></div>
    <div class="sig-row"><label class="chk"><input type="checkbox" class="sig-keep" checked> Save this signature for next time</label><span class="spacer"></span>
      <button class="btn" data-sig="cancel">Cancel</button><button class="btn primary" data-sig="use">Place Signature</button></div></div>`;
  const pad = $('.sig-pad', d);
  const ctx = pad.getContext('2d');
  ctx.lineCap = 'round'; ctx.lineJoin = 'round'; ctx.lineWidth = 5; ctx.strokeStyle = '#13235b';
  let drawn = false;
  pad.addEventListener('pointerdown', ev => {
    const r = pad.getBoundingClientRect();
    const sx = pad.width / r.width;
    let last = [(ev.clientX - r.left) * sx, (ev.clientY - r.top) * sx];
    ctx.beginPath(); ctx.arc(last[0], last[1], 2.5, 0, Math.PI * 2); ctx.fillStyle = ctx.strokeStyle; ctx.fill();
    drawn = true;
    dragLoop(ev, e => {
      for (const ce of (e.getCoalescedEvents ? e.getCoalescedEvents() : [e])) {
        const p = [(ce.clientX - r.left) * sx, (ce.clientY - r.top) * sx];
        const pressure = ce.pressure && ce.pointerType === 'pen' ? 2 + ce.pressure * 6 : 5;
        ctx.lineWidth = pressure;
        ctx.beginPath(); ctx.moveTo(last[0], last[1]); ctx.lineTo(p[0], p[1]); ctx.stroke();
        last = p;
      }
    }, () => {});
  });
  let font = fonts[0];
  const nameIn = $('.sig-name', d), preview = $('.sig-type-preview', d);
  const updatePreview = () => { preview.style.fontFamily = `'${font}'`; preview.textContent = nameIn.value || 'Your Name'; preview.style.color = nameIn.value ? '#13235b' : '#b9c0cc'; };
  updatePreview();
  nameIn.addEventListener('input', updatePreview);
  let tab = 'draw';
  const finish = () => d.close();
  d.addEventListener('close', () => d.classList.remove('wide'), { once: true });
  d.addEventListener('click', async ev => {
    const t = ev.target.closest('[data-st]');
    if (t) { tab = t.dataset.st; for (const b of $$('[data-st]', d)) b.classList.toggle('on', b === t); for (const p of $$('[data-sp]', d)) p.hidden = p.dataset.sp !== tab; if (tab === 'type') nameIn.focus(); return; }
    const f = ev.target.closest('[data-font]');
    if (f) { font = f.dataset.font; for (const b of $$('[data-font]', d)) b.classList.toggle('on', b === f); updatePreview(); return; }
    const del = ev.target.closest('[data-del]');
    if (del) { ev.stopPropagation(); S.signatures.splice(+del.dataset.del, 1); api.setSetting('signatures', S.signatures); del.closest('.s').remove(); return; }
    const use = ev.target.closest('[data-use]');
    if (use) { finish(); await placeImage(S.signatures[+use.dataset.use], 0.32, 'signature'); return; }
    const a = ev.target.closest('[data-sig]');
    if (!a) return;
    if (a.dataset.sig === 'cancel') { finish(); return; }
    if (a.dataset.sig === 'clear') { ctx.clearRect(0, 0, pad.width, pad.height); drawn = false; return; }
    if (a.dataset.sig === 'use') {
      let url = null;
      if (tab === 'draw') { if (!drawn) { toast('Draw your signature first.'); return; } url = trimCanvas(pad); }
      else if (tab === 'type') {
        if (!nameIn.value.trim()) { toast('Type your name first.'); nameIn.focus(); return; }
        url = typedSignature(nameIn.value.trim(), font);
      } else { toast('Click a saved signature.'); return; }
      if ($('.sig-keep', d).checked && !S.signatures.includes(url)) { S.signatures.unshift(url); S.signatures = S.signatures.slice(0, 8); api.setSetting('signatures', S.signatures); }
      finish();
      await placeImage(url, 0.32, 'signature');
    }
  });
  d.showModal();
}
function trimCanvas(c) {
  const ctx = c.getContext('2d');
  const { data, width, height } = ctx.getImageData(0, 0, c.width, c.height);
  let x0 = width, y0 = height, x1 = -1, y1 = -1;
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) if (data[(y * width + x) * 4 + 3] > 8) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
  if (x1 < 0) return c.toDataURL('image/png');
  const pad = 6;
  x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad); x1 = Math.min(width - 1, x1 + pad); y1 = Math.min(height - 1, y1 + pad);
  const out = document.createElement('canvas');
  out.width = x1 - x0 + 1; out.height = y1 - y0 + 1;
  out.getContext('2d').drawImage(c, x0, y0, out.width, out.height, 0, 0, out.width, out.height);
  return out.toDataURL('image/png');
}
function typedSignature(name, font) {
  const c = document.createElement('canvas');
  const ctx = c.getContext('2d');
  const size = 120;
  ctx.font = `${size}px '${font}'`;
  const w = Math.ceil(ctx.measureText(name).width) + 40;
  c.width = w; c.height = Math.round(size * 1.6);
  ctx.font = `${size}px '${font}'`;
  ctx.fillStyle = '#13235b';
  ctx.textBaseline = 'middle';
  ctx.fillText(name, 20, c.height / 2);
  return trimCanvas(c);
}

// ---------------------------------------------------------------- pages
function targetIndexes() {
  if (S.selPages.size) return S.entries.map((e, i) => (S.selPages.has(e.id) ? i : -1)).filter(i => i >= 0);
  return S.entries.length ? [S.current] : [];
}
function afterStructureChange(focusIndex) {
  const keep = focusIndex ?? S.current;
  buildPages();
  S.current = clamp(keep, 0, S.entries.length - 1);
  if (S.fit) applyFit(); else layoutAfterZoom();
  scrollToPage(S.current, false);
  commit();
  requestAnimationFrame(currentFromScroll);
}
function rotatePages(dir) {
  const idx = targetIndexes();
  for (const i of idx) S.entries[i].rot = (S.entries[i].rot + dir * 90 + 360) % 360;
  afterStructureChange(idx[0]);
}
async function deletePages() {
  const idx = targetIndexes();
  if (!idx.length) return;
  if (idx.length >= S.entries.length) { await alertBox('Delete Pages', 'A PDF needs at least one page, so the last page can’t be deleted.'); return; }
  if (idx.length > 1 && !(await confirmBox('Delete Pages', `Delete ${idx.length} pages?`, 'Delete', 'Cancel'))) return;
  const ids = new Set(idx.map(i => S.entries[i].id));
  S.entries = S.entries.filter(e => !ids.has(e.id));
  S.selPages.clear();
  select(null);
  afterStructureChange(Math.min(idx[0], S.entries.length - 1));
  toast(`Deleted ${idx.length} page${idx.length === 1 ? '' : 's'}.`);
}
function movePages(dir) {
  const idx = targetIndexes();
  if (!idx.length) return;
  if ((dir < 0 && idx[0] === 0) || (dir > 0 && idx[idx.length - 1] === S.entries.length - 1)) return;
  const order = dir < 0 ? idx : [...idx].reverse();
  for (const i of order) { const j = i + dir; [S.entries[i], S.entries[j]] = [S.entries[j], S.entries[i]]; }
  afterStructureChange(idx[0] + dir);
}
function cloneObjects(objs) { return objs.map(o => ({ ...o, id: uid() })); }
function duplicatePages() {
  const idx = targetIndexes();
  const last = idx[idx.length - 1];
  const copies = idx.map(i => ({ ...S.entries[i], id: uid(), objects: cloneObjects(S.entries[i].objects) }));
  S.entries.splice(last + 1, 0, ...copies);
  afterStructureChange(last + 1);
}
function insertBlank() {
  const cur = S.entries[S.current];
  const g = cur ? geom(cur) : { w1: 612, h1: 792 };
  const ent = { id: uid(), src: 0, index: 0, rot: 0, blank: { w: g.w1, h: g.h1 }, objects: [] };
  S.entries.splice(S.current + 1, 0, ent);
  afterStructureChange(S.current + 1);
}
async function insertPdf() {
  const after = S.current;
  const p = await api.openDialog('pdf');
  if (!p) return;
  const r = await api.readFileBytes(p);
  if (r.error) { await alertBox('Insert Pages', r.error); return; }
  await insertPdfBytes(r, after);
}
async function insertPdfBytes(r, after = S.current) {
  // the pages go after the page that was current when the command started
  const anchorId = S.entries[after] && S.entries[after].id;
  setBusy('Adding pages…');
  try {
    const loaded = await loadPdfjs(r.bytes, r.name);
    const src = S.sources.push({ bytes: await decryptWithPassword(r.bytes, loaded), name: r.name, ...loaded }) - 1;
    const add = loaded.pages.map((pg, i) => ({ id: uid(), src, index: i, rot: 0, blank: null, objects: [] }));
    const at = S.entries.findIndex(e => e.id === anchorId) + 1;
    S.entries.splice(at, 0, ...add);
    afterStructureChange(at);
    toast(`Added ${add.length} page${add.length === 1 ? '' : 's'} from ${r.name}.`);
  } catch (err) {
    if (err && /destroy/i.test(err.message)) return;
    await alertBox('Insert Pages', `"${r.name}" couldn't be read as a PDF.`);
  } finally { setBusy(null); }
}
async function insertImagePage() {
  const r = await api.pickImage();
  if (!r) return;
  if (r.error) { await alertBox('Image Page', r.error); return; }
  const { w: iw, h: ih } = await imageSize(r.url);
  const cur = S.entries[S.current];
  const g = cur ? geom(cur) : { w1: 612, h1: 792 };
  const portrait = ih >= iw;
  const pw = portrait ? Math.min(g.w1, g.h1) : Math.max(g.w1, g.h1);
  const ph = portrait ? Math.max(g.w1, g.h1) : Math.min(g.w1, g.h1);
  const fit = Math.min(pw / iw, ph / ih);
  const w = iw * fit, h = ih * fit;
  const ent = { id: uid(), src: 0, index: 0, rot: 0, blank: { w: pw, h: ph }, objects: [] };
  const geomNew = { V1: viewportMatrix([0, 0, pw, ph], 0, 1) };
  ent.objects.push({ id: uid(), type: 'image', m: mul(inv(geomNew.V1), T((pw - w) / 2, (ph - h) / 2)), w, h, src: r.url, kind: 'image' });
  S.entries.splice(S.current + 1, 0, ent);
  afterStructureChange(S.current + 1);
}
async function extractPages() {
  const idx = targetIndexes();
  if (!idx.length) return;
  setBusy('Extracting pages…');
  try {
    const bytes = await buildPdf({ only: idx.map(i => S.entries[i]) });
    const base = S.name.replace(/\.pdf$/i, '');
    const r = await api.savePdf({ bytes, name: `${base} (pages ${pageRangeLabel(idx)})`, title: 'Save Extracted Pages', copy: true }, true);
    if (r && r.path) toast(`Saved ${idx.length} page${idx.length === 1 ? '' : 's'} to ${r.name}.`, () => api.showInFolder(r.path), 'Show in folder');
    else if (r && r.error) await alertBox("Couldn't save", r.error);
  } catch (err) { await alertBox('Extract Pages', friendly(err)); } finally { setBusy(null); }
}
function pageRangeLabel(idx) {
  const n = idx.map(i => i + 1);
  const parts = [];
  for (let i = 0; i < n.length; i++) { let j = i; while (j + 1 < n.length && n[j + 1] === n[j] + 1) j++; parts.push(i === j ? `${n[i]}` : `${n[i]}-${n[j]}`); i = j; }
  return parts.join(', ');
}

// ---------------------------------------------------------------- thumbnails
const thumbIO = new IntersectionObserver(items => {
  for (const it of items) if (it.isIntersecting) renderThumb(it.target);
}, { root: $('#thumb-list'), rootMargin: '400px 0px' });
function renderThumbs() {
  thumbIO.disconnect();
  const list = $('#thumb-list');
  list.innerHTML = '';
  S.entries.forEach((ent, i) => {
    const g = geom(ent);
    const w = g.w1 >= g.h1 ? 130 : Math.round(130 * (g.w1 / g.h1));
    const h = Math.round((w * g.h1) / g.w1);
    const t = document.createElement('div');
    t.className = `pthumb${S.selPages.has(ent.id) ? ' sel' : ''}${i === S.current ? ' current' : ''}`;
    t.dataset.eid = ent.id;
    t.draggable = true;
    t.innerHTML = `<div class="tcan" style="width:${w}px;height:${h}px"></div><span class="num">${i + 1}</span>`;
    list.appendChild(t);
    thumbIO.observe(t);
  });
}
async function renderThumb(t) {
  if (t.dataset.done) return;
  t.dataset.done = '1';
  const ent = entryById(t.dataset.eid);
  if (!ent) return;
  const box = $('.tcan', t);
  const g = geom(ent);
  const w = parseFloat(box.style.width);
  const s = w / g.w1;
  const ratio = window.devicePixelRatio || 1;
  const c = document.createElement('canvas');
  c.width = Math.round(w * ratio); c.height = Math.round(parseFloat(box.style.height) * ratio);
  c.style.width = box.style.width; c.style.height = box.style.height;
  const ctx = c.getContext('2d', { alpha: false });
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, c.width, c.height);
  if (!ent.blank) {
    try { await pageProxy(ent).render({ canvasContext: ctx, viewport: pageProxy(ent).getViewport({ scale: s * ratio, rotation: g.rotation }), transform: g.cropped ? g.shift(s * ratio) : null }).promise; } catch { /* page removed meanwhile */ }
  }
  box.prepend(c);
  thumbObjects(t);
}
// miniature of the items added to a page, drawn over its thumbnail
function thumbObjects(t) {
  const ent = entryById(t.dataset.eid);
  const box = $('.tcan', t);
  if (!ent || !box) return;
  let layer = $('.tobj', box);
  if (!layer) { layer = document.createElement('div'); layer.className = 'tobj'; box.appendChild(layer); }
  const g = geom(ent);
  const s = parseFloat(box.style.width) / g.w1;
  const Vs = mul(Sc(s), g.V1);
  layer.innerHTML = '';
  for (const o of ent.objects) {
    if (o.type === 'note') continue;
    const el = objElement(o, Vs, s);
    el.removeAttribute('data-oid'); // only the real item on the page carries the id
    layer.appendChild(el);
  }
}
let thumbTimer = null;
function refreshThumbObjects() {
  clearTimeout(thumbTimer);
  thumbTimer = setTimeout(() => { for (const t of $$('.pthumb[data-done]')) thumbObjects(t); }, 250);
}
function syncThumbClasses() {
  for (const t of $$('.pthumb')) {
    const i = S.entries.findIndex(e => e.id === t.dataset.eid);
    t.classList.toggle('current', i === S.current);
    t.classList.toggle('sel', S.selPages.has(t.dataset.eid));
  }
}
let thumbAnchor = null;
$('#thumb-list').addEventListener('click', ev => {
  const t = ev.target.closest('.pthumb');
  if (!t) return;
  const i = S.entries.findIndex(e => e.id === t.dataset.eid);
  if (ev.shiftKey && thumbAnchor != null) {
    S.selPages.clear();
    const [a, b] = [Math.min(thumbAnchor, i), Math.max(thumbAnchor, i)];
    for (let k = a; k <= b; k++) S.selPages.add(S.entries[k].id);
  } else if (ev.ctrlKey || ev.metaKey) {
    if (S.selPages.has(t.dataset.eid)) S.selPages.delete(t.dataset.eid); else S.selPages.add(t.dataset.eid);
    thumbAnchor = i;
  } else {
    S.selPages.clear();
    S.selPages.add(t.dataset.eid);
    thumbAnchor = i;
  }
  S.current = i;
  syncThumbClasses();
  scrollToPage(i);
});
let dragIds = null;
$('#thumb-list').addEventListener('dragstart', ev => {
  const t = ev.target.closest('.pthumb');
  if (!t) return;
  dragIds = S.selPages.has(t.dataset.eid) ? S.entries.filter(e => S.selPages.has(e.id)).map(e => e.id) : [t.dataset.eid];
  ev.dataTransfer.effectAllowed = 'move';
  ev.dataTransfer.setData('text/x-nisar-pages', dragIds.join(','));
  for (const id of dragIds) { const el = $(`.pthumb[data-eid="${id}"]`); if (el) el.classList.add('dragging'); }
});
$('#thumb-list').addEventListener('dragover', ev => {
  if (!dragIds) return;
  ev.preventDefault();
  for (const el of $$('.pthumb.drop-before, .pthumb.drop-after')) el.classList.remove('drop-before', 'drop-after');
  const t = ev.target.closest('.pthumb');
  if (!t) return;
  const r = t.getBoundingClientRect();
  t.classList.add(ev.clientY < r.top + r.height / 2 ? 'drop-before' : 'drop-after');
});
$('#thumb-list').addEventListener('drop', ev => {
  if (!dragIds) return;
  ev.preventDefault();
  const t = $('.pthumb.drop-before, .pthumb.drop-after');
  if (t && !dragIds.includes(t.dataset.eid)) {
    const after = t.classList.contains('drop-after');
    const moving = S.entries.filter(e => dragIds.includes(e.id));
    const rest = S.entries.filter(e => !dragIds.includes(e.id));
    let at = rest.findIndex(e => e.id === t.dataset.eid) + (after ? 1 : 0);
    rest.splice(at, 0, ...moving);
    S.entries = rest;
    afterStructureChange(at);
  }
  dragIds = null;
});
$('#thumb-list').addEventListener('dragend', () => {
  dragIds = null;
  for (const el of $$('.pthumb.dragging, .pthumb.drop-before, .pthumb.drop-after')) el.classList.remove('dragging', 'drop-before', 'drop-after');
});

// ---------------------------------------------------------------- zoom, scrolling, status
function pageTop(i) { const ent = S.entries[i]; const v = ent && views.get(ent.id); return v ? v.el.offsetTop : 0; }
function scrollToPage(i, smooth = true) {
  const ent = S.entries[i];
  const v = ent && views.get(ent.id);
  if (!v) return;
  scroller.scrollTo({ top: v.el.offsetTop - 14, behavior: smooth ? 'smooth' : 'auto' });
  S.current = i;
  updateStatus();
}
function layoutAfterZoom() {
  const s = scale();
  for (const v of views.values()) { sizeView(v); renderObjects(v); if (v.visible) renderView(v); else releaseView(v); }
  positionObjbar();
  $('#zoom-input').value = `${Math.round(S.zoom * 100)}%`;
  $('#zoom-range').value = Math.round(S.zoom * 100);
  $('#st-zoom').textContent = `${Math.round(S.zoom * 100)}%`;
  void s;
}
function setZoom(z, fit = null) {
  z = clamp(Math.round(z * 1000) / 1000, 0.1, 4);
  const i = S.current;
  const v = S.entries[i] && views.get(S.entries[i].id);
  const frac = v ? (scroller.scrollTop - v.el.offsetTop) / (v.el.offsetHeight || 1) : 0;
  S.zoom = z;
  S.fit = fit;
  layoutAfterZoom();
  if (v) scroller.scrollTop = v.el.offsetTop + frac * v.el.offsetHeight;
  updateToggles();
}
function maxPageWidth() { return Math.max(1, ...S.entries.slice(Math.max(0, S.current - 3), S.current + 4).map(e => geom(e).w1)); }
function fitWidth() { if (!S.entries.length) return; setZoom((scroller.clientWidth - 48 - 16) / (maxPageWidth() * PT), 'width'); }
function fitPage() {
  const ent = S.entries[S.current];
  if (!ent) return;
  const g = geom(ent);
  setZoom(Math.min((scroller.clientWidth - 64) / (g.w1 * PT), (scroller.clientHeight - 40) / (g.h1 * PT)), 'page');
  scrollToPage(S.current, false);
}
function applyFit() { if (S.fit === 'width') fitWidth(); else if (S.fit === 'page') fitPage(); else layoutAfterZoom(); }
const ZOOMS = [0.1, 0.25, 0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4];
function stepZoom(dir) {
  const z = S.zoom;
  const next = dir > 0 ? ZOOMS.find(x => x > z + 0.001) || 4 : [...ZOOMS].reverse().find(x => x < z - 0.001) || 0.1;
  setZoom(next);
}
let scrollQueued = false;
// the current page is the one at the upper third of the window
function currentFromScroll() {
  const y = scroller.scrollTop + scroller.clientHeight * 0.35;
  let cur = 0;
  for (let i = 0; i < S.entries.length; i++) { const v = views.get(S.entries[i].id); if (v && v.el.offsetTop <= y) cur = i; else break; }
  if (cur !== S.current) { S.current = cur; updateStatus(); }
}
scroller.addEventListener('scroll', () => {
  if (scrollQueued) return;
  scrollQueued = true;
  requestAnimationFrame(() => {
    scrollQueued = false;
    currentFromScroll();
    positionObjbar();
    if (!$('#notepop').hidden) closeNote(true);
  });
}, { passive: true });
scroller.addEventListener('wheel', ev => {
  if (!ev.ctrlKey) return;
  ev.preventDefault();
  setZoom(S.zoom * (ev.deltaY < 0 ? 1.1 : 1 / 1.1));
}, { passive: false });
window.addEventListener('resize', () => { closeMenu(); if (S.fit) applyFit(); positionObjbar(); });
function updateStatus() {
  $('#st-page').textContent = S.entries.length ? `Page ${S.current + 1} of ${S.entries.length}` : 'No document';
  syncThumbClasses();
  const t = $('.pthumb.current');
  if (t) { const list = $('#thumb-list'); const tr = t.getBoundingClientRect(), lr = list.getBoundingClientRect(); if (tr.top < lr.top || tr.bottom > lr.bottom) t.scrollIntoView({ block: 'nearest' }); }
}
async function goToDest(dest) {
  try {
    const doc = S.sources[0].doc;
    const explicit = typeof dest === 'string' ? await doc.getDestination(dest) : dest;
    if (!explicit) return;
    const index = typeof explicit[0] === 'object' ? await doc.getPageIndex(explicit[0]) : explicit[0];
    const i = S.entries.findIndex(e => e.src === 0 && !e.blank && e.index === index);
    if (i >= 0) scrollToPage(i);
  } catch { /* broken link */ }
}
async function gotoPageDialog() {
  if (!S.entries.length) return;
  const r = await askForm('Go To Page', [{ id: 'n', label: `Page number (1–${S.entries.length})`, value: S.current + 1, type: 'number', attrs: `min="1" max="${S.entries.length}"` }], 'Go');
  if (r) scrollToPage(clamp((parseInt(r.n, 10) || 1) - 1, 0, S.entries.length - 1), false);
}

// ---------------------------------------------------------------- find
const find = { query: '', matches: [], index: -1 };
const findInput = $('#find-input');
function openFind() {
  $('#findbar').hidden = false;
  const s = getSelection().toString().trim();
  if (s && s.length < 80) findInput.value = s;
  findInput.focus();
  findInput.select();
  runFind();
}
function closeFind() {
  $('#findbar').hidden = true;
  find.query = '';
  find.matches = [];
  for (const l of $$('.search-layer')) l.innerHTML = '';
}
let findTimer = null;
async function runFind() {
  const q = findInput.value;
  const mc = $('#find-case').checked;
  find.query = q;
  find.matches = [];
  find.index = -1;
  if (!q) { $('#find-count').textContent = ''; for (const l of $$('.search-layer')) l.innerHTML = ''; return; }
  const needle = mc ? q : q.toLocaleLowerCase();
  for (let ei = 0; ei < S.entries.length; ei++) {
    const ent = S.entries[ei];
    if (ent.blank) continue;
    const { str } = await pageString(ent);
    if (find.query !== q) return; // a newer search started
    const hay = mc ? str : str.toLocaleLowerCase();
    let i = 0;
    while ((i = hay.indexOf(needle, i)) !== -1) { find.matches.push({ ei, start: i, end: i + q.length }); i += Math.max(1, q.length); }
  }
  $('#find-count').textContent = find.matches.length ? '' : 'No results';
  if (find.matches.length) {
    const from = find.matches.findIndex(m => m.ei >= S.current);
    gotoMatch(from < 0 ? 0 : from);
  }
  for (const v of views.values()) paintFind(v);
}
async function pageString(ent) {
  const tc = await textContent(ent);
  const items = tc.items.filter(it => typeof it.str === 'string');
  let str = '';
  const starts = [];
  for (const it of items) { starts.push(str.length); str += it.str + (it.hasEOL ? ' ' : ''); }
  return { str, items, starts };
}
function gotoMatch(i) {
  if (!find.matches.length) return;
  find.index = (i + find.matches.length) % find.matches.length;
  const m = find.matches[find.index];
  $('#find-count').textContent = `${find.index + 1} of ${find.matches.length}`;
  const ent = S.entries[m.ei];
  const v = views.get(ent.id);
  if (!v) return;
  S.current = m.ei;
  updateStatus();
  // scroll the page into view, then centre the match once its text layer exists
  if (v.el.offsetTop > scroller.scrollTop + scroller.clientHeight || v.el.offsetTop + v.el.offsetHeight < scroller.scrollTop) scroller.scrollTop = v.el.offsetTop - 20;
  const wait = () => { if (v.textLayer) { for (const vv of views.values()) paintFind(vv); const cur = $('.search-layer .cur', v.el); if (cur) { const r = cur.getBoundingClientRect(), sr = scroller.getBoundingClientRect(); if (r.top < sr.top + 40 || r.bottom > sr.bottom - 40) scroller.scrollTop += r.top - sr.top - sr.height / 2; } } else setTimeout(wait, 60); };
  wait();
}
async function paintFind(v) {
  const layer = v.el.querySelector('.search-layer');
  layer.innerHTML = '';
  if (!find.query || !v.textLayer || v.ent.blank) return;
  const ei = S.entries.indexOf(v.ent);
  const mine = find.matches.map((m, i) => ({ ...m, i })).filter(m => m.ei === ei);
  if (!mine.length) return;
  const { starts } = await pageString(v.ent);
  const divs = v.textLayer.textDivs;
  const pr = v.el.getBoundingClientRect();
  for (const m of mine) {
    // find the text-layer spans covering [start, end)
    const range = document.createRange();
    const pos = off => {
      let k = starts.length - 1;
      while (k > 0 && starts[k] > off) k--;
      const node = divs[k] && divs[k].firstChild;
      return node ? [node, clamp(off - starts[k], 0, node.length)] : null;
    };
    const a = pos(m.start), b = pos(m.end);
    if (!a || !b) continue;
    try { range.setStart(...a); range.setEnd(...b); } catch { continue; }
    for (const r of range.getClientRects()) {
      const d = document.createElement('div');
      d.className = m.i === find.index ? 'hit cur' : 'hit';
      Object.assign(d.style, { position: 'absolute', left: `${r.left - pr.left}px`, top: `${r.top - pr.top}px`, width: `${r.width}px`, height: `${r.height}px`, background: m.i === find.index ? 'rgba(255,150,40,.55)' : 'rgba(255,225,70,.5)', mixBlendMode: 'multiply', borderRadius: '2px' });
      layer.appendChild(d);
    }
  }
}
findInput.addEventListener('input', () => { clearTimeout(findTimer); findTimer = setTimeout(runFind, 200); });
$('#find-case').addEventListener('change', runFind);

// ---------------------------------------------------------------- history
function snapshot() {
  return { entries: S.entries.map(e => ({ ...e, objects: e.objects.map(o => ({ ...o })) })), form: { ...S.form } };
}
function resetHistory() { S.history = [snapshot()]; S.hIndex = 0; }
function commit() {
  S.history.splice(S.hIndex + 1);
  S.history.push(snapshot());
  if (S.history.length > 150) S.history.shift();
  S.hIndex = S.history.length - 1;
  markDirty();
  refreshThumbObjects();
}
function structureKey(entries) { return entries.map(e => `${e.id}/${e.rot}`).join('|'); }
function restore(i) {
  finishEditingSilently();
  const snap = S.history[i];
  const structural = structureKey(snap.entries) !== structureKey(S.entries);
  S.entries = snap.entries.map(e => ({ ...e, objects: e.objects.map(o => ({ ...o })) }));
  S.form = { ...snap.form };
  S.hIndex = i;
  select(null);
  if (structural) { buildPages(); applyFit(); }
  else for (const ent of S.entries) { const v = views.get(ent.id); if (v) { v.ent = ent; renderObjects(v); } }
  for (const el of $$('.fld')) {
    const name = el.dataset.name;
    if (!(name in S.form)) continue;
    const val = S.form[name];
    if (el.type === 'checkbox') el.checked = !!val; else if (el.type === 'radio') el.checked = el.value === val; else if (el.tagName !== 'SELECT') el.value = val;
  }
  S.dirty = true;
  updateTitle();
}
function finishEditingSilently() { if (S.editing) { const id = S.editing; S.editing = null; const el = objectEl(id); if (el) el.classList.remove('editing'); } }
function undo() { if (S.editing) { finishEditing(); } if (S.hIndex > 0) restore(S.hIndex - 1); }
function redo() { if (S.hIndex < S.history.length - 1) restore(S.hIndex + 1); }

// ---------------------------------------------------------------- writing the PDF
const hexRgb = hex => { const n = parseInt((hex || '#000000').slice(1), 16); return PL.rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255); };
const STD = {
  Helvetica: ['Helvetica', 'HelveticaBold', 'HelveticaOblique', 'HelveticaBoldOblique'],
  Times: ['TimesRoman', 'TimesRomanBold', 'TimesRomanItalic', 'TimesRomanBoldItalic'],
  Courier: ['Courier', 'CourierBold', 'CourierOblique', 'CourierBoldOblique'],
};
function wrapLines(text, maxW, measure) {
  const out = [];
  for (const para of String(text).split('\n')) {
    if (!para) { out.push(''); continue; }
    const words = para.split(/(\s+)/);
    let line = '';
    for (const w of words) {
      if (!w) continue;
      const tryLine = line + w;
      if (measure(tryLine.replace(/\s+$/, '')) <= maxW || !line) {
        if (measure(tryLine.replace(/\s+$/, '')) > maxW && !line) {
          // a single word wider than the box: break it
          let part = '';
          for (const ch of w) { if (measure(part + ch) > maxW && part) { out.push(part); part = ch; } else part += ch; }
          line = part;
        } else line = tryLine;
      } else { out.push(line.replace(/\s+$/, '')); line = /^\s+$/.test(w) ? '' : w; }
    }
    out.push(line.replace(/\s+$/, ''));
  }
  return out;
}
async function dataUrlBytes(url) { const r = await fetch(url); return new Uint8Array(await r.arrayBuffer()); }
async function toPngBytes(url) {
  const im = new Image();
  await new Promise((res, rej) => { im.onload = res; im.onerror = rej; im.src = url; });
  const c = document.createElement('canvas');
  c.width = im.naturalWidth; c.height = im.naturalHeight;
  c.getContext('2d').drawImage(im, 0, 0);
  return dataUrlBytes(c.toDataURL('image/png'));
}
async function embedImage(doc, url, cache) {
  if (cache.has(url)) return cache.get(url);
  let img;
  if (/^data:image\/png/.test(url)) img = await doc.embedPng(await dataUrlBytes(url));
  else if (/^data:image\/jpe?g/.test(url)) img = await doc.embedJpg(await dataUrlBytes(url));
  else img = await doc.embedPng(await toPngBytes(url));
  cache.set(url, img);
  return img;
}
// non-Latin text (Arabic, Urdu, Chinese…) can't use the built-in PDF fonts, so it is drawn as a sharp picture
function rasterText(o) {
  const k = 4;
  const c = document.createElement('canvas');
  c.width = Math.ceil(o.w * k); c.height = Math.ceil(o.h * k);
  const ctx = c.getContext('2d');
  ctx.scale(k, k);
  ctx.font = `${o.italic ? 'italic ' : ''}${o.bold ? '700 ' : '400 '}${o.size}px ${CSS_FONT[o.font] || CSS_FONT.Helvetica}`;
  ctx.fillStyle = o.color;
  const rtl = /[֐-ࣿיִ-ﻼ]/.test(o.text);
  ctx.direction = rtl ? 'rtl' : 'ltr';
  const lines = wrapLines(o.text, o.w - 2 * TEXT_PAD + 0.5, t => ctx.measureText(t).width);
  const lh = o.size * 1.2;
  lines.forEach((line, i) => {
    const align = o.align === 'center' ? 'center' : o.align === 'right' || (rtl && o.align !== 'left') ? 'right' : 'left';
    ctx.textAlign = align;
    const x = align === 'center' ? o.w / 2 : align === 'right' ? o.w - TEXT_PAD : TEXT_PAD;
    ctx.fillText(line, x, TEXT_PAD + i * lh + BASELINE[o.font] * o.size);
  });
  return c.toDataURL('image/png');
}
async function drawObject(doc, page, o, pre, cache) {
  const { pushGraphicsState, popGraphicsState, concatTransformationMatrix, LineCapStyle, BlendMode } = PL;
  if (o.type === 'note') return;
  const cm = mul(pre, mul(o.m, [1, 0, 0, -1, 0, o.h])); // local box with y up -> page space
  page.pushOperators(pushGraphicsState(), concatTransformationMatrix(...cm));
  const W = o.w, H = o.h;
  const color = hexRgb(o.color);
  const op = o.opacity ?? 1;
  const sw = o.width || 2;
  switch (o.type) {
    case 'text': {
      const names = STD[o.font] || STD.Helvetica;
      const fname = names[(o.bold ? 1 : 0) + (o.italic ? 2 : 0)];
      if (!cache.fonts.has(fname)) cache.fonts.set(fname, await doc.embedFont(PL.StandardFonts[fname]));
      const font = cache.fonts.get(fname);
      let encodable = true;
      try { font.encodeText(o.text.replace(/\n/g, '')); } catch { encodable = false; }
      if (!encodable) {
        const img = await embedImage(doc, rasterText(o), cache.images);
        page.drawImage(img, { x: 0, y: 0, width: W, height: H });
        break;
      }
      const lines = wrapLines(o.text, (o.autoW ? Infinity : W - 2 * TEXT_PAD + 0.5), t => font.widthOfTextAtSize(t, o.size));
      const lh = o.size * 1.2;
      lines.forEach((line, i) => {
        if (!line) return;
        const lw = font.widthOfTextAtSize(line, o.size);
        const x = o.align === 'center' ? (W - lw) / 2 : o.align === 'right' ? W - TEXT_PAD - lw : TEXT_PAD;
        page.drawText(line, { x, y: H - TEXT_PAD - BASELINE[o.font] * o.size - i * lh, size: o.size, font, color, opacity: o.opacity ?? 1 });
      });
      break;
    }
    case 'image': {
      const img = await embedImage(doc, o.src, cache.images);
      page.drawImage(img, { x: 0, y: 0, width: W, height: H });
      break;
    }
    case 'rect':
      page.drawRectangle({ x: sw / 2, y: sw / 2, width: Math.max(0, W - sw), height: Math.max(0, H - sw), borderColor: color, borderWidth: sw, borderOpacity: op, color: o.fill ? hexRgb(o.fill) : undefined, opacity: o.fill ? op : undefined });
      break;
    case 'ellipse':
      page.drawEllipse({ x: W / 2, y: H / 2, xScale: Math.max(0, (W - sw) / 2), yScale: Math.max(0, (H - sw) / 2), borderColor: color, borderWidth: sw, borderOpacity: op, color: o.fill ? hexRgb(o.fill) : undefined, opacity: o.fill ? op : undefined });
      break;
    case 'line':
    case 'arrow': {
      const [[x1, y1], [x2, y2]] = o.pts.map(([x, y]) => [x * W, y * H]);
      page.drawLine({ start: { x: x1, y: H - y1 }, end: { x: x2, y: H - y2 }, thickness: sw, color, opacity: op, lineCap: LineCapStyle.Round });
      if (o.type === 'arrow') {
        const head = arrowHead(x1, y1, x2, y2, sw);
        page.drawSvgPath(`M${head[0].join(' ')} L${head[1].join(' ')} L${head[2].join(' ')} Z`, { x: 0, y: H, color, opacity: op });
      }
      break;
    }
    case 'ink':
      page.drawSvgPath(inkPath(o), { x: 0, y: H, borderColor: color, borderWidth: sw, borderOpacity: op, borderLineCap: LineCapStyle.Round });
      break;
    case 'highlight':
      page.drawRectangle({ x: 0, y: 0, width: W, height: H, color, opacity: 1, blendMode: BlendMode.Multiply });
      break;
    case 'underline':
    case 'strike': {
      const th = Math.max(0.8, H * 0.075);
      const y = o.type === 'underline' ? H * 0.04 + th / 2 : H - (H * 0.52);
      page.drawLine({ start: { x: 0, y }, end: { x: W, y }, thickness: th, color });
      break;
    }
    case 'whiteout':
      page.drawRectangle({ x: 0, y: 0, width: W, height: H, color: PL.rgb(1, 1, 1) });
      break;
    case 'stamp': {
      if (!cache.fonts.has('HelveticaBold')) cache.fonts.set('HelveticaBold', await doc.embedFont(PL.StandardFonts.HelveticaBold));
      const font = cache.fonts.get('HelveticaBold');
      const bw = Math.max(1.5, H * 0.07);
      page.drawRectangle({ x: bw / 2, y: bw / 2, width: W - bw, height: H - bw, borderColor: color, borderWidth: bw, color: PL.rgb(1, 1, 1), opacity: 0.25 });
      let size = H * 0.46;
      const tw = font.widthOfTextAtSize(o.text, size) * 1.04;
      if (tw > W - 2 * bw - 4) size *= (W - 2 * bw - 4) / tw;
      page.drawText(o.text, { x: (W - font.widthOfTextAtSize(o.text, size)) / 2, y: H / 2 - size * 0.35, size, font, color });
      break;
    }
  }
  page.pushOperators(popGraphicsState());
}
// OCR results become invisible text over the scanned image, so the page can be searched,
// selected and copied (like Acrobat's "searchable image")
async function drawOcrText(doc, page, words, pre, cache) {
  if (!cache.fonts.has('Helvetica')) cache.fonts.set('Helvetica', await doc.embedFont(PL.StandardFonts.Helvetica));
  const font = cache.fonts.get('Helvetica');
  page.pushOperators(PL.pushGraphicsState(), PL.setTextRenderingMode(PL.TextRenderingMode.Invisible));
  for (const wd of words) {
    let text = '';
    for (const ch of wd.t) { try { font.encodeText(ch); text += ch; } catch { text += '?'; } }
    if (!text.trim()) continue;
    const [ox, oy] = apply(pre, wd.o[0], wd.o[1]);
    const [ux, uy] = (() => { const v = [pre[0] * wd.u[0] + pre[2] * wd.u[1], pre[1] * wd.u[0] + pre[3] * wd.u[1]]; const l = Math.hypot(...v) || 1; return [v[0] / l, v[1] / l]; })();
    const natural = font.widthOfTextAtSize(text, wd.size) || 1;
    page.pushOperators(PL.setCharacterSqueeze(clamp((wd.w / natural) * 100, 10, 1000)));
    page.drawText(text, { x: ox, y: oy, size: wd.size, font, rotate: PL.degrees((Math.atan2(uy, ux) * 180) / Math.PI) });
  }
  page.pushOperators(PL.setCharacterSqueeze(100), PL.popGraphicsState());
}
// "Remove hidden information": document properties, XMP metadata, attachments, scripts
function sanitizeDoc(doc) {
  const N = n => PL.PDFName.of(n);
  doc.setTitle(''); doc.setAuthor(''); doc.setSubject(''); doc.setKeywords([]); doc.setCreator('');
  doc.catalog.delete(N('Metadata'));
  const names = doc.catalog.lookupMaybe(N('Names'), PL.PDFDict);
  if (names) { names.delete(N('EmbeddedFiles')); names.delete(N('JavaScript')); }
  doc.catalog.delete(N('OpenAction'));
  doc.catalog.delete(N('AA'));
  for (const p of doc.getPages()) p.node.delete(N('AA'));
}
function addNoteAnnotation(doc, page, o, pre) {
  const M = mul(pre, o.m);
  const pts = [[0, 0], [o.w, 0], [0, o.h], [o.w, o.h]].map(([x, y]) => apply(M, x, y));
  const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
  const c = hexRgb(o.color);
  const now = new Date();
  const annot = doc.context.obj({
    Type: 'Annot', Subtype: 'Text',
    Rect: [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)],
    Contents: PL.PDFHexString.fromText(o.text || ''),
    T: PL.PDFHexString.fromText('Nisar Office Studio'),
    M: PL.PDFString.fromDate(now),
    Name: 'Comment', C: [c.red, c.green, c.blue], F: 4, Open: false,
  });
  page.node.addAnnot(doc.context.register(annot));
}
// draw a page into a picture (used to remove redacted content for good, and for locked PDFs)
async function rasterizeInto(doc, ent, g, redactions) {
  const k = Math.min(3, Math.sqrt(32e6 / (g.w1 * g.h1)));
  const c = document.createElement('canvas');
  c.width = Math.ceil(g.w1 * k); c.height = Math.ceil(g.h1 * k);
  const ctx = c.getContext('2d', { alpha: false });
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, c.width, c.height);
  if (!ent.blank) await pageProxy(ent).render({ canvasContext: ctx, viewport: pageProxy(ent).getViewport({ scale: k, rotation: g.rotation }), transform: g.cropped ? g.shift(k) : null, intent: 'print', annotationMode: pdfjsLib.AnnotationMode.ENABLE_STORAGE }).promise;
  ctx.fillStyle = '#000';
  for (const o of redactions) {
    const M = mul(Sc(k), mul(g.V1, o.m));
    const pts = [[0, 0], [o.w, 0], [o.w, o.h], [0, o.h]].map(([x, y]) => apply(M, x, y));
    ctx.beginPath();
    pts.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y)));
    ctx.closePath();
    ctx.fill();
  }
  const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.92));
  const img = await doc.embedJpg(new Uint8Array(await blob.arrayBuffer()));
  const page = doc.addPage([g.w1, g.h1]);
  page.drawImage(img, { x: 0, y: 0, width: g.w1, height: g.h1 });
  c.width = 0;
  return page;
}
const isEncryptedError = err => /encrypt/i.test(String(err && err.message));
async function buildPdf({ only = null, keepEdits = false, forView = false } = {}) {
  const { PDFDocument, degrees } = PL;
  const list = only || S.entries;
  const loadLib = async bytes => PDFDocument.load(bytes, { updateMetadata: false });
  let out, fresh = !!only, locked = false;
  if (!fresh) {
    try { out = await loadLib(S.sources[0].bytes); } catch (err) { if (!isEncryptedError(err)) throw err; locked = true; fresh = true; }
  }
  if (fresh) out = await PDFDocument.create();
  const originals = fresh ? [] : out.getPages();
  if (!fresh) for (let i = out.getPageCount() - 1; i >= 0; i--) out.removePage(i);
  const libDocs = new Map();
  const libDoc = async key => {
    if (!libDocs.has(key)) {
      const srcIndex = key === 'dup0' ? 0 : key;
      try { libDocs.set(key, await loadLib(S.sources[srcIndex].bytes)); } catch (err) { if (!isEncryptedError(err)) throw err; libDocs.set(key, null); }
    }
    return libDocs.get(key);
  };
  const used = new Set();
  const cache = { fonts: new Map(), images: new Map() };
  const editedPages = [];
  const newFields = [];
  for (const ent of list) {
    const g = geom(ent);
    const redactions = ent.objects.filter(o => o.type === 'redact');
    let page, pre;
    if (redactions.length || (locked && !ent.blank) || (fresh && !ent.blank && !(await libDoc(ent.src === 0 ? 'dup0' : ent.src)))) {
      page = await rasterizeInto(out, ent, g, redactions);
      pre = mul([1, 0, 0, -1, 0, g.h1], g.V1); // original page space -> picture page space
    } else {
      if (ent.blank) page = out.addPage([ent.blank.w, ent.blank.h]);
      else if (!fresh && ent.src === 0 && !used.has(ent.index)) { used.add(ent.index); page = out.addPage(originals[ent.index]); }
      else {
        const d = await libDoc(ent.src === 0 ? 'dup0' : ent.src);
        const [cp] = await out.copyPages(d, [ent.index]);
        page = out.addPage(cp);
      }
      page.setRotation(degrees(g.rotation));
      if (g.cropped) {
        const [x0, y0, x1, y1] = g.view;
        page.setCropBox(x0, y0, x1 - x0, y1 - y0);
        page.setTrimBox(x0, y0, x1 - x0, y1 - y0);
      }
      pre = [1, 0, 0, 1, 0, 0];
    }
    if (ent.ocr && ent.ocr.length) await drawOcrText(out, page, ent.ocr, pre, cache);
    for (const o of ent.objects) if (o.type !== 'redact' && o.type !== 'field') await drawObject(out, page, o, pre, cache);
    for (const o of ent.objects) if (o.type === 'note') addNoteAnnotation(out, page, o, pre);
    for (const o of ent.objects) if (o.type === 'field') newFields.push({ o, page, pre });
    editedPages.push(page);
  }
  // fillable fields drawn with Prepare Form
  if (newFields.length) {
    const form = out.getForm();
    const taken = new Set(form.getFields().map(f => f.getName()));
    for (const { o, page, pre } of newFields) {
      const M = mul(pre, o.m);
      const pts = [[0, 0], [o.w, 0], [0, o.h], [o.w, o.h]].map(([x, y]) => apply(M, x, y));
      const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
      const rect = { x: Math.min(...xs), y: Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys), borderColor: PL.rgb(0.45, 0.55, 0.75), borderWidth: 0.75, backgroundColor: PL.rgb(1, 1, 1) };
      let name = (o.name || 'Field').replace(/[.\s]+/g, '_');
      for (let k = 2; taken.has(name); k++) name = `${(o.name || 'Field').replace(/[.\s]+/g, '_')}_${k}`;
      taken.add(name);
      try {
        if (o.kind === 'check') form.createCheckBox(name).addToPage(page, rect);
        else if (o.kind === 'dropdown') { const d = form.createDropdown(name); d.addOptions((o.options || []).filter(Boolean)); d.addToPage(page, rect); }
        else { const t = form.createTextField(name); if (o.multiline) t.enableMultiline(); t.addToPage(page, rect); }
      } catch (err) { console.warn('new field', err); }
    }
  }
  if (S.sanitize) sanitizeDoc(out);
  // form fields
  let formProblem = false;
  if (!fresh) {
    try {
      const form = out.getForm();
      for (const [name, value] of Object.entries(S.form)) {
        let f;
        try { f = form.getField(name); } catch { continue; }
        try {
          if (f instanceof PL.PDFTextField) f.setText(value == null ? '' : String(value));
          else if (f instanceof PL.PDFCheckBox) (value ? f.check() : f.uncheck());
          else if (f instanceof PL.PDFRadioGroup) {
            // pdf.js reports the widget's state name; when the field lists option labels
            // separately (/Opt), the state names are their positions ("0", "1", …)
            const opts = f.getOptions();
            const pick = opts.includes(value) ? value : /^\d+$/.test(String(value)) && opts[+value] !== undefined ? opts[+value] : null;
            if (value && value !== 'Off' && pick != null) f.select(pick); else if (!value || value === 'Off') f.clear();
          }
          else if (f instanceof PL.PDFDropdown) (value ? f.select(Array.isArray(value) ? value[0] : value) : f.clear());
          else if (f instanceof PL.PDFOptionList) f.select(Array.isArray(value) ? value : [value]);
        } catch (err) { console.warn('field', name, err); formProblem = true; }
      }
      if (S.flattenForms) form.flatten();
    } catch (err) { console.warn('form', err); }
  }
  const now = new Date(Math.floor(Date.now() / 1000) * 1000);
  out.setModificationDate(now);
  out.setProducer('Nisar Office Studio');
  if (!out.getTitle() && !S.sanitize) out.setTitle(S.name.replace(/\.pdf$/i, ''), { showInWindowTitleBar: true });
  if (keepEdits && !fresh && !forView && !S.sanitize) {
    const hasEdits = S.entries.some(e => e.objects.length || (e.ocr && e.ocr.length) || e.crop);
    const total = S.sources.reduce((n, s) => n + s.bytes.length, 0);
    if (hasEdits && total < 60e6) {
      const data = { v: 1, mod: now.toISOString(), sources: S.sources.map(s => ({ name: s.name, b64: s.bytes.toBase64() })), entries: S.entries, form: S.form, flatten: S.flattenForms };
      await out.attach(new TextEncoder().encode(JSON.stringify(data)), EDIT_FILE, { mimeType: 'application/json', description: 'Nisar Office Studio editing data (keeps your added items editable)', creationDate: now, modificationDate: now });
    }
  }
  const bytes = await out.save({ updateMetadata: false, useObjectStreams: true });
  return Object.assign(bytes, { locked, formProblem });
}
function friendly(err) {
  const m = String(err && err.message ? err.message : err);
  if (isEncryptedError(err)) return 'This PDF is protected against changes.';
  return m;
}

// ---------------------------------------------------------------- file commands
function updateTitle() {
  document.title = S.open ? `${S.dirty ? '• ' : ''}${S.name} - Nisar Office Studio` : 'PDF - Nisar Office Studio';
  $('#docname').textContent = S.open ? `${S.name}${S.dirty ? ' • Not saved' : ''}` : '';
  $('#docname').title = S.path || '';
}
function markDirty() { if (!S.dirty) { S.dirty = true; updateTitle(); } }
let busyToast = null;
function setBusy(msg) {
  document.body.style.cursor = msg ? 'progress' : '';
  clearTimeout(busyToast);
  if (msg) busyToast = setTimeout(() => toast(msg), 350); else $('#toast').hidden = true;
}
async function save(saveAs = false) {
  if (!S.open) return false;
  finishEditing();
  closeNote(true);
  try {
    if (!(await ensureEditable())) return false;
    setBusy('Saving…');
    let bytes = await buildPdf({ keepEdits: true });
    const formProblem = bytes.formProblem;
    if (S.protect) {
      const enc = await api.qpdf({ op: 'encrypt', bytes, ...S.protect });
      if (enc.error) throw new Error(`Couldn't add the password: ${enc.error}`);
      bytes = enc.bytes;
    }
    const r = await api.savePdf({ path: S.path, name: S.name, bytes }, saveAs || !S.path);
    if (!r || r.canceled) return false;
    if (r.error) { await alertBox("Couldn't save", r.error); return false; }
    S.path = r.path;
    S.name = r.name;
    S.dirty = false;
    api.pdfRecoveryClear();
    recoveryWritten.clear();
    updateTitle();
    if (formProblem) bytes.formProblem = true;
    $('#st-saved').textContent = `Saved ${new Date().toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}`;
    toast(bytes.formProblem ? 'Saved. Some form text uses characters the form can’t display.' : `Saved to ${r.path}`, () => api.showInFolder(r.path), 'Show in folder');
    return true;
  } catch (err) {
    console.error(err);
    await alertBox("Couldn't save", friendly(err));
    return false;
  } finally { setBusy(null); }
}
// A PDF whose author restricted changes can only be saved with its permissions password
// (the same rule Acrobat follows). With it, qpdf removes the restriction for this save.
async function ensureEditable() {
  for (const src of S.sources) {
    let locked = false;
    try { await PL.PDFDocument.load(src.bytes, { updateMetadata: false }); } catch (err) { if (isEncryptedError(err)) locked = true; else throw err; }
    if (!locked) continue;
    for (let attempt = 0; attempt < 3; attempt++) {
      const r = await askForm('Permissions Password', [{ id: 'pw', label: 'Permissions password', type: 'password' }], 'Unlock',
        attempt ? 'That password is not correct. Please try again.' : `The author of "${src.name}" has protected it against changes. Enter its permissions password to save your changes.`);
      if (!r) return false;
      const dec = await api.qpdf({ op: 'decrypt', bytes: src.bytes, password: r.pw });
      if (dec.bytes) {
        src.bytes = dec.bytes;
        if (src === S.sources[0] && !S.protect) S.protect = { user: '', owner: r.pw, allowPrint: true, allowCopy: false, allowModify: false };
        toast('Unlocked. The file stays protected with the same permissions password; use Protect › Password Protect to change what is allowed.');
        break;
      }
      if (!dec.badPassword) { await alertBox('Permissions Password', dec.error || 'This PDF could not be unlocked.'); return false; }
      if (attempt === 2) return false;
    }
  }
  return true;
}
// PDFs that needed a password to open are decrypted with it (so editing keeps real text)
async function decryptWithPassword(bytes, loaded) {
  if (!loaded.password) return bytes;
  const dec = await api.qpdf({ op: 'decrypt', bytes, password: loaded.password });
  return dec.bytes || bytes;
}
async function openDialog() {
  const p = await api.openDialog('all');
  if (p) await openPath(p);
}
async function openPath(p) {
  hideBackstage();
  if (!/\.pdf$/i.test(p)) { api.newWindow({ file: p }); return; }
  if (S.open) { api.newWindow({ file: p }); return; }
  const r = await api.readPdf(p);
  if (r.error) { await alertBox("Couldn't open", r.error); return; }
  try { await openBytes(r.bytes, r.name, r.path); } catch (err) {
    if (err && /destroy/i.test(String(err.message))) return; // password prompt cancelled
    console.error(err);
    await alertBox("Couldn't open", `"${r.name}" couldn't be opened. It may be damaged or not a real PDF.`);
  }
}
async function printPdf() {
  if (!S.open) return;
  finishEditing();
  setBusy('Preparing to print…');
  try {
    const bytes = await buildPdf({ forView: true });
    await api.previewPdfBytes({ bytes, name: S.name });
    toast('Use the printer button at the top right of the preview to print.');
  } catch (err) { await alertBox('Print', friendly(err)); } finally { setBusy(null); }
}
// rebuilt PDF (with everything added) opened in pdf.js, for images and printing
async function flattenedDoc() {
  const bytes = await buildPdf({ forView: true });
  return pdfjsLib.getDocument({ data: bytes, ...PDF_OPTS }).promise;
}
async function exportImages(which) {
  if (!S.open) return;
  finishEditing();
  const idx = which === 'all' ? S.entries.map((e, i) => i) : which === 'selected' ? targetIndexes() : [S.current];
  setBusy('Creating images…');
  try {
    const doc = await flattenedDoc();
    const base = S.name.replace(/\.pdf$/i, '');
    const files = [];
    for (const i of idx) {
      const page = await doc.getPage(i + 1);
      const vp = page.getViewport({ scale: 2 });
      const c = document.createElement('canvas');
      c.width = Math.ceil(vp.width); c.height = Math.ceil(vp.height);
      const ctx = c.getContext('2d', { alpha: false });
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: ctx, viewport: vp, intent: 'print' }).promise;
      const blob = await new Promise(r => c.toBlob(r, 'image/png'));
      files.push({ name: `${base} - page ${i + 1}.png`, bytes: new Uint8Array(await blob.arrayBuffer()) });
      c.width = 0;
    }
    doc.destroy();
    const r = await api.saveImages({ files });
    if (r && r.error) await alertBox('Export Images', r.error);
    else if (r && !r.canceled) toast(`Saved ${r.count} image${r.count === 1 ? '' : 's'}.`, () => api.showInFolder(r.path), 'Show in folder');
  } catch (err) { await alertBox('Export Images', friendly(err)); } finally { setBusy(null); }
}
// PDF -> editable document: rebuild lines and paragraphs from the positioned text
async function exportToWord() {
  if (!S.open) return;
  setBusy('Converting to a document…');
  try {
    const pages = [];
    for (const ent of S.entries) {
      if (ent.blank) { pages.push(''); continue; }
      const tc = await textContent(ent);
      const items = tc.items.filter(it => typeof it.str === 'string' && it.str.length);
      const g = geom(ent);
      const runs = items.map(it => {
        const [x, y] = apply(g.V1, it.transform[4], it.transform[5]);
        const size = Math.hypot(it.transform[2], it.transform[3]) || it.height || 10;
        return { x, y, size, w: it.width, str: it.str, bold: /bold|black|heavy/i.test((tc.styles[it.fontName] || {}).fontFamily || it.fontName) };
      }).filter(r => r.str.trim() || r.str === ' ');
      runs.sort((a, b) => a.y - b.y || a.x - b.x);
      const lines = [];
      for (const r of runs) {
        const L = lines.find(l => Math.abs(l.y - r.y) < Math.max(l.size, r.size) * 0.45);
        if (L) L.runs.push(r); else lines.push({ y: r.y, size: r.size, runs: [r] });
      }
      lines.sort((a, b) => a.y - b.y);
      const sizes = lines.map(l => l.size).sort((a, b) => a - b);
      const body = sizes[Math.floor(sizes.length / 2)] || 11;
      const paras = [];
      let cur = null, prev = null;
      for (const l of lines) {
        l.runs.sort((a, b) => a.x - b.x);
        let text = '';
        let lastEnd = null;
        for (const r of l.runs) {
          if (lastEnd != null && r.x - lastEnd > r.size * 0.2 && !/\s$/.test(text) && !/^\s/.test(r.str)) text += ' ';
          text += esc(r.str);
          lastEnd = r.x + r.w;
        }
        text = text.replace(/\s+/g, ' ').trim();
        if (!text) continue;
        const size = Math.max(...l.runs.map(r => r.size));
        const bold = l.runs.every(r => r.bold);
        const heading = size > body * 1.28 ? (size > body * 1.8 ? 'h1' : 'h2') : null;
        const gap = prev ? l.y - prev.y : 0;
        const newPara = !cur || heading || cur.heading || gap > Math.max(size, prev.size) * 1.75 || Math.abs(size - prev.size) > 1.5 || /[.!?:]$/.test(cur.text) && gap > size * 1.45;
        if (newPara) { cur = { text, heading, bold, size }; paras.push(cur); }
        else { cur.text += (cur.text.endsWith('-') ? '' : ' ') + text; cur.bold = cur.bold && bold; }
        prev = l;
      }
      pages.push(paras.map(p => p.heading ? `<${p.heading}>${p.text}</${p.heading}>` : `<p>${p.bold ? `<b>${p.text}</b>` : p.text}</p>`).join(''));
    }
    const html = pages.map(p => p || '<p><br></p>').join('<div class="page-break" contenteditable="false"></div>');
    if (!html.replace(/<[^>]+>/g, '').trim()) { await alertBox('Export to Word', 'This PDF has no text that can be copied (it may be a scan). Nothing was converted.'); return; }
    await api.newWindow({ doc: { html, name: `${S.name.replace(/\.pdf$/i, '')}.docx` } });
    toast('Opened as a new document. Save it to keep it.');
  } catch (err) { await alertBox('Export to Word', friendly(err)); } finally { setBusy(null); }
}
async function confirmDiscard() {
  finishEditing();
  closeNote(true);
  if (!S.dirty) return true;
  const c = await api.confirmUnsaved(S.name);
  if (c === 0) return save(false);
  return c === 1;
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
function hideBackstage() { $('#backstage').hidden = true; }
async function refreshRecent() {
  const list = await api.getRecent();
  const ul = $('#recent');
  if (!list.length) { ul.innerHTML = '<li class="empty">Files you open or save will appear here.</li>'; return; }
  ul.innerHTML = list.map((p, i) => `<li data-i="${i}">${svg(/\.pdf$/i.test(p) ? 'pdf' : 'doc')}<div class="meta"><div class="nm">${esc(p.split(/[\\/]/).pop())}</div><div class="pt">${esc(p)}</div></div><button class="x" data-rm="${i}" title="Remove from list">${svg('close')}</button></li>`).join('');
  ul.onclick = ev => {
    const rm = ev.target.closest('[data-rm]');
    if (rm) { api.removeRecent(list[+rm.dataset.rm]); refreshRecent(); return; }
    const li = ev.target.closest('li[data-i]');
    if (li) openPath(list[+li.dataset.i]);
  };
}

// ---------------------------------------------------------------- page content: header & footer, watermark
const textWidth = (text, size, font = 'Helvetica', bold = false) => {
  const c = textWidth.c || (textWidth.c = document.createElement('canvas').getContext('2d'));
  c.font = `${bold ? '700 ' : ''}${size}px ${CSS_FONT[font]}`;
  return c.measureText(text).width;
};
function pagesFor(scope) {
  if (scope === 'selected') { const idx = targetIndexes(); return idx.map(i => S.entries[i]); }
  if (scope === 'current') return [S.entries[S.current]].filter(Boolean);
  return S.entries;
}
async function headerFooterDialog() {
  const prev = S.hf || {};
  const r = await askForm('Header & Footer', [
    { id: 'hl', label: 'Header left', value: prev.hl || '' }, { id: 'hc', label: 'Header centre', value: prev.hc || '' }, { id: 'hr', label: 'Header right', value: prev.hr || '' },
    { id: 'fl', label: 'Footer left', value: prev.fl || '' }, { id: 'fc', label: 'Footer centre', value: prev.fc ?? 'Page {page} of {pages}' }, { id: 'fr', label: 'Footer right', value: prev.fr || '' },
    { id: 'size', label: 'Font size (pt)', type: 'number', value: prev.size || 10, attrs: 'min="5" max="48"' },
    { id: 'margin', label: 'Distance from edge (pt)', type: 'number', value: prev.margin || 30, attrs: 'min="6" max="144"' },
    { id: 'start', label: 'First page number', type: 'number', value: prev.start || 1, attrs: 'min="0"' },
    { id: 'bates', label: 'Bates prefix (used by {bates})', value: prev.bates || '' },
    { id: 'scope', label: 'Pages', type: 'select', value: prev.scope || 'all', options: [{ value: 'all', label: 'All pages' }, { value: 'selected', label: 'Selected pages' }] },
  ], 'Apply', 'Codes you can use: {page}, {pages}, {date}, {bates} (e.g. ABC000123). Leave every box empty to remove headers and footers.', true);
  if (!r) return;
  S.hf = r;
  const size = clamp(parseFloat(r.size) || 10, 5, 48), margin = clamp(parseFloat(r.margin) || 30, 6, 144);
  const start = parseInt(r.start, 10) || 1;
  const targets = new Set(pagesFor(r.scope).map(e => e.id));
  const date = new Date().toLocaleDateString();
  let made = 0;
  S.entries.forEach((ent, i) => {
    ent.objects = ent.objects.filter(o => !o.hf);
    if (!targets.has(ent.id)) return;
    const v = views.get(ent.id);
    const g = geom(ent);
    const n = i + start;
    const fill = t => t.replace(/\{page\}/g, n).replace(/\{pages\}/g, S.entries.length + start - 1).replace(/\{date\}/g, date).replace(/\{bates\}/g, `${r.bates}${String(n).padStart(6, '0')}`);
    const lh = size * 1.2 + 2 * TEXT_PAD;
    for (const [key, align, top] of [['hl', 'left', true], ['hc', 'center', true], ['hr', 'right', true], ['fl', 'left', false], ['fc', 'center', false], ['fr', 'right', false]]) {
      if (!r[key].trim()) continue;
      const o = { id: uid(), type: 'text', hf: true, m: mul(inv(g.V1), T(margin, top ? margin - TEXT_PAD : g.h1 - margin - lh + TEXT_PAD)), w: g.w1 - 2 * margin, h: lh, text: fill(r[key]), font: 'Helvetica', size, bold: false, italic: false, color: '#333333', align, autoW: false };
      ent.objects.push(o);
      made++;
    }
    if (v) renderObjects(v);
  });
  commit();
  toast(made ? `Added to ${targets.size} page${targets.size === 1 ? '' : 's'}. Each one is editable text.` : 'Headers and footers removed.');
}
async function watermarkDialog() {
  const prev = S.wm || {};
  const r = await askForm('Watermark', [
    { id: 'text', label: 'Text (empty removes the watermark)', value: prev.text ?? 'CONFIDENTIAL' },
    { id: 'size', label: 'Font size (pt)', type: 'number', value: prev.size || 72, attrs: 'min="8" max="300"' },
    { id: 'color', label: 'Colour', type: 'color', value: prev.color || '#c4314b' },
    { id: 'opacity', label: 'Opacity', type: 'select', value: String(prev.opacity || 0.25), options: [['0.15', '15 %'], ['0.25', '25 %'], ['0.4', '40 %'], ['0.6', '60 %']].map(([value, label]) => ({ value, label })) },
    { id: 'angle', label: 'Angle', type: 'select', value: String(prev.angle ?? 45), options: [['45', 'Diagonal'], ['0', 'Horizontal'], ['90', 'Vertical']].map(([value, label]) => ({ value, label })) },
    { id: 'scope', label: 'Pages', type: 'select', value: 'all', options: [{ value: 'all', label: 'All pages' }, { value: 'selected', label: 'Selected pages' }] },
  ], 'Apply', null, true);
  if (!r) return;
  S.wm = r;
  const targets = new Set(pagesFor(r.scope).map(e => e.id));
  const size = clamp(parseFloat(r.size) || 72, 8, 300);
  const text = r.text.trim();
  const ang = (parseFloat(r.angle) || 0) * Math.PI / 180;
  for (const ent of S.entries) {
    ent.objects = ent.objects.filter(o => !o.wm);
    if (!text || !targets.has(ent.id)) continue;
    const g = geom(ent);
    const w = textWidth(text, size, 'Helvetica', true) + 2 * TEXT_PAD + 2, h = size * 1.2 + 2 * TEXT_PAD;
    // centred, rotated counter-clockwise on screen
    const R = [Math.cos(-ang), Math.sin(-ang), -Math.sin(-ang), Math.cos(-ang), 0, 0];
    const m = mul(inv(g.V1), mul(T(g.w1 / 2, g.h1 / 2), mul(R, T(-w / 2, -h / 2))));
    ent.objects.unshift({ id: uid(), type: 'text', wm: true, m, w, h, text, font: 'Helvetica', size, bold: true, italic: false, color: r.color, align: 'center', autoW: false, opacity: parseFloat(r.opacity) || 0.25 });
    const v = views.get(ent.id);
    if (v) renderObjects(v);
  }
  for (const v of views.values()) renderObjects(v);
  commit();
}

// ---------------------------------------------------------------- crop
async function cropDialog() {
  const idx = targetIndexes();
  if (!idx.length) return;
  const first = S.entries[idx[0]];
  const g = geom(first);
  const unit = /^en-US/.test(navigator.language) ? 'in' : 'mm';
  const k = unit === 'in' ? 72 : 72 / 25.4;
  const cur = first.crop ? displayMargins(first) : { top: 0, bottom: 0, left: 0, right: 0 };
  const f = (id, label) => ({ id, label: `${label} (${unit})`, type: 'number', value: Math.round((cur[id] / k) * 100) / 100, attrs: 'step="0.1" min="0"' });
  const r = await askForm('Crop Pages', [f('top', 'Top'), f('bottom', 'Bottom'), f('left', 'Left'), f('right', 'Right'),
    { id: 'scope', label: 'Apply to', type: 'select', value: idx.length > 1 ? 'selected' : 'current', options: [{ value: 'current', label: 'This page' }, { value: 'selected', label: 'Selected pages' }, { value: 'all', label: 'All pages' }] }],
  'Crop', `Trims the edges of the page (the page is ${Math.round((g.w1 / k) * 10) / 10} × ${Math.round((g.h1 / k) * 10) / 10} ${unit}). Use 0 everywhere to undo a crop.`, true);
  if (!r) return;
  const m = { top: (parseFloat(r.top) || 0) * k, bottom: (parseFloat(r.bottom) || 0) * k, left: (parseFloat(r.left) || 0) * k, right: (parseFloat(r.right) || 0) * k };
  let done = 0;
  for (const ent of pagesFor(r.scope === 'current' ? 'current' : r.scope)) {
    const ok = setCropFromDisplay(ent, m);
    if (ok) done++;
  }
  if (!done) { await alertBox('Crop Pages', 'Those margins would leave nothing of the page.'); return; }
  afterStructureChange(S.current);
}
// margins seen on screen (top/right/bottom/left of the displayed page) -> PDF-space crop
function setCropFromDisplay(ent, m) {
  if (ent.blank) return false;
  const saved = ent.crop;
  ent.crop = null;
  const g = geom(ent);
  if (m.left + m.right >= g.w1 - 10 || m.top + m.bottom >= g.h1 - 10) { ent.crop = saved; return false; }
  if (!m.left && !m.right && !m.top && !m.bottom) return true;
  const Vi = inv(g.V1);
  const pts = [[m.left, m.top], [g.w1 - m.right, m.top], [m.left, g.h1 - m.bottom], [g.w1 - m.right, g.h1 - m.bottom]].map(([x, y]) => apply(Vi, x, y));
  const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
  const [x0, y0, x1, y1] = g.full;
  ent.crop = { l: Math.min(...xs) - x0, b: Math.min(...ys) - y0, r: x1 - Math.max(...xs), t: y1 - Math.max(...ys) };
  return true;
}
function displayMargins(ent) {
  const g = geom(ent);
  const full = { ...ent, crop: null };
  const gf = geom(full);
  const Vf = gf.V1;
  const pts = [[g.view[0], g.view[1]], [g.view[2], g.view[3]]].map(([x, y]) => apply(Vf, x, y));
  const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
  return { left: Math.min(...xs), top: Math.min(...ys), right: gf.w1 - Math.max(...xs), bottom: gf.h1 - Math.max(...ys) };
}

// ---------------------------------------------------------------- split, combine, pictures to pages
async function splitDialog() {
  const n = S.entries.length;
  const r = await askForm('Split PDF', [
    { id: 'mode', label: 'Split into', type: 'select', value: 'every', options: [{ value: 'every', label: 'Files of N pages each' }, { value: 'single', label: 'One file per page' }, { value: 'ranges', label: 'Page ranges' }] },
    { id: 'every', label: 'Pages per file (N)', type: 'number', value: Math.max(1, Math.ceil(n / 2)), attrs: `min="1" max="${n}"` },
    { id: 'ranges', label: 'Ranges (for example 1-3, 4-7, 8)', value: '', placeholder: `1-${Math.ceil(n / 2)}, ${Math.ceil(n / 2) + 1}-${n}` },
  ], 'Split', `This PDF has ${n} pages. The new files are saved into a folder you choose; this document is not changed.`);
  if (!r) return;
  const groups = [];
  if (r.mode === 'single') for (let i = 0; i < n; i++) groups.push([i]);
  else if (r.mode === 'every') { const k = clamp(parseInt(r.every, 10) || 1, 1, n); for (let i = 0; i < n; i += k) groups.push(Array.from({ length: Math.min(k, n - i) }, (_, j) => i + j)); }
  else {
    for (const part of r.ranges.split(/[,;]+/).map(s => s.trim()).filter(Boolean)) {
      const m = part.match(/^(\d+)\s*(?:-\s*(\d+))?$/);
      if (!m) { await alertBox('Split PDF', `"${part}" is not a page range.`); return; }
      const a = clamp(+m[1], 1, n), b = clamp(+(m[2] || m[1]), 1, n);
      groups.push(Array.from({ length: Math.abs(b - a) + 1 }, (_, j) => Math.min(a, b) - 1 + j));
    }
  }
  if (!groups.length) return;
  setBusy('Splitting…');
  try {
    const base = S.name.replace(/\.pdf$/i, '');
    const files = [];
    for (const gIdx of groups) files.push({ name: `${base} (pages ${pageRangeLabel(gIdx)}).pdf`, bytes: await buildPdf({ only: gIdx.map(i => S.entries[i]) }) });
    const res = await api.saveFiles({ title: 'Choose a folder for the split files', files });
    if (res && res.error) await alertBox('Split PDF', res.error);
    else if (res && !res.canceled) toast(`Saved ${res.count} PDF files.`, () => api.showInFolder(res.path), 'Show in folder');
  } catch (err) { await alertBox('Split PDF', friendly(err)); } finally { setBusy(null); }
}
async function combineFiles() {
  const paths = await api.openMany('pdfs');
  if (!paths.length) return;
  let after = S.entries.length - 1;
  let added = 0;
  for (const p of paths) {
    const r = await api.readFileBytes(p);
    if (r.error) { await alertBox('Combine Files', r.error); continue; }
    const before = S.entries.length;
    await insertPdfBytes(r, after);
    added += S.entries.length - before;
    after = S.entries.length - 1;
  }
  if (added) toast(`Added ${added} page${added === 1 ? '' : 's'} from ${paths.length} file${paths.length === 1 ? '' : 's'}.`);
}
async function imagesToPages() {
  const paths = await api.openMany('images');
  if (!paths.length) return;
  let at = S.current + 1;
  for (const p of paths) {
    const r = await api.readImage(p);
    if (r.error) { toast(r.error); continue; }
    const { w: iw, h: ih } = await imageSize(r.url);
    const portrait = ih >= iw;
    const a4 = !/^en-(US|CA)/.test(navigator.language);
    const [pw0, ph0] = a4 ? [595.28, 841.89] : [612, 792];
    const pw = portrait ? pw0 : ph0, ph = portrait ? ph0 : pw0;
    const fit = Math.min((pw - 36) / iw, (ph - 36) / ih, 1.5);
    const w = iw * fit, h = ih * fit;
    const ent = { id: uid(), src: 0, index: 0, rot: 0, blank: { w: pw, h: ph }, objects: [] };
    ent.objects.push({ id: uid(), type: 'image', m: mul(inv(viewportMatrix([0, 0, pw, ph], 0, 1)), T((pw - w) / 2, (ph - h) / 2)), w, h, src: r.url, kind: 'image' });
    S.entries.splice(at++, 0, ent);
  }
  afterStructureChange(S.current + 1);
}

// ---------------------------------------------------------------- find & redact, protect, hidden info, size
function itemBoxes(item, start, end, g) {
  // approximate the PDF-space box of characters [start, end) of a text item
  const [a, b, c, d, e, f] = item.transform;
  const len = item.str.length || 1;
  const fs = Math.hypot(c, d) || item.height || 10;
  const ux = a / (Math.hypot(a, b) || 1), uy = b / (Math.hypot(a, b) || 1);
  const x0 = (item.width * start) / len, x1 = (item.width * end) / len;
  const corner = (along, up) => [e + ux * along - uy * up, f + uy * along + ux * up];
  const pts = [corner(x0, -0.25 * fs), corner(x1, -0.25 * fs), corner(x0, 0.95 * fs), corner(x1, 0.95 * fs)].map(([x, y]) => apply(g.V1, x, y));
  const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
  return { x: Math.min(...xs) - 1, y: Math.min(...ys) - 1, w: Math.max(...xs) - Math.min(...xs) + 2, h: Math.max(...ys) - Math.min(...ys) + 2 };
}
async function findRedactDialog() {
  const r = await askForm('Find & Redact', [
    { id: 'q', label: 'Words, names or numbers to remove (separate several with |)', value: getSelection().toString().trim() },
    { id: 'case', label: 'Match case', type: 'checkbox', value: false },
  ], 'Find', 'Every match is covered with a black box and removed for good when you save.');
  if (!r || !r.q.trim()) return;
  const terms = r.q.split('|').map(s => s.trim()).filter(Boolean);
  let count = 0;
  for (const ent of S.entries) {
    if (ent.blank) continue;
    const { str, items, starts } = await pageString(ent);
    const g = geom(ent);
    const v = views.get(ent.id);
    const hay = r.case ? str : str.toLocaleLowerCase();
    for (const term of terms) {
      const needle = r.case ? term : term.toLocaleLowerCase();
      let i = 0;
      while ((i = hay.indexOf(needle, i)) !== -1) {
        const end = i + needle.length;
        for (let k = 0; k < items.length; k++) {
          const s0 = starts[k], s1 = s0 + items[k].str.length;
          if (s1 <= i || s0 >= end) continue;
          const b = itemBoxes(items[k], Math.max(i, s0) - s0, Math.min(end, s1) - s0, g);
          ent.objects.push({ id: uid(), type: 'redact', m: mul(inv(g.V1), T(b.x, b.y)), w: b.w, h: b.h });
        }
        count++;
        i = end;
      }
    }
    if (v) renderObjects(v);
  }
  if (!count) { toast('No matches found.'); return; }
  commit();
  toast(`Marked ${count} match${count === 1 ? '' : 'es'} for redaction. They are removed for good when you save.`);
}
async function protectDialog() {
  const p = S.protect || {};
  const r = await askForm('Password Protect', [
    { id: 'user', label: 'Password to open the PDF (optional)', type: 'password', value: p.user || '' },
    { id: 'owner', label: 'Permissions password (needed to change the settings below)', type: 'password', value: p.owner || '' },
    { id: 'print', label: 'Allow printing', type: 'checkbox', value: p.allowPrint !== false },
    { id: 'copy', label: 'Allow copying text and pictures', type: 'checkbox', value: p.allowCopy !== false },
    { id: 'modify', label: 'Allow changes', type: 'checkbox', value: p.allowModify !== false },
  ], 'Apply', 'Uses 256-bit AES encryption when you save. Keep your passwords safe: a lost open password cannot be recovered.');
  if (!r) return;
  if (!r.user && !r.owner) { await alertBox('Password Protect', 'Enter at least one password.'); return; }
  if (!r.owner && (!r.print || !r.copy || !r.modify)) { await alertBox('Password Protect', 'Limiting printing, copying or changes needs a permissions password.'); return; }
  S.protect = { user: r.user, owner: r.owner || r.user, allowPrint: r.print, allowCopy: r.copy, allowModify: r.modify };
  markDirty();
  toast('The PDF will be password-protected when you save it.');
}
function unprotect() {
  if (!S.protect) { toast('This PDF has no password.'); return; }
  S.protect = null;
  markDirty();
  toast('The password will be removed when you save.');
}
function sanitizeToggle() {
  S.sanitize = !S.sanitize;
  markDirty();
  updateToggles();
  toast(S.sanitize ? 'Document properties, metadata, attachments and scripts will be removed when you save.' : 'Hidden information will be kept.');
}
async function compressCopy() {
  const r = await askForm('Reduce File Size', [
    { id: 'q', label: 'Picture quality', type: 'select', value: 'medium', options: [{ value: 'high', label: 'High (print quality)' }, { value: 'medium', label: 'Medium (recommended)' }, { value: 'low', label: 'Low (smallest file)' }] },
  ], 'Save Smaller Copy', 'Large photos are scaled down and the file is repacked. This document stays as it is; a new copy is saved.');
  if (!r) return;
  const preset = { high: [2400, 0.85], medium: [1600, 0.72], low: [1100, 0.55] }[r.q];
  setBusy('Making the file smaller…');
  try {
    const before = await buildPdf({});
    let bytes = await downsampleImages(before, preset[0], preset[1]);
    const opt = await api.qpdf({ op: 'optimize', bytes });
    if (opt.bytes && opt.bytes.length < bytes.length) bytes = opt.bytes;
    const original = S.sources[0].bytes.length;
    const mb = n => (n < 1048576 ? `${Math.max(1, Math.round(n / 1024))} KB` : `${(n / 1048576).toFixed(1)} MB`);
    if (bytes.length >= original * 0.98) { setBusy(null); await alertBox('Reduce File Size', `This PDF is already as small as it can be made (${mb(original)}). No copy was saved.`); return; }
    const res = await api.savePdf({ bytes, name: `${S.name.replace(/\.pdf$/i, '')} (smaller)`, title: 'Save Smaller Copy', copy: true }, true);
    if (res && res.path) toast(`Saved: ${mb(original)} → ${mb(bytes.length)}.`, () => api.showInFolder(res.path), 'Show in folder');
    else if (res && res.error) await alertBox("Couldn't save", res.error);
  } catch (err) { await alertBox('Reduce File Size', friendly(err)); } finally { setBusy(null); }
}
// re-encode big JPEG photos at a lower resolution
async function downsampleImages(bytes, maxSide, quality) {
  const doc = await PL.PDFDocument.load(bytes, { updateMetadata: false });
  const N = n => PL.PDFName.of(n);
  let changed = 0;
  for (const [ref, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PL.PDFRawStream)) continue;
    const dict = obj.dict;
    if (String(dict.get(N('Subtype'))) !== '/Image' || String(dict.get(N('Filter'))) !== '/DCTDecode') continue;
    const cs = dict.get(N('ColorSpace'));
    const csName = String(cs instanceof PL.PDFRef ? doc.context.lookup(cs) : cs);
    if (!/DeviceRGB|DeviceGray|ICCBased/.test(csName) || dict.get(N('Decode'))) continue;
    const w = dict.lookup(N('Width')).asNumber(), h = dict.lookup(N('Height')).asNumber();
    if (Math.max(w, h) <= maxSide) continue;
    try {
      const bmp = await createImageBitmap(new Blob([obj.contents], { type: 'image/jpeg' }));
      const k = maxSide / Math.max(w, h);
      const c = document.createElement('canvas');
      c.width = Math.max(1, Math.round(w * k)); c.height = Math.max(1, Math.round(h * k));
      c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
      bmp.close();
      const gray = /DeviceGray/.test(csName);
      const blob = await new Promise(res => c.toBlob(res, 'image/jpeg', quality));
      const out = new Uint8Array(await blob.arrayBuffer());
      c.width = 0;
      if (out.length >= obj.contents.length) continue;
      dict.set(N('Width'), PL.PDFNumber.of(Math.round(w * k)));
      dict.set(N('Height'), PL.PDFNumber.of(Math.round(h * k)));
      dict.set(N('ColorSpace'), N(gray ? 'DeviceRGB' : 'DeviceRGB'));
      dict.set(N('BitsPerComponent'), PL.PDFNumber.of(8));
      dict.delete(N('DecodeParms'));
      doc.context.assign(ref, PL.PDFRawStream.of(dict, out));
      changed++;
    } catch { /* leave this image as it is */ }
  }
  return changed ? doc.save({ updateMetadata: false, useObjectStreams: true }) : bytes;
}

// ---------------------------------------------------------------- OCR (tesseract.js, runs on this PC)
let ocrWorker = null;
async function getOcrWorker(onProgress) {
  if (ocrWorker) return ocrWorker;
  const mod = await import('../node_modules/tesseract.js/dist/tesseract.esm.min.js');
  const Tess = mod.createWorker ? mod : mod.default;
  const base = new URL('../node_modules/', import.meta.url).href;
  ocrWorker = await Tess.createWorker('eng', 1, {
    workerPath: `${base}tesseract.js/dist/worker.min.js`,
    corePath: `${base}tesseract.js-core/`,
    langPath: `${base}@tesseract.js-data/eng/4.0.0_best_int`,
    gzip: true, workerBlobURL: false, cacheMethod: 'none',
    logger: m => onProgress && onProgress(m),
  });
  return ocrWorker;
}
function ocrProgressUi() {
  let el = $('.ocr-progress');
  if (!el) { el = document.createElement('div'); el.className = 'ocr-progress'; document.body.appendChild(el); }
  el.innerHTML = '<div class="row"><span class="msg">Preparing text recognition…</span><button class="btn" data-ocr-cancel>Cancel</button></div><div class="bar"><i></i></div>';
  return el;
}
async function recognizeText(scope) {
  const list = pagesFor(scope).filter(e => !e.blank || e.objects.some(o => o.type === 'image'));
  if (!list.length) return;
  const ui = ocrProgressUi();
  let cancelled = false;
  ui.querySelector('[data-ocr-cancel]').onclick = () => { cancelled = true; };
  const msg = t => { ui.querySelector('.msg').textContent = t; };
  const bar = f => { ui.querySelector('.bar i').style.width = `${Math.round(f * 100)}%`; };
  let pageNo = 0, words = 0;
  try {
    const worker = await getOcrWorker(m => { if (m.status === 'recognizing text' && typeof m.progress === 'number') bar((pageNo - 1 + m.progress) / list.length); });
    for (const ent of list) {
      if (cancelled) break;
      pageNo++;
      msg(`Recognizing text on page ${S.entries.indexOf(ent) + 1} (${pageNo} of ${list.length})…`);
      const g = geom(ent);
      // render the page (with added pictures) at about 300 dpi
      const k = Math.min(300 / 72, Math.sqrt(30e6 / (g.w1 * g.h1)));
      const c = document.createElement('canvas');
      c.width = Math.ceil(g.w1 * k); c.height = Math.ceil(g.h1 * k);
      const ctx = c.getContext('2d', { alpha: false });
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      if (!ent.blank) await pageProxy(ent).render({ canvasContext: ctx, viewport: pageProxy(ent).getViewport({ scale: k, rotation: g.rotation }), transform: g.cropped ? g.shift(k) : null, intent: 'print' }).promise;
      for (const o of ent.objects.filter(x => x.type === 'image')) {
        const im = new Image(); im.src = o.src; await im.decode().catch(() => {});
        const M = mul(Sc(k), mul(g.V1, o.m));
        ctx.save(); ctx.setTransform(...M); ctx.drawImage(im, 0, 0, o.w, o.h); ctx.restore();
      }
      const { data } = await worker.recognize(c, {}, { blocks: true, text: false });
      c.width = 0;
      const Vi = inv(g.V1);
      const out = [];
      for (const b of data.blocks || []) for (const p of b.paragraphs || []) for (const line of p.lines || []) for (const wd of line.words || []) {
        const t = (wd.text || '').trim();
        if (!t || wd.confidence < 35) continue;
        const { x0, y0, x1, y1 } = wd.bbox;
        const baseY = (line.baseline && line.baseline.y0 != null ? Math.max(line.baseline.y0, line.baseline.y1) : y1) / k;
        const o = apply(Vi, x0 / k, Math.min(baseY, y1 / k));
        const ex = apply(Vi, x1 / k, Math.min(baseY, y1 / k));
        const len = Math.hypot(ex[0] - o[0], ex[1] - o[1]) || 1;
        out.push({ t, o, u: [(ex[0] - o[0]) / len, (ex[1] - o[1]) / len], w: len, size: Math.max(4, ((y1 - y0) / k) * 0.85) });
      }
      ent.ocr = out;
      words += out.length;
      injectOcrText(ent);
      bar(pageNo / list.length);
    }
    commit();
    toast(cancelled ? `Stopped. Text recognized on ${pageNo - 1} page(s).` : `Recognized ${words} words. The pages are now searchable; save to keep the text.`);
  } catch (err) {
    console.error(err);
    await alertBox('Recognize Text', `Text recognition failed: ${friendly(err)}`);
  } finally {
    ui.remove();
  }
}
// make recognised words selectable and searchable right away (pdf.js text layer items)
function injectOcrText(ent) {
  if (!ent.ocr || !ent.ocr.length) return;
  const key = ent.blank ? `blank:${ent.id}` : keyOf(ent);
  const base = ent.blank ? Promise.resolve({ items: [], styles: {} }) : pageProxy(ent).getTextContent({ includeMarkedContent: true });
  const merged = base.then(tc => ({
    ...tc,
    styles: { ...tc.styles, nos_ocr: { fontFamily: 'sans-serif', ascent: 0.8, descent: -0.2, vertical: false } },
    items: [...tc.items, ...ent.ocr.map(w => ({ str: w.t, dir: 'ltr', width: w.w, height: w.size, transform: [w.u[0] * w.size, w.u[1] * w.size, -w.u[1] * w.size, w.u[0] * w.size, w.o[0], w.o[1]], fontName: 'nos_ocr', hasEOL: false }))],
  }));
  if (!ent.blank) textCache.set(key, merged);
  const v = views.get(ent.id);
  if (v) { v.done = 0; if (v.visible) renderView(v); }
}

// ---------------------------------------------------------------- side panel: bookmarks & comments
let sideTab = 'pages';
function showSide(tab) {
  sideTab = tab;
  for (const b of $$('.side-tabs [data-side]')) b.classList.toggle('on', b.dataset.side === tab);
  $('#thumb-list').hidden = tab !== 'pages';
  $('#bookmark-list').hidden = tab !== 'bookmarks';
  $('#comment-list').hidden = tab !== 'comments';
  if (tab === 'bookmarks') renderBookmarks();
  if (tab === 'comments') renderCommentList();
}
async function renderBookmarks() {
  const box = $('#bookmark-list');
  const src = S.sources[0];
  const outline = src ? await src.doc.getOutline().catch(() => null) : null;
  if (!outline || !outline.length) { box.innerHTML = '<div class="side-empty">This PDF has no bookmarks.</div>'; return; }
  const flat = [];
  const build = items => items.map(it => { const i = flat.push(it) - 1; return `<button class="bm" data-bm="${i}" title="${esc(it.title)}">${esc(it.title)}</button>${it.items && it.items.length ? `<div class="bm-children">${build(it.items)}</div>` : ''}`; }).join('');
  box.innerHTML = build(outline);
  box.onclick = ev => {
    const b = ev.target.closest('[data-bm]');
    const it = b && flat[+b.dataset.bm];
    if (!it) return;
    if (it.dest) goToDest(it.dest); else if (it.url) api.openExternal(it.url);
  };
}
function renderCommentList() {
  const box = $('#comment-list');
  const names = { note: 'Comment', highlight: 'Highlight', underline: 'Underline', strike: 'Strikethrough', text: 'Text', stamp: 'Stamp', ink: 'Drawing', rect: 'Rectangle', ellipse: 'Ellipse', line: 'Line', arrow: 'Arrow', image: 'Picture', redact: 'Redaction', whiteout: 'Whiteout', field: 'Form field' };
  const rows = [];
  S.entries.forEach((ent, i) => ent.objects.forEach(o => { if (!o.hf && !o.wm) rows.push({ ent, i, o }); }));
  if (!rows.length) { box.innerHTML = '<div class="side-empty">Comments, highlights, drawings and other things you add will be listed here.</div>'; return; }
  box.innerHTML = rows.map((r, k) => `<button class="cm-item" data-cm="${k}"><div class="k"><span class="sw" style="background:${r.o.color || (r.o.type === 'redact' ? '#000' : '#1f5fbf')}"></span>${esc(names[r.o.type] || r.o.type)}<small>Page ${r.i + 1}</small></div>${r.o.text ? `<div class="t">${esc(r.o.text)}</div>` : ''}</button>`).join('');
  box.onclick = ev => {
    const b = ev.target.closest('[data-cm]');
    const r = b && rows[+b.dataset.cm];
    if (!r) return;
    scrollToPage(r.i, false);
    setTool('select');
    setTimeout(() => { select(r.ent.id, r.o.id); const el = objectEl(r.o.id); if (el) el.scrollIntoView({ block: 'center' }); if (r.o.type === 'note') openNote(r.o); }, 150);
  };
}
$('.side-tabs').addEventListener('click', ev => { const b = ev.target.closest('[data-side]'); if (b) showSide(b.dataset.side); });

// ---------------------------------------------------------------- crash recovery
const recoveryWritten = new Set();
setInterval(() => {
  if (!S.open || !S.dirty) return;
  const sources = [];
  S.sources.forEach((s, index) => { if (!recoveryWritten.has(s)) { sources.push({ index, bytes: s.bytes }); recoveryWritten.add(s); } });
  api.pdfRecoverySave({ sources, state: { entries: S.entries, form: S.form, flattenForms: S.flattenForms, name: S.name, path: S.path, sourceNames: S.sources.map(s => s.name) } });
}, 45000);
async function restoreRecovery(rec) {
  const st = rec.state;
  const when = new Date(st.time || Date.now()).toLocaleString();
  if (!(await confirmBox('Recover unsaved PDF?', `Nisar Office Studio closed before "${st.name || 'a PDF'}" was saved. A copy of your changes from ${when} was kept.`, 'Recover', 'Discard'))) { window.close(); return; }
  setBusy('Recovering…');
  try {
    S.sources = [];
    for (let i = 0; i < rec.sources.length; i++) S.sources.push({ bytes: rec.sources[i], name: st.sourceNames[i], ...(await loadPdfjs(rec.sources[i], st.sourceNames[i])) });
    S.entries = st.entries;
    S.form = st.form || {};
    S.flattenForms = !!st.flattenForms;
    S.name = st.name || 'Recovered.pdf';
    S.path = st.path || null;
    S.open = true;
    $('#empty-state').hidden = true;
    buildPages();
    fitWidth();
    for (const ent of S.entries) injectOcrText(ent);
    resetHistory();
    S.dirty = true;
    updateTitle();
    toast('Recovered. Save it to keep your changes.');
  } catch (err) { await alertBox('Recover', friendly(err)); } finally { setBusy(null); }
}

// ---------------------------------------------------------------- menus
window.MENUS = {
  zoom: () => [...[0.5, 0.75, 1, 1.25, 1.5, 2, 3].map(z => item(`${z * 100}%`, () => setZoom(z), { check: Math.abs(S.zoom - z) < 0.001 && !S.fit })), '-', item('Fit Width', fitWidth, { check: S.fit === 'width' }), item('Fit Page', fitPage, { check: S.fit === 'page' })],
  sizes: () => [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72].map(n => item(String(n), () => setTextStyle({ size: n }))),
  color: () => colorMenu(c => { S.style.color = c || '#000000'; $('#color-swatch').style.background = S.style.color; if (S.sel) applyColorToSelected(S.style.color); }, null),
  textColor: () => colorMenu(c => setTextStyle({ color: c || '#000000' }), 'Automatic'),
  width: () => [0.5, 1, 2, 3, 4, 6, 8, 12].map(w => item(`${w} pt`, () => { S.style.width = w; if (S.sel) updateSelectedIf(o => 'width' in o, { width: w }); }, { check: S.style.width === w })),
  opacity: () => [1, 0.75, 0.5, 0.25].map(a => item(`${a * 100}%`, () => { S.style.opacity = a; if (S.sel) updateSelectedIf(o => 'opacity' in o, { opacity: a }); }, { check: S.style.opacity === a })),
  stamps: () => [...STAMPS.map(([t, c]) => item(t, () => placeStamp(t, c), { pre: `<span class="swatch" style="background:${c}"></span>` })), '-',
    item(`Date: ${new Date().toLocaleDateString()}`, () => placeStamp(new Date().toLocaleDateString(undefined, { dateStyle: 'medium' }).toUpperCase(), '#1f5fbf')),
    item('Custom Stamp…', async () => { const r = await askForm('Custom Stamp', [{ id: 't', label: 'Stamp text', value: 'REVIEWED' }], 'Add'); if (r && r.t.trim()) placeStamp(r.t.trim().toUpperCase().slice(0, 40), '#1f5fbf'); })],
  images: () => [item('Current Page', () => exportImages('current')), item('Selected Pages', () => exportImages('selected')), item('All Pages', () => exportImages('all'))],
  ocr: () => [item('Current Page', () => recognizeText('current')), item('Selected Pages', () => recognizeText('selected')), item('All Pages', () => recognizeText('all')), '-', { head: 'English text. Runs on this PC; nothing is uploaded.' }],
  objColor: () => colorMenu(c => applyColorToSelected(c || '#000000'), null),
  objFill: () => colorMenu(c => updateSelected({ fill: c }), 'No Fill'),
  objWidth: () => [0.5, 1, 2, 3, 4, 6, 8, 12].map(w => item(`${w} pt`, () => updateSelected({ width: w }))),
  objOpacity: () => [1, 0.75, 0.5, 0.25].map(a => item(`${a * 100}%`, () => updateSelected({ opacity: a }))),
};
function applyColorToSelected(c) {
  const hit = S.sel && objById(S.sel.oid);
  if (!hit) return;
  if (hit.o.type === 'highlight') S.style.hl = c;
  updateSelected({ color: c });
}
function updateSelectedIf(test, changes) {
  const hit = S.sel && objById(S.sel.oid);
  if (hit && test(hit.o)) updateSelected(changes);
}

// ---------------------------------------------------------------- commands
const COMMANDS = {
  save: () => save(false),
  saveAs: () => save(true),
  undo, redo,
  open: openDialog,
  print: printPdf,
  backstage: () => showBackstage('home'),
  closeBackstage: hideBackstage,
  closeDoc: () => { hideBackstage(); window.close(); },
  newBlank: async () => { hideBackstage(); if (S.open) api.newWindow({ pdf: true, template: 'blank' }); else await newBlankPdf(); },
  newDocument: () => { hideBackstage(); api.newWindow(); },
  zoomIn: () => stepZoom(1),
  zoomOut: () => stepZoom(-1),
  fitWidth,
  fitPage,
  find: openFind,
  goto: gotoPageDialog,
  thumbs: () => { document.body.classList.toggle('no-thumbs'); updateToggles(); if (S.fit) applyFit(); },
  toWord: () => { hideBackstage(); return exportToWord(); },
  collapseRibbon: () => document.body.classList.toggle('ribbon-collapsed'),
  applyColor: () => { if (S.sel) applyColorToSelected(S.style.color); },
  applyTextColor: () => setTextStyle({ color: S.style.textColor }),
  bold: () => { const hit = S.sel && objById(S.sel.oid); setTextStyle({ bold: !(hit && hit.o.type === 'text' ? hit.o.bold : S.style.bold) }); },
  italic: () => { const hit = S.sel && objById(S.sel.oid); setTextStyle({ italic: !(hit && hit.o.type === 'text' ? hit.o.italic : S.style.italic) }); },
  alignLeft: () => setTextStyle({ align: 'left' }),
  alignCenter: () => setTextStyle({ align: 'center' }),
  alignRight: () => setTextStyle({ align: 'right' }),
  image: insertImage,
  signature: signatureDialog,
  clearForm: async () => {
    if (!(await confirmBox('Clear Form', 'Clear every form field in this PDF?', 'Clear', 'Cancel'))) return;
    for (const el of $$('.fld')) { if (el.type === 'checkbox' || el.type === 'radio') el.checked = false; else if (el.tagName === 'SELECT') el.selectedIndex = -1; else el.value = ''; }
    const names = new Set();
    for (const ent of S.entries) if (!ent.blank && ent.src === 0) for (const a of await annotations(ent)) if (a.subtype === 'Widget' && a.fieldName) names.add(a);
    for (const a of names) S.form[a.fieldName] = a.fieldType === 'Btn' ? (a.checkBox ? false : 'Off') : a.fieldType === 'Ch' && a.multiSelect ? [] : '';
    commit();
  },
  flattenForms: () => { S.flattenForms = !S.flattenForms; updateToggles(); markDirty(); toast(S.flattenForms ? 'Form fields will become plain text when you save.' : 'Form fields will stay fillable when you save.'); },
  insertBlank, insertPdf, insertImagePage,
  rotateLeft: () => rotatePages(-1),
  rotateRight: () => rotatePages(1),
  moveUp: () => movePages(-1),
  moveDown: () => movePages(1),
  duplicate: duplicatePages,
  deletePages,
  extract: extractPages,
  headerFooter: headerFooterDialog,
  watermark: watermarkDialog,
  crop: cropDialog,
  split: splitDialog,
  combine: combineFiles,
  imagesToPdf: async () => { if (!S.open) { await newBlankPdf(); const first = S.entries[0]; await imagesToPages(); if (S.entries.length > 1 && first && !first.objects.length) { S.entries = S.entries.filter(e => e !== first); afterStructureChange(0); } } else await imagesToPages(); },
  findRedact: findRedactDialog,
  protect: protectDialog,
  unprotect,
  sanitize: sanitizeToggle,
  compress: compressCopy,
};
const NEEDS_DOC = new Set(['save', 'saveAs', 'print', 'zoomIn', 'zoomOut', 'fitWidth', 'fitPage', 'find', 'goto', 'toWord', 'image', 'signature', 'clearForm', 'flattenForms', 'insertBlank', 'insertPdf', 'insertImagePage', 'rotateLeft', 'rotateRight', 'moveUp', 'moveDown', 'duplicate', 'deletePages', 'extract', 'headerFooter', 'watermark', 'crop', 'split', 'combine', 'findRedact', 'protect', 'unprotect', 'sanitize', 'compress']);
function runCommand(name, el) {
  const fn = COMMANDS[name];
  if (!fn) return;
  if (NEEDS_DOC.has(name) && !S.open) { toast('Open a PDF first.'); return; }
  try { const r = fn(el); if (r && r.catch) r.catch(reportError); } catch (err) { reportError(err); }
}
function updateToggles() {
  const set = (cmd, on) => $$(`[data-cmd="${cmd}"]`).forEach(b => b.classList.toggle('on', !!on));
  set('thumbs', !document.body.classList.contains('no-thumbs'));
  set('flattenForms', S.flattenForms);
  set('sanitize', S.sanitize);
  set('protect', !!S.protect);
  set('fitWidth', S.fit === 'width');
  set('fitPage', S.fit === 'page');
}

// ---------------------------------------------------------------- events
document.addEventListener('mousedown', ev => {
  if (ev.target.closest('#ribbon button, #statusbar button, .menu button, #objbar button')) ev.preventDefault();
  if (!ev.target.closest('.menu') && !ev.target.closest('[data-menu]')) closeMenu();
  if (!$('#notepop').hidden && !ev.target.closest('#notepop') && !ev.target.closest('.obj-note')) closeNote(true);
});
document.addEventListener('click', ev => {
  const m = ev.target.closest('[data-menu]');
  if (m) { openMenu(m.dataset.menu, m); return; }
  const t = ev.target.closest('[data-tool]');
  if (t && t.tagName === 'BUTTON') { if (!S.open) { toast('Open a PDF first.'); return; } setTool(t.dataset.tool); return; }
  const c = ev.target.closest('[data-cmd]');
  if (c) { closeMenu(); runCommand(c.dataset.cmd, c); return; }
  const tab = ev.target.closest('[data-tab]');
  if (tab) {
    for (const x of $$('.tab[data-tab]')) x.classList.toggle('active', x === tab);
    for (const p of $$('.panel')) p.classList.toggle('active', p.dataset.panel === tab.dataset.tab);
    document.body.classList.remove('ribbon-collapsed');
    return;
  }
  const bs = ev.target.closest('[data-bs]');
  if (bs) { switchPane(bs.dataset.bs); return; }
  const f = ev.target.closest('[data-find]');
  if (f) { ({ prev: () => gotoMatch(find.index - 1), next: () => gotoMatch(find.index + 1), close: closeFind })[f.dataset.find](); return; }
  const ob = ev.target.closest('[data-ob]');
  if (ob && ob.tagName === 'BUTTON') {
    const hit = S.sel && objById(S.sel.oid);
    if (!hit) return;
    const a = ob.dataset.ob;
    if (a === 'delete') { deleteObject(hit.o.id); commit(); }
    else if (a === 'duplicate') duplicateObject();
    else if (a === 'bold' || a === 'italic') setTextStyle({ [a]: !hit.o[a] });
    else if (a === 'edit') startEditing(hit.o.id);
    else if (a === 'note') openNote(hit.o);
    else if (a === 'fmulti') updateSelected({ multiline: !hit.o.multiline });
    else if (a === 'fopts') {
      askForm('Drop-down Options', [{ id: 'opts', label: 'Choices, separated by commas', value: (hit.o.options || []).join(', ') }], 'OK').then(r => { if (r) updateSelected({ options: r.opts.split(',').map(s => s.trim()).filter(Boolean) }); });
    }
    return;
  }
  const np = ev.target.closest('[data-np]');
  if (np) {
    const hit = objById($('#notepop').dataset.oid);
    if (np.dataset.np === 'delete' && hit) { closeNote(false); deleteObject(hit.o.id); commit(); }
    else closeNote(true);
  }
});
$('#objbar').addEventListener('change', ev => {
  const k = ev.target.dataset.ob;
  if (k === 'font') setTextStyle({ font: ev.target.value });
  if (k === 'size') { const n = parseFloat(ev.target.value); if (n >= 2 && n <= 400) setTextStyle({ size: n }); }
  if (k === 'fname') { const n = ev.target.value.trim().replace(/[.]/g, '_'); if (n) updateSelected({ name: n }); }
});
$('#font-select').addEventListener('change', ev => setTextStyle({ font: ev.target.value }));
$('#size-input').addEventListener('change', ev => { const n = parseFloat(ev.target.value); if (n >= 2 && n <= 400) setTextStyle({ size: n }); else updateFontControls(); });
$('#zoom-input').addEventListener('change', ev => { const n = parseFloat(ev.target.value); if (n > 0) setZoom(n / 100); else layoutAfterZoom(); });
$('#zoom-range').addEventListener('input', ev => setZoom(ev.target.value / 100));
function duplicateObject() {
  const hit = S.sel && objById(S.sel.oid);
  if (!hit) return;
  const copy = { ...hit.o, id: uid(), m: mul(hit.o.m, T(12, 12)) };
  hit.ent.objects.push(copy);
  const v = views.get(hit.ent.id);
  if (v) renderObjects(v);
  select(hit.ent.id, copy.id);
  commit();
}
function pasteObject() {
  if (!S.clip) return false;
  const v = currentView();
  if (!v) return false;
  const c = visibleCenter(v);
  const g = c.g;
  const o = { ...S.clip, id: uid(), m: mul(inv(g.V1), T(c.x - S.clip.w / 2, c.y - S.clip.h / 2)) };
  v.ent.objects.push(o);
  renderObjects(v);
  select(v.ent.id, o.id);
  commit();
  return true;
}
const typingIn = el => el && (el.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(el.tagName));
document.addEventListener('keydown', ev => {
  const k = ev.key, ctrl = ev.ctrlKey || ev.metaKey;
  const typing = typingIn(document.activeElement);
  if (k === 'Escape') {
    if (openMenuEl) { closeMenu(); return; }
    if ($('#modal').open) return;
    if (!$('#backstage').hidden) { hideBackstage(); return; }
    if (!$('#notepop').hidden) { closeNote(true); return; }
    if (S.editing) { finishEditing(); return; }
    if (!$('#findbar').hidden) { closeFind(); return; }
    if (S.sel) { select(null); return; }
    if (S.tool !== 'select') setTool('select');
    return;
  }
  if ($('#modal').open) return;
  if (ev.target === findInput && k === 'Enter') { ev.preventDefault(); gotoMatch(find.index + (ev.shiftKey ? -1 : 1)); return; }
  if (ctrl) {
    const key = k.toLowerCase();
    const map = {
      s: () => save(ev.shiftKey), o: openDialog, p: printPdf, f: openFind, w: () => window.close(), n: () => api.newWindow(),
      '=': () => stepZoom(1), '+': () => stepZoom(1), '-': () => stepZoom(-1), '0': fitWidth,
    };
    if (!typing || ['s', 'o', 'p', 'w', 'n'].includes(key)) {
      if (key === 'z' && !ev.shiftKey && !typing) { ev.preventDefault(); undo(); return; }
      if ((key === 'y' || (key === 'z' && ev.shiftKey)) && !typing) { ev.preventDefault(); redo(); return; }
      if (key === 'c' && S.sel && !typing && getSelection().isCollapsed) { const hit = objById(S.sel.oid); if (hit) { S.clip = { ...hit.o }; ev.preventDefault(); toast('Copied. Press Ctrl+V to paste.'); } return; }
      if (key === 'v' && !typing && S.clip) { ev.preventDefault(); pasteObject(); return; }
      if (key === 'd' && S.sel && !typing) { ev.preventDefault(); duplicateObject(); return; }
      if (map[key] && (S.open || ['o', 'w', 'n'].includes(key))) { ev.preventDefault(); const r = map[key](); if (r && r.catch) r.catch(reportError); return; }
    }
    if (k === 'F1') { ev.preventDefault(); runCommand('collapseRibbon'); }
    return;
  }
  if (typing) return;
  if ((k === 'Delete' || k === 'Backspace') && S.sel) { ev.preventDefault(); deleteObject(S.sel.oid); commit(); return; }
  if (S.sel && k.startsWith('Arrow')) {
    ev.preventDefault();
    const hit = objById(S.sel.oid);
    const v = hit && views.get(hit.ent.id);
    if (!v) return;
    const step = ev.shiftKey ? 10 : 1;
    const d = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[k];
    const [lx, ly] = localDelta(v, hit.o, d[0], d[1], 1);
    hit.o.m = mul(hit.o.m, T(lx, ly));
    renderObjects(v);
    positionObjbar();
    clearTimeout(nudgeTimer);
    nudgeTimer = setTimeout(commit, 400);
    return;
  }
  if (!S.open) return;
  const tools = { v: 'select', h: 'hand', t: 'text', d: 'ink' };
  if (tools[k.toLowerCase()] && !ev.altKey) { setTool(tools[k.toLowerCase()]); return; }
  if (k === 'Home') { ev.preventDefault(); scrollToPage(0, false); }
  if (k === 'End') { ev.preventDefault(); scrollToPage(S.entries.length - 1, false); }
});
let nudgeTimer = null;

// paste an image from the clipboard onto the page
document.addEventListener('paste', ev => {
  if (typingIn(document.activeElement) || !S.open) return;
  const file = [...(ev.clipboardData ? ev.clipboardData.files : [])].find(f => f.type.startsWith('image/'));
  if (!file) return;
  ev.preventDefault();
  const fr = new FileReader();
  fr.onload = () => placeImage(fr.result);
  fr.readAsDataURL(file);
});

// drag & drop files
document.addEventListener('dragover', ev => { if ([...ev.dataTransfer.types].includes('Files')) { ev.preventDefault(); document.body.classList.add('dropping'); } });
document.addEventListener('dragleave', ev => { if (!ev.relatedTarget) document.body.classList.remove('dropping'); });
document.addEventListener('drop', async ev => {
  document.body.classList.remove('dropping');
  const files = [...(ev.dataTransfer ? ev.dataTransfer.files : [])];
  if (!files.length) return;
  ev.preventDefault();
  for (const f of files) {
    if (f.type.startsWith('image/') && S.open) {
      const fr = new FileReader();
      fr.onload = () => placeImage(fr.result);
      fr.readAsDataURL(f);
      continue;
    }
    const p = api.pathForFile(f);
    if (p) await openPath(p);
  }
});

api.on('request-close', async () => { if (await confirmDiscard()) api.closeNow(); });
api.on('cmd', () => {});

// ---------------------------------------------------------------- start
const SHORTCUTS = [
  ['Open', 'Ctrl+O'], ['Save / Save As', 'Ctrl+S / Ctrl+Shift+S'], ['Print', 'Ctrl+P'], ['Find', 'Ctrl+F'], ['Undo / Redo', 'Ctrl+Z / Ctrl+Y'],
  ['Zoom in / out', 'Ctrl+= / Ctrl+-'], ['Fit width', 'Ctrl+0'], ['Zoom with the mouse', 'Ctrl+wheel'], ['First / last page', 'Home / End'],
  ['Select tool', 'V'], ['Hand tool', 'H'], ['Add text', 'T'], ['Draw', 'D'], ['Delete selected item', 'Delete'],
  ['Move selected item', 'Arrow keys (Shift = 10×)'], ['Copy / paste item', 'Ctrl+C / Ctrl+V'], ['Duplicate item', 'Ctrl+D'], ['Edit added text', 'Double-click'],
  ['Straight lines at 45°', 'Shift while drawing'], ['New document window', 'Ctrl+N'], ['Close', 'Ctrl+W'], ['Cancel / back', 'Esc'],
];
(async function init() {
  for (const el of $$('[data-icon]')) el.insertAdjacentHTML('afterbegin', svg(el.dataset.icon));
  $('#shortcuts').innerHTML = SHORTCUTS.map(([a, b]) => `<div><span>${esc(a)}</span><kbd>${esc(b)}</kbd></div>`).join('');
  $('#color-swatch').style.background = S.style.color;
  setTool('select');
  updateFontControls();
  updateToggles();
  updateTitle();
  let init = {};
  try { init = await api.init(); } catch (err) { console.error(err); }
  if (init.dark) document.body.classList.add('dark');
  S.signatures = Array.isArray(init.signatures) ? init.signatures : [];
  S.untitled = init.untitled || 1;
  if (init.pdfPath) {
    const r = await api.readPdf(init.pdfPath);
    if (r.error) { $('#empty-state').hidden = false; await alertBox("Couldn't open", r.error); return; }
    try { await openBytes(r.bytes, r.name, r.path); } catch (err) {
      $('#empty-state').hidden = false;
      if (!(err && /destroy/i.test(String(err.message)))) await alertBox("Couldn't open", `"${r.name}" couldn't be opened. It may be damaged or not a real PDF.`);
    }
  } else if (init.recovery) {
    await restoreRecovery(init.recovery);
  } else if (init.template === 'blank') {
    await newBlankPdf();
  } else {
    $('#empty-state').hidden = false;
  }
})();

// for automated tests
window.__pdf = { recognizeText, downsampleImages, renderBookmarks, renderCommentList, showSide, ensureEditable, S, views, find, buildPdf, setTool, createObject, geom, openBytes, save, select, commit, undo, redo, exportToWord, runFind, gotoMatch, viewportMatrix, rotatePages, movePages, duplicatePages, deletePages, insertBlank, insertPdfBytes, markupFromSelection, startEditing, finishEditing, setZoom, fitWidth, scrollToPage, pdfjsLib, PDF_OPTS };
