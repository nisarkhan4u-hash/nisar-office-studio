'use strict';
// Nisar Office Studio: main process (windows, files, .docx conversion, PDF, printing).
const { app, BrowserWindow, ipcMain, dialog, Menu, shell, clipboard, session } = require('electron');
const path = require('path');
const fs = require('fs');
const fsp = fs.promises;
const crypto = require('crypto');

const APP_NAME = 'Nisar Office Studio';
const OPENABLE = ['.docx', '.html', '.htm', '.txt', '.pdf'];
const isPdfPath = p => typeof p === 'string' && path.extname(p).toLowerCase() === '.pdf';
const PAGES = { doc: 'index.html', pdf: 'pdf.html' };
// Our own copy of the editor HTML, stored inside saved .docx files so formatting
// round-trips perfectly. It is ignored if the file was later changed by another app.
const DATA_PART = 'nisarword/document.json';
const DATA_REL = 'http://schemas.nisarword.local/2026/relationships/document-data';
// html-to-docx has no strikethrough support; runs in this marker font become <w:strike/>.
const STRIKE_FONT = 'NisarWordStrike';

app.setName(APP_NAME);
if (process.platform === 'win32') app.setAppUserModelId('com.nisar.officestudio');

const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  start();
}

let settings = { recent: [], dark: false };
let untitledCounter = 0;
let leftoverRecovery = [];
let recoveryOffered = false;

const settingsFile = () => path.join(app.getPath('userData'), 'settings.json');
const recoveryDir = () => path.join(app.getPath('userData'), 'recovery');

// the app used to be called "Nisar Studio" and before that "Nisar Word": bring the newest
// settings, recent files and custom dictionary across the first time this version starts
function migrateOldProfile() {
  try {
    const now = app.getPath('userData');
    const old = ['Nisar Studio', 'Nisar Word'].map(n => path.join(app.getPath('appData'), n)).find(p => fs.existsSync(path.join(p, 'settings.json')));
    if (!old || fs.existsSync(path.join(now, 'settings.json'))) return;
    fs.mkdirSync(now, { recursive: true });
    for (const f of ['settings.json', 'Custom Dictionary.txt']) {
      if (fs.existsSync(path.join(old, f))) fs.copyFileSync(path.join(old, f), path.join(now, f));
    }
  } catch { /* nothing to migrate */ }
}
function loadSettings() {
  try { settings = { ...settings, ...JSON.parse(fs.readFileSync(settingsFile(), 'utf8')) }; } catch { /* first run */ }
  if (!Array.isArray(settings.recent)) settings.recent = [];
}
function saveSettings() {
  try { fs.writeFileSync(settingsFile(), JSON.stringify(settings, null, 2)); } catch { /* not fatal */ }
}
function addRecent(p) {
  settings.recent = [p, ...settings.recent.filter(r => r.toLowerCase() !== p.toLowerCase())].slice(0, 15);
  saveSettings();
}

function fileFromArgv(argv) {
  return argv.slice(1).find(a => !a.startsWith('-') && OPENABLE.includes(path.extname(a).toLowerCase()) && fs.existsSync(a));
}

function defaultUserName() {
  try { const n = require('os').userInfo().username; return n ? n.charAt(0).toUpperCase() + n.slice(1) : 'Author'; } catch { return 'Author'; }
}
const escHtml = s => String(s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const sha = s => crypto.createHash('sha256').update(s).digest('hex');

function createWindow(opts = {}) {
  const win = new BrowserWindow({
    width: 1320,
    height: 880,
    minWidth: 860,
    minHeight: 560,
    show: false,
    title: APP_NAME,
    icon: path.join(__dirname, 'build', 'icon.png'),
    backgroundColor: settings.dark ? '#1f1f1f' : '#f3f3f3',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: true,
      // fonts a document asks for but this PC doesn't have fall back to Calibri, not Times
      defaultFontFamily: { standard: 'Calibri', sansSerif: 'Calibri', serif: 'Cambria', monospace: 'Consolas' },
    },
  });
  win.pendingFile = opts.file || null;
  win.pendingTemplate = opts.template || null;
  win.pendingDoc = opts.doc || null; // { html, name } handed over from the PDF editor
  win.untitled = ++untitledCounter;
  win.allowClose = false;
  win.inTable = false;
  win.recoveryFile = path.join(recoveryDir(), `${Date.now()}-${win.id}.json`);

  win.once('ready-to-show', () => { win.maximize(); win.show(); });
  loadKind(win, opts.pdf || isPdfPath(opts.file) ? 'pdf' : 'doc');

  win.on('close', e => {
    if (!win.allowClose) {
      e.preventDefault();
      win.webContents.send('request-close');
    }
  });
  win.on('closed', () => {
    fs.rm(win.recoveryFile, { force: true }, () => {});
    fs.rm(pdfRecoveryDir(win), { recursive: true, force: true }, () => {});
  });
  win.pendingRecover = opts.recover || null;

  const wc = win.webContents;
  wc.on('will-navigate', e => e.preventDefault());
  wc.setWindowOpenHandler(({ url }) => { openExternal(url); return { action: 'deny' }; });
  wc.on('render-process-gone', (e, details) => {
    if (details.reason !== 'clean-exit') {
      dialog.showMessageBox(win, { type: 'error', title: APP_NAME, message: 'Nisar Office Studio ran into a problem and needs to reload this window.', detail: 'Your latest autosaved copy will be offered for recovery.' })
        .then(() => { if (!win.isDestroyed()) wc.reload(); });
    }
  });
  wc.on('context-menu', (e, p) => buildContextMenu(win, p));
  return win;
}

// a window shows either the document editor or the PDF editor
function loadKind(win, kind) {
  win.kind = kind;
  win.inTable = false;
  win.loadFile(path.join(__dirname, 'src', PAGES[kind]));
}

function buildContextMenu(win, p) {
  const send = (...a) => win.webContents.send(...a);
  const t = [];
  if (p.misspelledWord) {
    for (const s of p.dictionarySuggestions.slice(0, 6)) t.push({ label: s, click: () => win.webContents.replaceMisspelling(s) });
    if (!p.dictionarySuggestions.length) t.push({ label: 'No spelling suggestions', enabled: false });
    t.push({ label: 'Add to Dictionary', click: () => win.webContents.session.addWordToSpellCheckerDictionary(p.misspelledWord) });
    t.push({ type: 'separator' });
  }
  if (p.isEditable) {
    t.push(
      { label: 'Cut', role: 'cut', enabled: p.editFlags.canCut },
      { label: 'Copy', role: 'copy', enabled: p.editFlags.canCopy },
      { label: 'Paste', role: 'paste', enabled: p.editFlags.canPaste },
      { label: 'Paste as Plain Text', click: () => send('cmd', 'pastePlain') },
      { type: 'separator' },
      { label: 'Select All', role: 'selectAll' },
    );
    if (p.linkURL) {
      t.push({ type: 'separator' },
        { label: 'Open Link', click: () => openExternal(p.linkURL) },
        { label: 'Edit Link…', click: () => send('cmd', 'link') },
        { label: 'Remove Link', click: () => send('cmd', 'unlink') });
    }
    if (win.inTable) {
      t.push({ type: 'separator' }, {
        label: 'Table',
        submenu: [
          { label: 'Insert Row Above', click: () => send('cmd', 'rowAbove') },
          { label: 'Insert Row Below', click: () => send('cmd', 'rowBelow') },
          { label: 'Insert Column Left', click: () => send('cmd', 'colLeft') },
          { label: 'Insert Column Right', click: () => send('cmd', 'colRight') },
          { type: 'separator' },
          { label: 'Delete Row', click: () => send('cmd', 'delRow') },
          { label: 'Delete Column', click: () => send('cmd', 'delCol') },
          { label: 'Delete Table', click: () => send('cmd', 'delTable') },
        ],
      });
    }
  } else if (p.selectionText) {
    t.push({ label: 'Copy', role: 'copy' });
  }
  if (t.length) Menu.buildFromTemplate(t).popup({ window: win });
}

function openExternal(url) {
  if (typeof url === 'string' && /^(https?:|mailto:)/i.test(url)) shell.openExternal(url);
}

// ---------- reading documents ----------

const PAPERS = { letter: [8.5, 11], legal: [8.5, 14], a4: [8.27, 11.69], a5: [5.83, 8.27], executive: [7.25, 10.5], tabloid: [11, 17] };

function parsePageSetup(xml) {
  const sect = xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>(?![\s\S]*<w:sectPr)/);
  if (!sect) return null;
  const s = sect[0];
  const attr = (tag, name) => { const m = s.match(new RegExp(`<w:${tag}\\b[^>]*\\bw:${name}="(-?\\d+)"`)); return m ? +m[1] : null; };
  const w = attr('pgSz', 'w'), h = attr('pgSz', 'h');
  if (!w || !h) return null;
  const wi = w / 1440, hi = h / 1440;
  const landscape = wi > hi;
  const pw = Math.min(wi, hi), ph = Math.max(wi, hi);
  let size = 'letter';
  for (const [k, [a, b]] of Object.entries(PAPERS)) if (Math.abs(a - pw) < 0.12 && Math.abs(b - ph) < 0.12) size = k;
  const m = n => { const v = attr('pgMar', n); return v == null ? 1 : Math.max(0, Math.round((Math.abs(v) / 1440) * 100) / 100); };
  const cols = s.match(/<w:cols\b[^>]*\bw:num="(\d+)"/);
  return {
    size, orientation: landscape ? 'landscape' : 'portrait',
    margins: { top: m('top'), right: m('right'), bottom: m('bottom'), left: m('left') },
    columns: cols ? Math.min(3, Math.max(1, +cols[1])) : 1,
  };
}

async function docxToHtml(buf) {
  const JSZip = require('jszip');
  let page = null;
  try {
    const zip = await JSZip.loadAsync(buf);
    const docPart = zip.file('word/document.xml');
    const xml = docPart ? await docPart.async('string') : '';
    const part = zip.file(DATA_PART);
    if (part && xml) {
      const payload = JSON.parse(await part.async('string'));
      if (payload && payload.hash === sha(xml) && typeof payload.html === 'string') {
        return { html: payload.html, page: payload.page, extra: payload.extra || {} };
      }
    }
    page = parsePageSetup(xml);
    // our own converter keeps fonts, colours, sizes, spacing, lists and tables
    const { convertDocx } = require('./docx-import');
    const out = await convertDocx(zip, parsePageSetup);
    if (out.html.trim()) return { html: out.html, page: { ...page, ...out.page }, extra: out.extra || {}, imported: true };
  } catch (err) {
    console.error('docx import fell back to the basic converter:', err);
  }

  const mammoth = require('mammoth');
  const r = await mammoth.convertToHtml({ buffer: buf }, {
    styleMap: [
      'u => u',
      'strike => s',
      "p[style-name='Title'] => h1.doc-title:fresh",
      "p[style-name='Subtitle'] => p.doc-subtitle:fresh",
      "p[style-name='Quote'] => blockquote:fresh",
      "p[style-name='Intense Quote'] => blockquote:fresh",
    ],
  });
  return { html: r.value, page, imported: true };
}

async function readDocument(p) {
  const ext = path.extname(p).toLowerCase();
  const name = path.basename(p);
  try {
    if (ext === '.pdf') return { pdf: true, path: p, name };
    if (ext === '.doc') return { error: `"${name}" is an old Word 97–2003 (.doc) file. Nisar Office Studio opens .docx files. Re-save it as .docx and try again.` };
    if (!OPENABLE.includes(ext)) return { error: `Nisar Office Studio can't open "${name}". Supported files: .docx, .html, .txt` };
    const buf = await fsp.readFile(p);
    let out;
    if (ext === '.docx') out = await docxToHtml(buf);
    else if (ext === '.txt') out = { text: buf.toString('utf8').replace(/^﻿/, '') };
    else out = { html: buf.toString('utf8') };
    addRecent(p);
    return { ...out, path: p, name };
  } catch (err) {
    if (err.code === 'ENOENT') {
      settings.recent = settings.recent.filter(r => r !== p);
      saveSettings();
      return { error: `"${name}" couldn't be found. It may have been moved or deleted.` };
    }
    return { error: `"${name}" couldn't be opened. ${err.message || err}` };
  }
}

// ---------- writing documents ----------

async function buildDocx(d) {
  const HTMLtoDOCX = require('html-to-docx');
  const JSZip = require('jszip');
  const p = d.page;
  const tw = n => Math.round(n * 1440);
  const landscape = p.orientation === 'landscape';
  const hasHeader = !!p.header;
  const hasFooter = !!(p.footer || p.pageNumbers);
  const footerText = escHtml(p.footer) + (p.footer && p.pageNumbers ? ' · ' : '') + (p.pageNumbers ? 'Page ' : '');

  let buf = await HTMLtoDOCX(
    `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${d.exportHtml}</body></html>`,
    hasHeader ? `<p style="text-align:center">${escHtml(p.header)}</p>` : null,
    {
      orientation: landscape ? 'landscape' : 'portrait',
      pageSize: { width: tw(p.w), height: tw(p.h) },
      margins: { top: tw(p.margins.top), right: tw(p.margins.right), bottom: tw(p.margins.bottom), left: tw(p.margins.left), header: 720, footer: 720, gutter: 0 },
      title: d.name,
      creator: 'Nisar Office Studio',
      lastModifiedBy: 'Nisar Office Studio',
      font: 'Calibri',
      fontSize: 22,
      complexScriptFontSize: 22,
      header: hasHeader,
      footer: hasFooter,
      pageNumber: !!p.pageNumbers,
      table: { row: { cantSplit: true } },
    },
    hasFooter ? `<p style="text-align:center">${footerText}</p>` : null,
  );
  if (!Buffer.isBuffer(buf)) buf = Buffer.from(await buf.arrayBuffer());

  const zip = await JSZip.loadAsync(buf);
  let xml = await zip.file('word/document.xml').async('string');
  // html-to-docx writes portrait dimensions even for landscape pages; write the real ones.
  xml = xml.replace(/<w:pgSz\b[^>]*\/>/, `<w:pgSz w:w="${tw(p.w)}" w:h="${tw(p.h)}" w:orient="${landscape ? 'landscape' : 'portrait'}"/>`);
  if (p.columns > 1) xml = xml.replace(/(<w:pgMar\b[^>]*\/>)/, `$1<w:cols w:num="${p.columns}" w:space="720"/>`);
  xml = xml.replace(new RegExp(`<w:rFonts w:ascii="${STRIKE_FONT}" w:hAnsi="${STRIKE_FONT}"/>`, 'g'), '<w:strike/>');
  // comments, tracked changes and footnotes become real Word parts
  const { applyDocxExtras } = require('./docx-extras');
  xml = await applyDocxExtras(zip, xml, d.docxExtras);
  zip.file('word/document.xml', xml);

  zip.file(DATA_PART, JSON.stringify({ app: APP_NAME, version: 1, hash: sha(xml), html: d.raw, page: p, extra: d.extra || {} }));
  let types = await zip.file('[Content_Types].xml').async('string');
  if (!/Extension="json"/i.test(types)) types = types.replace(/<Types\b[^>]*>/, m => `${m}<Default Extension="json" ContentType="application/json"/>`);
  zip.file('[Content_Types].xml', types);
  let rels = await zip.file('_rels/.rels').async('string');
  if (!rels.includes(DATA_REL)) rels = rels.replace('</Relationships>', `<Relationship Id="rIdNisarWord1" Type="${DATA_REL}" Target="${DATA_PART}"/></Relationships>`);
  zip.file('_rels/.rels', rels);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function writeAtomic(target, data) {
  const tmp = `${target}.${process.pid}.tmp`;
  await fsp.writeFile(tmp, data);
  try {
    await fsp.rename(tmp, target);
  } catch (err) {
    await fsp.rm(tmp, { force: true });
    throw err;
  }
}

function saveErrorMessage(err, target) {
  if (err.code === 'EBUSY' || err.code === 'EPERM' || err.code === 'EACCES') {
    return `"${path.basename(target)}" is open in another program or the folder is read-only. Close it there, or use Save As to pick another name.`;
  }
  return err.message || String(err);
}

// ---------- IPC ----------

// ---------- qpdf (bundled, Apache-2.0) ----------
function qpdfExe() {
  const packaged = path.join(process.resourcesPath || '', 'qpdf', 'qpdf.exe');
  if (fs.existsSync(packaged)) return packaged;
  return path.join(__dirname, 'vendor', 'qpdf', 'bin', 'qpdf.exe');
}
// d: { op: 'encrypt' | 'decrypt' | 'optimize' | 'check', bytes, user, owner, allowPrint, allowCopy, allowModify, password }
async function runQpdf(d) {
  const { execFile } = require('child_process');
  const dir = await fsp.mkdtemp(path.join(app.getPath('temp'), 'nos-qpdf-'));
  const inFile = path.join(dir, 'in.pdf');
  const outFile = path.join(dir, 'out.pdf');
  try {
    await fsp.writeFile(inFile, Buffer.from(d.bytes));
    const args = [];
    if (d.password) args.push(`--password=${d.password}`);
    if (d.op === 'check') args.push('--requires-password', inFile);
    else {
      if (d.op === 'encrypt') {
        args.push(inFile, '--encrypt', String(d.user || ''), String(d.owner || d.user || ''), '256',
          `--print=${d.allowPrint === false ? 'none' : 'full'}`, `--extract=${d.allowCopy === false ? 'n' : 'y'}`, `--modify=${d.allowModify === false ? 'none' : 'all'}`, '--', outFile);
      } else if (d.op === 'decrypt') args.push('--decrypt', inFile, outFile);
      else if (d.op === 'optimize') args.push('--object-streams=generate', '--compress-streams=y', '--recompress-flate', '--compression-level=9', '--remove-unreferenced-resources=yes', inFile, outFile);
      else return { error: 'Unknown operation.' };
    }
    const result = await new Promise(resolve => {
      execFile(qpdfExe(), args, { windowsHide: true, timeout: 180000, maxBuffer: 8 * 1024 * 1024 }, (err, stdout, stderr) => resolve({ code: err ? (typeof err.code === 'number' ? err.code : 1) : 0, stdout: String(stdout), stderr: String(stderr) }));
    });
    if (d.op === 'check') return { needsPassword: result.code === 0 };
    // qpdf exit code 3 means "succeeded with warnings"
    if ((result.code === 0 || result.code === 3) && fs.existsSync(outFile)) {
      const out = await fsp.readFile(outFile);
      return { bytes: new Uint8Array(out.buffer, out.byteOffset, out.length) };
    }
    if (/invalid password/i.test(result.stderr)) return { error: 'That password is not correct.', badPassword: true };
    return { error: (result.stderr || 'qpdf failed').split('\n').filter(Boolean).slice(-1)[0] };
  } catch (err) {
    return { error: err.message };
  } finally {
    fs.rm(dir, { recursive: true, force: true }, () => {});
  }
}
const pdfRecoveryDir = win => path.join(recoveryDir(), `pdf-${win.recoveryFile.match(/(\d+-\d+)\.json$/)[1]}`);

// Print Preview: Chromium's PDF viewer in a child window (it has the Print button)
async function openPdfViewerWindow(win, data, name) {
  const dir = path.join(app.getPath('temp'), 'NisarOfficeStudioPreview');
  await fsp.mkdir(dir, { recursive: true });
  const file = path.join(dir, `preview-${win.id}-${Date.now()}.pdf`);
  await fsp.writeFile(file, data);
  const pv = new BrowserWindow({
    parent: win, width: 1000, height: Math.min(1100, win.getBounds().height), title: `Print Preview - ${name || 'Document'}`,
    icon: path.join(__dirname, 'build', 'icon.png'), autoHideMenuBar: true,
    webPreferences: { plugins: true, sandbox: true, contextIsolation: true, nodeIntegration: false },
  });
  pv.setMenu(null);
  pv.on('page-title-updated', ev => ev.preventDefault()); // the PDF's own title would replace ours
  pv.webContents.on('will-navigate', ev => ev.preventDefault());
  pv.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  pv.on('closed', () => fs.rm(file, { force: true }, () => {}));
  await pv.loadFile(file);
  return { ok: true };
}

function whitePaper(win) {
  if (!win || win.isDestroyed()) return () => {};
  const before = win.getBackgroundColor();
  win.setBackgroundColor('#ffffff');
  return () => { if (!win.isDestroyed()) win.setBackgroundColor(before); };
}

function registerIpc() {
  const winOf = e => BrowserWindow.fromWebContents(e.sender);

  ipcMain.handle('get-init', async e => {
    const win = winOf(e);
    const out = { kind: win.kind, untitled: win.untitled, dark: !!settings.dark, recent: settings.recent, template: win.pendingTemplate, signatures: settings.signatures || [], userName: settings.userName || defaultUserName(), initials: settings.initials || '' };
    win.pendingTemplate = null;
    if (win.kind === 'pdf') {
      // the PDF editor reads its file itself (as bytes)
      out.pdfPath = win.pendingFile;
      out.blankPdf = !win.pendingFile;
      win.pendingFile = null;
      if (win.pendingRecover) {
        const dir = win.pendingRecover;
        win.pendingRecover = null;
        try {
          const st = JSON.parse(await fsp.readFile(path.join(dir, 'state.json'), 'utf8'));
          const sources = [];
          for (let i = 0; i < (st.sourceNames || []).length; i++) { const b = await fsp.readFile(path.join(dir, `src-${i}.pdf`)); sources.push(new Uint8Array(b.buffer, b.byteOffset, b.length)); }
          out.recovery = { state: st, sources };
        } catch { /* unreadable */ }
        fs.rm(dir, { recursive: true, force: true }, () => {});
      }
      return out;
    }
    if (win.pendingFile) { out.doc = await readDocument(win.pendingFile); win.pendingFile = null; }
    if (win.pendingDoc) { out.doc = { html: win.pendingDoc.html, name: win.pendingDoc.name, imported: 'pdf' }; win.pendingDoc = null; }
    if (!recoveryOffered) {
      recoveryOffered = true;
      for (const f of leftoverRecovery) {
        try {
          const data = JSON.parse(await fsp.readFile(f, 'utf8'));
          if (!out.recovery && data && typeof data.html === 'string') out.recovery = data;
        } catch { /* unreadable leftover */ }
        fs.rm(f, { force: true }, () => {});
      }
      leftoverRecovery = [];
    }
    return out;
  });

  // several files at once (combine PDFs, pictures to PDF)
  ipcMain.handle('open-many', async (e, kind) => {
    const r = await dialog.showOpenDialog(winOf(e), {
      title: kind === 'images' ? 'Choose pictures' : 'Choose PDF files',
      defaultPath: app.getPath(kind === 'images' ? 'pictures' : 'documents'),
      properties: ['openFile', 'multiSelections'],
      filters: kind === 'images' ? [{ name: 'Pictures', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'] }] : [{ name: 'PDF Files', extensions: ['pdf'] }],
    });
    return r.canceled ? [] : r.filePaths;
  });
  ipcMain.handle('read-image', async (e, p) => {
    const ext = path.extname(String(p)).slice(1).toLowerCase();
    const mime = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' }[ext];
    if (!mime) return { error: 'Not a picture.' };
    try {
      const stat = await fsp.stat(p);
      if (stat.size > 40 * 1024 * 1024) return { error: `${path.basename(p)} is larger than 40 MB.` };
      return { url: `data:${mime};base64,${(await fsp.readFile(p)).toString('base64')}`, name: path.basename(p) };
    } catch (err) { return { error: err.message }; }
  });
  // save several files into one folder the user picks (split PDF, export pages)
  ipcMain.handle('save-files', async (e, d) => {
    const r = await dialog.showOpenDialog(winOf(e), { title: d.title || 'Choose a folder', defaultPath: app.getPath('documents'), properties: ['openDirectory', 'createDirectory'] });
    if (r.canceled || !r.filePaths[0]) return { canceled: true };
    try {
      for (const f of d.files) await writeAtomic(path.join(r.filePaths[0], path.basename(f.name)), Buffer.from(f.bytes));
      return { path: path.join(r.filePaths[0], path.basename(d.files[0].name)), count: d.files.length };
    } catch (err) { return { error: saveErrorMessage(err, r.filePaths[0]) }; }
  });
  // qpdf: AES-256 encryption, decryption with a password, and optimisation
  ipcMain.handle('qpdf', async (e, d) => runQpdf(d));
  // crash recovery for the PDF editor: the document's pages/items plus copies of its source PDFs
  ipcMain.on('pdf-recovery-save', async (e, d) => {
    const win = winOf(e);
    if (!win) return;
    try {
      const dir = pdfRecoveryDir(win);
      await fsp.mkdir(dir, { recursive: true });
      for (const s of d.sources || []) if (s.bytes) await writeAtomic(path.join(dir, `src-${s.index}.pdf`), Buffer.from(s.bytes));
      await writeAtomic(path.join(dir, 'state.json'), JSON.stringify({ ...d.state, time: Date.now() }));
    } catch { /* best effort */ }
  });
  ipcMain.on('pdf-recovery-clear', e => { const win = winOf(e); if (win) fs.rm(pdfRecoveryDir(win), { recursive: true, force: true }, () => {}); });

  ipcMain.handle('open-dialog', async (e, kind) => {
    const filters = kind === 'pdf'
      ? [{ name: 'PDF Files', extensions: ['pdf'] }]
      : kind === 'image'
        ? [{ name: 'Pictures', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'] }]
        : [
          { name: 'All Supported Files', extensions: ['docx', 'pdf', 'html', 'htm', 'txt'] },
          { name: 'Word Documents', extensions: ['docx'] },
          { name: 'PDF Files', extensions: ['pdf'] },
          { name: 'Web Pages', extensions: ['html', 'htm'] },
          { name: 'Text Files', extensions: ['txt'] },
        ];
    const r = await dialog.showOpenDialog(winOf(e), {
      title: kind === 'pdf' ? 'Open PDF' : 'Open',
      defaultPath: app.getPath(kind === 'image' ? 'pictures' : 'documents'),
      properties: ['openFile'],
      filters,
    });
    return r.canceled ? null : r.filePaths[0];
  });

  // ---- PDF editor
  ipcMain.handle('read-pdf', async (e, p) => {
    if (!isPdfPath(p)) return { error: 'That is not a PDF file.' };
    try {
      const bytes = await fsp.readFile(p);
      if (bytes.length > 400 * 1024 * 1024) return { error: 'This PDF is larger than 400 MB, which is too big to edit here.' };
      addRecent(p);
      return { bytes: new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.length), path: p, name: path.basename(p) };
    } catch (err) {
      if (err.code === 'ENOENT') {
        settings.recent = settings.recent.filter(r => r !== p);
        saveSettings();
        return { error: `"${path.basename(p)}" couldn't be found. It may have been moved or deleted.` };
      }
      return { error: `"${path.basename(p)}" couldn't be opened. ${err.message}` };
    }
  });
  ipcMain.handle('read-file-bytes', async (e, p) => {
    // used to insert pages from another PDF (picked by the user in a dialog)
    if (!isPdfPath(p)) return { error: 'That is not a PDF file.' };
    try { const b = await fsp.readFile(p); return { bytes: new Uint8Array(b.buffer, b.byteOffset, b.length), name: path.basename(p) }; } catch (err) { return { error: err.message }; }
  });
  ipcMain.handle('save-pdf', async (e, d, saveAs) => {
    const win = winOf(e);
    let target = d.path;
    if (saveAs || !target) {
      const r = await dialog.showSaveDialog(win, {
        title: d.title || 'Save PDF As',
        defaultPath: path.join(target ? path.dirname(target) : app.getPath('documents'), `${(d.name || 'Document').replace(/\.pdf$/i, '')}.pdf`),
        filters: [{ name: 'PDF', extensions: ['pdf'] }],
      });
      if (r.canceled || !r.filePath) return { canceled: true };
      target = r.filePath;
    }
    if (!isPdfPath(target)) target += '.pdf';
    try { await writeAtomic(target, Buffer.from(d.bytes)); } catch (err) { return { error: saveErrorMessage(err, target) }; }
    if (!d.copy) addRecent(target);
    return { path: target, name: path.basename(target) };
  });
  ipcMain.handle('save-images', async (e, d) => {
    // d.files: [{ name, bytes }] — one file goes to a Save dialog, several to a chosen folder
    const win = winOf(e);
    try {
      if (d.files.length === 1) {
        const r = await dialog.showSaveDialog(win, { title: 'Save Image', defaultPath: path.join(app.getPath('pictures'), d.files[0].name), filters: [{ name: 'PNG Image', extensions: ['png'] }] });
        if (r.canceled || !r.filePath) return { canceled: true };
        await writeAtomic(r.filePath, Buffer.from(d.files[0].bytes));
        return { path: r.filePath, count: 1 };
      }
      const r = await dialog.showOpenDialog(win, { title: 'Choose a folder for the images', defaultPath: app.getPath('pictures'), properties: ['openDirectory', 'createDirectory'] });
      if (r.canceled || !r.filePaths[0]) return { canceled: true };
      for (const f of d.files) await writeAtomic(path.join(r.filePaths[0], path.basename(f.name)), Buffer.from(f.bytes));
      return { path: path.join(r.filePaths[0], path.basename(d.files[0].name)), count: d.files.length };
    } catch (err) {
      return { error: err.message };
    }
  });
  ipcMain.handle('preview-pdf-bytes', async (e, d) => {
    const win = winOf(e);
    return openPdfViewerWindow(win, Buffer.from(d.bytes), d.name);
  });
  // show the other editor in this (empty) window
  ipcMain.handle('switch-editor', (e, kind, file) => {
    const win = winOf(e);
    if (!win || !PAGES[kind]) return false;
    win.pendingFile = typeof file === 'string' ? file : null;
    loadKind(win, kind);
    return true;
  });

  ipcMain.handle('open-path', (e, p) => (typeof p === 'string' ? readDocument(p) : { error: 'Invalid file.' }));

  ipcMain.handle('save', async (e, d, saveAs) => {
    const win = winOf(e);
    let target = d.path;
    if (saveAs || !target) {
      const dir = target ? path.dirname(target) : app.getPath('documents');
      const r = await dialog.showSaveDialog(win, {
        title: 'Save As',
        defaultPath: path.join(dir, `${d.name || 'Document'}.docx`),
        filters: [
          { name: 'Word Document', extensions: ['docx'] },
          { name: 'Web Page', extensions: ['html'] },
          { name: 'Plain Text', extensions: ['txt'] },
        ],
      });
      if (r.canceled || !r.filePath) return { canceled: true };
      target = r.filePath;
    }
    let ext = path.extname(target).toLowerCase();
    if (!['.docx', '.html', '.htm', '.txt'].includes(ext)) { target += '.docx'; ext = '.docx'; }
    try {
      if (ext === '.docx') await writeAtomic(target, await buildDocx(d));
      else if (ext === '.txt') await writeAtomic(target, '﻿' + d.text.replace(/\r?\n/g, '\r\n'));
      else await writeAtomic(target, d.htmlDoc);
    } catch (err) {
      return { error: saveErrorMessage(err, target) };
    }
    addRecent(target);
    return { path: target, name: path.basename(target), format: ext.slice(1) };
  });

  ipcMain.handle('export-pdf', async (e, d) => {
    const win = winOf(e);
    const r = await dialog.showSaveDialog(win, {
      title: 'Export to PDF',
      defaultPath: path.join(d.dir || app.getPath('documents'), `${d.name || 'Document'}.pdf`),
      filters: [{ name: 'PDF', extensions: ['pdf'] }],
    });
    if (r.canceled || !r.filePath) return { canceled: true };
    const hf = !!(d.header || d.footer || d.pageNumbers);
    const box = 'font-family:Calibri,Arial,sans-serif;font-size:9px;color:#444;width:100%;text-align:center;';
    const footer = escHtml(d.footer) + (d.footer && d.pageNumbers ? ' · ' : '') + (d.pageNumbers ? 'Page <span class="pageNumber"></span> of <span class="totalPages"></span>' : '');
    // the window colour shows through the page margins, so print on white
    const restoreBg = whitePaper(win);
    try {
      const data = await e.sender.printToPDF({
        printBackground: true,
        preferCSSPageSize: true,
        displayHeaderFooter: hf,
        headerTemplate: hf ? `<div style="${box}">${escHtml(d.header)}</div>` : '<span></span>',
        footerTemplate: hf ? `<div style="${box}">${footer}</div>` : '<span></span>',
      });
      await writeAtomic(r.filePath, data);
    } catch (err) {
      return { error: saveErrorMessage(err, r.filePath) };
    } finally {
      restoreBg();
    }
    win.lastPdf = r.filePath;
    return { path: r.filePath };
  });

  // exact page preview: render the document to PDF and show it in Chromium's PDF viewer
  ipcMain.handle('print-preview', async (e, d) => {
    const win = winOf(e);
    const hf = !!(d.header || d.footer || d.pageNumbers);
    const box = 'font-family:Calibri,Arial,sans-serif;font-size:9px;color:#444;width:100%;text-align:center;';
    const footer = escHtml(d.footer) + (d.footer && d.pageNumbers ? ' · ' : '') + (d.pageNumbers ? 'Page <span class="pageNumber"></span> of <span class="totalPages"></span>' : '');
    const restoreBg = whitePaper(win);
    let data;
    try {
      data = await e.sender.printToPDF({
        printBackground: true, preferCSSPageSize: true, displayHeaderFooter: hf,
        headerTemplate: hf ? `<div style="${box}">${escHtml(d.header)}</div>` : '<span></span>',
        footerTemplate: hf ? `<div style="${box}">${footer}</div>` : '<span></span>',
      });
    } catch (err) {
      return { error: err.message || String(err) };
    } finally {
      restoreBg();
    }
    return openPdfViewerWindow(win, data, d.name);
  });

  ipcMain.on('open-last-pdf', e => { const win = winOf(e); if (win && win.lastPdf) shell.openPath(win.lastPdf); });
  ipcMain.on('show-in-folder', (e, p) => { if (typeof p === 'string' && fs.existsSync(p)) shell.showItemInFolder(p); });

  ipcMain.handle('print', e => new Promise(resolve => {
    const restoreBg = whitePaper(winOf(e));
    e.sender.print({ printBackground: true }, (success, reason) => { restoreBg(); resolve({ success, reason }); });
  }));

  ipcMain.handle('new-window', (e, opts) => {
    const o = opts && typeof opts === 'object' ? opts : {};
    const doc = o.doc && typeof o.doc.html === 'string' ? { html: o.doc.html, name: String(o.doc.name || 'Converted document') } : null;
    createWindow({ file: typeof o.file === 'string' ? o.file : null, template: typeof o.template === 'string' ? o.template : null, pdf: o.pdf === true, doc });
  });

  ipcMain.handle('pick-image', async e => {
    const r = await dialog.showOpenDialog(winOf(e), {
      title: 'Insert Picture',
      defaultPath: app.getPath('pictures'),
      properties: ['openFile'],
      filters: [{ name: 'Pictures', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'] }],
    });
    if (r.canceled) return null;
    const f = r.filePaths[0];
    const stat = await fsp.stat(f);
    if (stat.size > 25 * 1024 * 1024) return { error: 'That picture is larger than 25 MB. Please choose a smaller one.' };
    const ext = path.extname(f).slice(1).toLowerCase();
    const mime = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' }[ext] || 'image/png';
    return { url: `data:${mime};base64,${(await fsp.readFile(f)).toString('base64')}` };
  });

  ipcMain.handle('confirm-unsaved', async (e, name) => {
    const r = await dialog.showMessageBox(winOf(e), {
      type: 'warning',
      title: APP_NAME,
      message: `Do you want to save your changes to "${name}"?`,
      detail: "Your changes will be lost if you don't save them.",
      buttons: ['Save', "Don't Save", 'Cancel'],
      defaultId: 0,
      cancelId: 2,
      noLink: true,
    });
    return r.response;
  });

  ipcMain.on('close-now', e => {
    const win = winOf(e);
    if (!win) return;
    win.allowClose = true;
    win.close();
  });

  ipcMain.on('ctx-state', (e, s) => { const win = winOf(e); if (win) win.inTable = !!(s && s.inTable); });
  ipcMain.on('open-external', (e, url) => openExternal(url));
  ipcMain.handle('clipboard-text', () => clipboard.readText());
  ipcMain.on('edit', (e, action) => { if (['cut', 'copy', 'paste', 'selectAll'].includes(action)) e.sender[action](); });

  ipcMain.on('recovery-save', async (e, data) => {
    const win = winOf(e);
    if (!win) return;
    try {
      await fsp.mkdir(recoveryDir(), { recursive: true });
      await writeAtomic(win.recoveryFile, JSON.stringify(data));
    } catch { /* best effort */ }
  });
  ipcMain.on('recovery-clear', e => { const win = winOf(e); if (win) fs.rm(win.recoveryFile, { force: true }, () => {}); });

  ipcMain.on('set-setting', (e, key, value) => {
    if (key === 'dark') { settings.dark = !!value; saveSettings(); }
    if ((key === 'userName' || key === 'initials') && typeof value === 'string') { settings[key] = value.slice(0, 60); saveSettings(); }
    // saved signatures: up to 8 small PNG data URLs
    if (key === 'signatures' && Array.isArray(value)) {
      settings.signatures = value.filter(s => typeof s === 'string' && s.startsWith('data:image/png;base64,') && s.length < 2e6).slice(0, 8);
      saveSettings();
    }
  });
  ipcMain.handle('get-recent', () => settings.recent);
  ipcMain.on('remove-recent', (e, p) => { settings.recent = settings.recent.filter(r => r !== p); saveSettings(); });
}

function start() {
  app.on('second-instance', (e, argv) => {
    const file = fileFromArgv(argv);
    if (file) { createWindow({ file }); return; }
    const w = BrowserWindow.getAllWindows()[0];
    if (w) { if (w.isMinimized()) w.restore(); w.focus(); } else createWindow();
  });

  app.whenReady().then(() => {
    migrateOldProfile();
    loadSettings();
    Menu.setApplicationMenu(null);
    try {
      const cutoff = Date.now() - 14 * 24 * 3600 * 1000;
      leftoverRecovery = fs.readdirSync(recoveryDir())
        .filter(f => f.endsWith('.json'))
        .map(f => path.join(recoveryDir(), f))
        .map(f => ({ f, t: fs.statSync(f).mtimeMs }))
        .filter(x => x.t > cutoff || (fs.rmSync(x.f, { force: true }), false))
        .sort((a, b) => b.t - a.t)
        .map(x => x.f);
    } catch { leftoverRecovery = []; }

    // Only allow the permissions a word processor needs (reading installed font names).
    session.defaultSession.setPermissionRequestHandler((wc, perm, cb) => cb(perm === 'local-fonts'));
    session.defaultSession.setPermissionCheckHandler((wc, perm) => perm === 'local-fonts');

    registerIpc();
    createWindow({ file: fileFromArgv(process.argv) });
    // PDFs that were being edited when the app closed unexpectedly
    try {
      const dirs = fs.readdirSync(recoveryDir()).filter(f => f.startsWith('pdf-')).map(f => path.join(recoveryDir(), f))
        .filter(d => fs.existsSync(path.join(d, 'state.json')));
      for (const d of dirs.slice(0, 3)) createWindow({ pdf: true, recover: d });
      for (const d of dirs.slice(3)) fs.rmSync(d, { recursive: true, force: true });
    } catch { /* none */ }
  });

  app.on('window-all-closed', () => app.quit());
}

module.exports = { createWindow };
