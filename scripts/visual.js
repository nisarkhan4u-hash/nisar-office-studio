'use strict';
// Visual walkthrough: drives every feature and saves screenshots for review.
// Run: electron scripts/visual.js <outDir>
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const out = process.argv.find(a => a.startsWith('--out='))?.slice(6) || path.join(os.tmpdir(), 'nw-visual');
fs.mkdirSync(out, { recursive: true });
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'nw-visual-profile-'));
app.setPath('userData', path.join(tmp, 'profile'));
app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion');
app.commandLine.appendSwitch('disable-renderer-backgrounding');
require('../main.js');

const wait = ms => new Promise(r => setTimeout(r, ms));
const notes = [];
const errors = [];

let started = false;
app.on('browser-window-created', (e, win) => {
  if (started) return;
  started = true;
  win.webContents.on('console-message', ev => { if (ev.level === 'error' || ev.level === 'warning') errors.push(`${ev.level}: ${ev.message}`); });
  win.webContents.once('did-finish-load', () => run(win).catch(err => { notes.push(`CRASH ${err.stack}`); done(); }));
});

async function run(win) {
  const wc = win.webContents;
  await wait(1200);
  win.unmaximize();
  win.setBounds({ x: 40, y: 40, width: 1500, height: 940 });
  await wait(600);
  wc.setBackgroundThrottling(false);
  const js = c => wc.executeJavaScript(c, true);
  const type = async t => { wc.insertText(t); await wait(200); };
  let n = 0;
  const shot = async (name, settle = 250) => {
    await wait(settle);
    const img = await wc.capturePage();
    fs.writeFileSync(path.join(out, `${String(++n).padStart(2, '0')}-${name}.png`), img.toPNG());
  };
  const note = async (label, code) => notes.push(`${label}: ${JSON.stringify(await js(code))}`);

  await note('dirty at startup', 'state.dirty');
  await note('title at startup', 'document.title');
  await shot('startup');

  // sample content typed through the real input pipeline
  await js('editor.focus(); placeCaret(editor.firstElementChild, true); true');
  await type('Quarterly Report');
  await js('applyStyle("title"); document.execCommand("insertParagraph"); applyStyle("normal"); true');
  await type('This document shows formatting. Some words are misspeled on purpose. ');
  await js(`(() => { const p = editor.lastElementChild; const t = p.firstChild; const r = document.createRange(); r.setStart(t, 5); r.setEnd(t, 13); setRange(r); runCommand('bold'); r.setStart(p.lastChild, 1); return true; })()`);
  await js('placeCaret(editor.lastElementChild, false); document.execCommand("insertParagraph"); applyStyle("h1"); true');
  await type('Heading One');
  await js('document.execCommand("insertParagraph"); applyStyle("normal"); true');
  await type('A normal paragraph after a heading. ');
  await js('runCommand("bullets"); true');
  await type('First bullet');
  await js('document.execCommand("insertParagraph"); true');
  await type('Second bullet');
  await js('document.execCommand("insertParagraph"); document.execCommand("insertParagraph"); true');
  await note('after list exit', 'editor.lastElementChild.outerHTML');
  await type('Paragraph after list.');
  await shot('typed-content');
  await note('dirty after typing', 'state.dirty');

  // home menus
  for (const m of ['sizes', 'case', 'fore', 'hilite', 'shade', 'bulletStyle', 'numberStyle', 'spacing']) {
    await js(`openMenu('${m}', document.querySelector('[data-menu="${m}"]')); true`);
    await shot(`menu-${m}`, 150);
    await js('closeMenu(); true');
  }
  await js('document.getElementById("font-family").focus(); true');
  await shot('font-select-focus');
  await js('editor.focus(); true');

  // insert tab
  await js('document.querySelector(\'[data-tab="insert"]\').click(); true');
  await shot('tab-insert');
  await js(`openMenu('table', document.querySelector('[data-menu="table"]')); (() => { const c = document.querySelector('.tgrid [data-r="3"][data-c="4"]'); c.dispatchEvent(new MouseEvent('mouseover', { bubbles: true })); })(); true`);
  await shot('menu-table');
  await js('closeMenu(); placeCaret(editor.lastElementChild, false); document.execCommand("insertParagraph"); insertTable(3, 4); true');
  await type('Name');
  await js('handleTab(false); true');
  await type('Value');
  await js('handleTab(false); true');
  await type('A much longer cell text that wraps across lines');
  await shot('table');
  for (const m of ['symbols', 'date']) {
    await js(`openMenu('${m}', document.querySelector('[data-menu="${m}"]')); true`);
    await shot(`menu-${m}`, 150);
    await js('closeMenu(); true');
  }

  // layout tab with table tools active
  await js('document.querySelector(\'[data-tab="layout"]\').click(); true');
  await shot('tab-layout-in-table');
  for (const m of ['margins', 'orientation', 'size', 'columns']) {
    await js(`openMenu('${m}', document.querySelector('[data-menu="${m}"]')); true`);
    await shot(`menu-${m}`, 150);
    await js('closeMenu(); true');
  }
  await js('document.querySelector(\'[data-tab="review"]\').click(); true');
  await shot('tab-review');
  await js('document.querySelector(\'[data-tab="view"]\').click(); true');
  await shot('tab-view');
  await js('document.querySelector(\'[data-tab="home"]\').click(); true');

  // picture + toolbar
  const png = fs.readFileSync(path.join(__dirname, '..', 'build', 'icon.png')).toString('base64');
  await js(`placeCaret(editor.lastElementChild, false); document.execCommand('insertParagraph'); insertImage('data:image/png;base64,${png}'); true`);
  await wait(300);
  await js('const im = editor.querySelector("img"); im.scrollIntoView({block:"center"}); selectImage(im); true');
  await shot('image-selected');
  await js('imageAction("25"); imageAction("floatRight"); true');
  await shot('image-small-float');
  await js('selectImage(null); true');

  // find & replace
  await js('openFind(true); findInput.value = "bullet"; refreshFind(false); true');
  await shot('find-replace');
  await js('closeFind(); true');

  // dialogs
  for (const [name, code] of [['dlg-link', 'insertLink()'], ['dlg-headerfooter', 'editHeaderFooter()'], ['dlg-wordcount', 'showWordCount()'], ['dlg-margins', 'customMargins()'], ['dlg-zoom', 'zoomDialog()'], ['dlg-goto', 'gotoPage()']]) {
    await js(`${code}; true`);
    await shot(name);
    await js('document.getElementById("modal").close(); true');
  }

  // multi-page, header/footer, page numbers
  await js(`(() => { for (let i = 0; i < 40; i++) { const p = document.createElement('p'); p.textContent = 'Filler paragraph number ' + (i + 1) + '. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.'; editor.appendChild(p); } state.page.header = 'Company Confidential'; state.page.footer = 'Quarterly Report'; state.page.pageNumbers = true; applyPage(); markDirty(); return true; })()`);
  await wait(400);
  await js('scroller.scrollTop = 900; true');
  await shot('multipage-boundary');
  await note('pages', 'state.pageCount + " / status: " + document.getElementById("st-page").textContent');
  await js('fitOnePage(); scroller.scrollTop = 0; true');
  await shot('one-page-zoom');
  await js('setZoom(1); setPage({ orientation: "landscape" }); setPage({ columns: 2 }); scroller.scrollTop = 0; true');
  await shot('landscape-2col');
  await js('setPage({ orientation: "portrait", columns: 1 }); true');

  // pdf
  win.setBackgroundColor('#ffffff');
  const pdf = await wc.printToPDF({ printBackground: true, preferCSSPageSize: true, displayHeaderFooter: true, headerTemplate: '<div style="font-family:Calibri,Arial;font-size:9px;color:#444;width:100%;text-align:center">Company Confidential</div>', footerTemplate: '<div style="font-family:Calibri,Arial;font-size:9px;color:#444;width:100%;text-align:center">Quarterly Report · Page <span class="pageNumber"></span> of <span class="totalPages"></span></div>' });
  fs.writeFileSync(path.join(out, 'export.pdf'), pdf);
  {
    const pv = new BrowserWindow({ width: 1100, height: 1300, show: true, webPreferences: { plugins: true } });
    await pv.loadFile(path.join(out, 'export.pdf'));
    await wait(2500);
    fs.writeFileSync(path.join(out, `${String(++n).padStart(2, '0')}-pdf-view.png`), (await pv.webContents.capturePage()).toPNG());
    pv.destroy();
  }
  // typing at the bottom of a page moves the paragraph onto the next page
  await js('setZoom(1); (() => { const ps = [...editor.querySelectorAll("p")]; const target = ps.find(p => p.hasAttribute("data-nw-push")); if (target) { target.scrollIntoView({ block: "center" }); placeCaret(target.previousElementSibling, false); } return !!target; })()');
  await shot('page-gap');
  await type(' Extra words typed at the end of the last paragraph on page one so that it needs another line and must move.');
  await shot('page-gap-after-typing', 400);
  await note('pushed blocks', 'editor.querySelectorAll("[data-nw-push]").length');
  await note('save html has no layout attrs', '!cleanClone().innerHTML.includes("nw-push")');

  // views
  await js('setDark(true); scroller.scrollTop = 0; true');
  await shot('dark');
  await js(`openMenu('fore', document.querySelector('[data-menu="fore"]')); true`);
  await shot('dark-menu');
  await js('closeMenu(); showBackstage("home"); true');
  await shot('dark-backstage');
  await js('hideBackstage(); setDark(false); setWebLayout(true); true');
  await shot('web-layout');
  await js('setWebLayout(false); setFocusMode(true); true');
  await shot('focus-mode');
  await js('setFocusMode(false); runCommand("collapseRibbon"); true');
  await shot('ribbon-collapsed');
  await js('runCommand("collapseRibbon"); showBackstage("home"); true');
  await shot('backstage-home');
  await js('switchPane("new"); true');
  await shot('backstage-new');
  await js('switchPane("about"); true');
  await shot('backstage-about');
  await js('hideBackstage(); true');

  // templates
  for (const t of ['letter', 'resume', 'report', 'notes']) {
    await js(`editor.innerHTML = TEMPLATES['${t}'].html(); state.page = defaultPage(); applyPage(); scroller.scrollTop = 0; true`);
    await shot(`template-${t}`);
  }

  // narrow window
  win.setBounds({ x: 40, y: 40, width: 900, height: 700 });
  await shot('narrow-window', 600);

  done();
}

function done() {
  console.log(notes.join('\n'));
  if (errors.length) console.log('CONSOLE:\n' + errors.join('\n'));
  console.log('screenshots in', out);
  app.exit(0);
}

