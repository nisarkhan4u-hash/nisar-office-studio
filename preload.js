'use strict';
const { contextBridge, ipcRenderer, webUtils } = require('electron');

const EVENTS = ['request-close', 'cmd'];

contextBridge.exposeInMainWorld('api', {
  init: () => ipcRenderer.invoke('get-init'),
  openDialog: kind => ipcRenderer.invoke('open-dialog', kind || 'all'),
  readPdf: p => ipcRenderer.invoke('read-pdf', p),
  openMany: kind => ipcRenderer.invoke('open-many', kind),
  readImage: p => ipcRenderer.invoke('read-image', p),
  saveFiles: data => ipcRenderer.invoke('save-files', data),
  qpdf: data => ipcRenderer.invoke('qpdf', data),
  pdfRecoverySave: d => ipcRenderer.send('pdf-recovery-save', d),
  pdfRecoveryClear: () => ipcRenderer.send('pdf-recovery-clear'),
  readFileBytes: p => ipcRenderer.invoke('read-file-bytes', p),
  savePdf: (data, saveAs) => ipcRenderer.invoke('save-pdf', data, !!saveAs),
  saveImages: data => ipcRenderer.invoke('save-images', data),
  previewPdfBytes: data => ipcRenderer.invoke('preview-pdf-bytes', data),
  switchEditor: (kind, file) => ipcRenderer.invoke('switch-editor', kind, file),
  openPath: p => ipcRenderer.invoke('open-path', p),
  save: (data, saveAs) => ipcRenderer.invoke('save', data, !!saveAs),
  exportPdf: data => ipcRenderer.invoke('export-pdf', data),
  openLastPdf: () => ipcRenderer.send('open-last-pdf'),
  showInFolder: p => ipcRenderer.send('show-in-folder', p),
  print: () => ipcRenderer.invoke('print'),
  printPreview: data => ipcRenderer.invoke('print-preview', data),
  newWindow: opts => ipcRenderer.invoke('new-window', opts || {}),
  pickImage: () => ipcRenderer.invoke('pick-image'),
  confirmUnsaved: name => ipcRenderer.invoke('confirm-unsaved', name),
  closeNow: () => ipcRenderer.send('close-now'),
  setCtx: s => ipcRenderer.send('ctx-state', s),
  openExternal: url => ipcRenderer.send('open-external', url),
  clipboardText: () => ipcRenderer.invoke('clipboard-text'),
  edit: action => ipcRenderer.send('edit', action),
  recoverySave: d => ipcRenderer.send('recovery-save', d),
  recoveryClear: () => ipcRenderer.send('recovery-clear'),
  setSetting: (k, v) => ipcRenderer.send('set-setting', k, v),
  getRecent: () => ipcRenderer.invoke('get-recent'),
  removeRecent: p => ipcRenderer.send('remove-recent', p),
  pathForFile: f => { try { return webUtils.getPathForFile(f); } catch { return ''; } },
  on: (channel, fn) => {
    if (EVENTS.includes(channel)) ipcRenderer.on(channel, (e, ...args) => fn(...args));
  },
});
