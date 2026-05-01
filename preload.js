
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  selectTemplate: () => ipcRenderer.invoke('select-template'),
  openExcel: () => ipcRenderer.invoke('open-excel'),
  validateSchema: (tplPath) => ipcRenderer.invoke('validate-schema', tplPath),
  mergeDocs: (tpl, groupingCol) => ipcRenderer.invoke('merge-docs', tpl, groupingCol),
  onProgress: cb => ipcRenderer.on('progress', (_, p) => cb(p))
});
