'use strict';
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('og', {
  async call(name, ...args) {
    const r = await ipcRenderer.invoke('api', name, ...args);
    if (!r.ok) { const e = new Error(r.error); e.extra = r.extra; throw e; }
    return r.value;
  },
  onFsChanged(cb) { ipcRenderer.on('fs-changed', () => cb()); },
});
