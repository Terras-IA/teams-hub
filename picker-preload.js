const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('picker', {
  onFontes: (cb) => ipcRenderer.on('picker:fontes', (_e, dados) => cb(dados)),
  escolher: (id) => ipcRenderer.send('picker:escolha', id),
});
