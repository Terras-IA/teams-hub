const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('hub', {
  config: () => ipcRenderer.invoke('hub:config'),
  cores: () => ipcRenderer.invoke('hub:cores'),
  agenda: (dias) => ipcRenderer.invoke('hub:agenda', dias),
  salvarContas: (contas) => ipcRenderer.invoke('hub:contas:salvar', contas),
  espelhar: () => ipcRenderer.invoke('hub:espelhar'),
  google: {
    listar: () => ipcRenderer.invoke('hub:google:listar'),
    salvar: (fonte) => ipcRenderer.invoke('hub:google:salvar', fonte),
    autorizar: (nome, escrita) => ipcRenderer.invoke('hub:google:autorizar', nome, escrita),
    calendarios: (nome) => ipcRenderer.invoke('hub:google:calendarios', nome),
    remover: (nome) => ipcRenderer.invoke('hub:google:remover', nome),
  },
  badges: (porConta) => ipcRenderer.send('hub:badges', porConta),
  notify: (msg) => ipcRenderer.send('hub:notify', msg),
  openExternal: (url) => ipcRenderer.send('hub:open-external', url),
  onFocusTab: (cb) => ipcRenderer.on('hub:focus-tab', (_e, tab) => cb(tab)),
});
