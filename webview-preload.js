// Roda dentro de cada Teams. Troca a Notification da página por uma que avisa
// a janela principal, para o aviso nativo dizer de qual conta veio e, ao
// clicar, abrir a aba certa.
const { ipcRenderer, webFrame } = require('electron');

window.addEventListener('message', (e) => {
  if (e.source !== window || !e.data || e.data.__hub !== 'notif') return;
  ipcRenderer.sendToHost('notif', { title: String(e.data.title || ''), body: String(e.data.body || '') });
});

// Outlook: o título não traz as não lidas. O item "Caixa de Entrada" da árvore
// de pastas traz, no texto acessível ("261 não lidos" / "261 unread").
if (/(^|\.)outlook\.(office|office365|live)\.com$|^outlook\.cloud\.microsoft$/.test(location.hostname)) {
  let ultimo = -1;
  const contar = () => {
    const nome = [...document.querySelectorAll('[role=treeitem] span')]
      .find((s) => s.childElementCount === 0 && /^(Caixa de Entrada|Inbox)$/.test(s.textContent.trim()));
    const item = nome?.closest('[role=treeitem]');
    if (!item) return;
    const m = /(\d+)\s+(não lid|unread)/i.exec(item.innerText || '');
    const n = m ? Number(m[1]) : 0;
    if (n !== ultimo) { ultimo = n; ipcRenderer.sendToHost('nao-lidas', n); }
  };
  window.addEventListener('load', () => { setTimeout(contar, 5000); setInterval(contar, 30000); });
}

webFrame.executeJavaScript(`(() => {
  const post = (title, opts) => window.postMessage({ __hub: 'notif', title, body: (opts && opts.body) || '' }, '*');
  class HubNotification extends EventTarget {
    constructor(title, opts) { super(); this.title = title; Object.assign(this, opts || {}); post(title, opts); }
    close() {}
    static get permission() { return 'granted'; }
    static requestPermission(cb) { cb && cb('granted'); return Promise.resolve('granted'); }
  }
  window.Notification = HubNotification;
  if (window.ServiceWorkerRegistration) {
    ServiceWorkerRegistration.prototype.showNotification = function (title, opts) { post(title, opts); return Promise.resolve(); };
  }
})();`);
