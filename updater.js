// Atualização pelo GitHub Releases (Terras-IA/teams-hub). Nada baixa nem
// instala sem o usuário aceitar: o app avisa, mostra as novidades e só então
// baixa; a instalação do .deb pede a senha do sistema (pkexec).
const { app, ipcMain } = require('electron');
const fs = require('fs');
const path = require('path');

const VERIFICAR_A_CADA = 6 * 60 * 60_000;

// As notas da release chegam em HTML; a interface mostra só texto.
function notasEmTexto(notas) {
  const html = Array.isArray(notas) ? notas.map((n) => `<h3>${n.version}</h3>${n.note || ''}`).join('') : String(notas || '');
  return html
    .replace(/<li[^>]*>/gi, '• ')
    .replace(/<\/(p|li|h\d|ul|ol)>|<br\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function registrar({ config, saveConfig, logDiag, getWin, notificar }) {
  config.atualizacoes = { automatico: true, pular: null, ...config.atualizacoes };
  const cfg = () => config.atualizacoes;
  let estado = { estado: app.isPackaged ? 'parado' : 'dev', versaoAtual: app.getVersion() };
  let updater = null;

  const publicar = (mudanca) => {
    estado = { ...estado, ...mudanca };
    const w = getWin();
    if (w && !w.isDestroyed()) w.webContents.send('upd:estado', estado);
  };

  ipcMain.handle('upd:estado', () => ({ ...estado, automatico: cfg().automatico, pular: cfg().pular }));
  ipcMain.handle('upd:changelog', () => {
    try { return fs.readFileSync(path.join(__dirname, 'CHANGELOG.md'), 'utf8'); } catch { return ''; }
  });
  ipcMain.handle('upd:auto', (_e, ligado) => {
    cfg().automatico = !!ligado;
    saveConfig();
    return cfg().automatico;
  });

  // Rodando do código-fonte não há pacote para trocar: atualiza com git pull.
  if (!app.isPackaged) {
    for (const canal of ['upd:verificar', 'upd:baixar', 'upd:instalar', 'upd:pular']) ipcMain.handle(canal, () => estado);
    return;
  }

  ({ autoUpdater: updater } = require('electron-updater'));
  updater.autoDownload = false;
  updater.autoInstallOnAppQuit = false;
  updater.logger = { info: () => {}, warn: (m) => logDiag(`[atualização] ${m}`), error: (m) => logDiag(`[atualização] ${m}`), debug: () => {} };

  let manual = false;
  updater.on('checking-for-update', () => publicar({ estado: 'verificando', erro: null }));
  updater.on('update-not-available', () => publicar({ estado: 'em-dia', verificadoEm: new Date().toISOString() }));
  updater.on('update-available', (info) => {
    publicar({ estado: 'disponivel', versao: info.version, notas: notasEmTexto(info.releaseNotes), data: info.releaseDate, verificadoEm: new Date().toISOString() });
    // Versão pulada só reaparece quando a verificação é pedida à mão.
    if (!manual && cfg().pular === info.version) return;
    publicar({ perguntar: true });
    notificar({ tab: 'CONTAS', title: `Teams Hub ${info.version} disponível`, body: 'Clique para ver as novidades e atualizar.' });
  });
  updater.on('download-progress', (p) => publicar({ estado: 'baixando', progresso: Math.round(p.percent) }));
  updater.on('update-downloaded', () => publicar({ estado: 'pronto', progresso: 100 }));
  updater.on('error', (err) => {
    logDiag(`[atualização] erro: ${err.message}`);
    publicar({ estado: 'erro', erro: String(err.message).slice(0, 300) });
  });

  const verificar = async (pedido) => {
    manual = pedido;
    publicar({ perguntar: false });
    try { await updater.checkForUpdates(); } catch { /* o evento error já avisou */ }
    return estado;
  };

  ipcMain.handle('upd:verificar', () => verificar(true));
  ipcMain.handle('upd:baixar', async () => {
    publicar({ estado: 'baixando', progresso: 0, perguntar: false });
    try { await updater.downloadUpdate(); } catch { /* idem */ }
    return estado;
  });
  ipcMain.handle('upd:instalar', () => {
    if (estado.estado !== 'pronto') return estado;
    logDiag(`[atualização] instalando ${estado.versao}`);
    // Fecha de verdade (sem ir para a bandeja), instala e abre de novo.
    setImmediate(() => updater.quitAndInstall(false, true));
    return estado;
  });
  ipcMain.handle('upd:pular', () => {
    cfg().pular = estado.versao || null;
    saveConfig();
    publicar({ perguntar: false });
    return estado;
  });

  app.whenReady().then(() => {
    setTimeout(() => cfg().automatico && verificar(false), 30_000);
    setInterval(() => cfg().automatico && estado.estado !== 'baixando' && estado.estado !== 'pronto' && verificar(false), VERIFICAR_A_CADA);
  });
}

module.exports = { registrar, notasEmTexto };
