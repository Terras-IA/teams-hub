// Teams Hub: as contas do Teams lado a lado, cada uma numa sessão isolada,
// mais a agenda unificada do terras-agenda. Fechar a janela só esconde (fica
// na bandeja) para os avisos continuarem chegando.
const { app, BrowserWindow, Tray, Menu, Notification, ipcMain, session, shell, nativeImage, desktopCapturer } = require('electron');
const { execFile } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const UA = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36';
const CONFIG_DIR = path.join(os.homedir(), '.config', 'teams-hub');
const CONFIG_FILE = path.join(CONFIG_DIR, 'config.json');

const DEFAULT_CONFIG = {
  // Começa vazio: as contas são cadastradas na aba Contas.
  contas: [],
  extras: [
    { id: 'GCAL', nome: 'Google Agenda', cor: '#db4437', url: 'https://calendar.google.com/calendar/r/week', externo: true },
  ],
  agenda: {
    script: path.join(os.homedir(), '.claude', 'skills', 'terras-agenda', 'scripts', 'terras_agenda.py'),
    dias: 7,
    lembrete_min: 5,
  },
};

function loadConfig() {
  fs.mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
  if (!fs.existsSync(CONFIG_FILE)) {
    fs.writeFileSync(CONFIG_FILE, JSON.stringify(DEFAULT_CONFIG, null, 2), { mode: 0o600 });
    return DEFAULT_CONFIG;
  }
  return { ...DEFAULT_CONFIG, ...JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8')) };
}

function saveConfig() {
  const tmp = CONFIG_FILE + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(config, null, 2) + '\n', { mode: 0o600 });
  fs.renameSync(tmp, CONFIG_FILE);
}

const config = loadConfig();
const AGENDA_CONFIG = path.join(os.homedir(), '.config', 'terras-agenda', 'config.json');
let win = null;
let tray = null;
let quitting = false;

if (!app.requestSingleInstanceLock()) app.quit();
app.on('second-instance', () => showWindow());
app.userAgentFallback = UA;
app.setName('Teams Hub');

const icon = (name) => nativeImage.createFromPath(path.join(__dirname, 'icons', name));

function showWindow(tab) {
  if (!win) return;
  win.show();
  win.focus();
  if (tab) win.webContents.send('hub:focus-tab', tab);
}

// Domínios que ficam dentro do app (login, reunião em janela própria). O Google
// fica fora: ele recusa login em navegador embutido.
const INTERNAL = /(^|\.)(microsoft\.com|microsoftonline\.com|office\.com|office365\.com|live\.com|skype\.com|teams\.cdn\.office\.net|sharepoint\.com|cloud\.microsoft)$/;

// Diagnóstico: permissões recusadas e erros das páginas vão para este arquivo.
const LOG_FILE = path.join(os.homedir(), '.cache', 'teams-hub', 'diagnostico.log');
function logDiag(linha) {
  try {
    fs.mkdirSync(path.dirname(LOG_FILE), { recursive: true });
    if (fs.existsSync(LOG_FILE) && fs.statSync(LOG_FILE).size > 2 * 1024 * 1024) fs.renameSync(LOG_FILE, LOG_FILE + '.1');
    fs.appendFileSync(LOG_FILE, `${new Date().toISOString()} ${linha}\n`);
  } catch { /* log nunca derruba o app */ }
}

// Compartilhar tela: o Electron não tem seletor próprio no Linux, então o
// pedido de getDisplayMedia abre esta janela com as telas e janelas.
let pickerAberto = null;
async function escolherFonte(conta) {
  if (pickerAberto) { pickerAberto.focus(); return null; }
  const fontes = await desktopCapturer.getSources({ types: ['screen', 'window'], thumbnailSize: { width: 400, height: 225 } });
  const proprias = new Set(BrowserWindow.getAllWindows().map((w) => w.getMediaSourceId()));
  const lista = fontes
    .filter((f) => !proprias.has(f.id))
    .map((f) => ({
      id: f.id,
      nome: f.name.replace(/^Screen (\d+)$/, 'Tela $1').replace(/^Entire screen$/, 'Tela inteira'),
      tipo: f.id.startsWith('screen:') ? 'screen' : 'window',
      miniatura: f.thumbnail.toDataURL(),
    }));
  return new Promise((resolve) => {
    const janela = new BrowserWindow({
      parent: win, modal: true, width: 900, height: 620, title: 'Compartilhar tela', autoHideMenuBar: true,
      webPreferences: { preload: path.join(__dirname, 'picker-preload.js'), contextIsolation: true },
    });
    pickerAberto = janela;
    let feito = false;
    const fim = (id) => {
      if (feito) return;
      feito = true;
      ipcMain.removeListener('picker:escolha', ouvir);
      pickerAberto = null;
      if (!janela.isDestroyed()) janela.close();
      resolve(fontes.find((f) => f.id === id) || null);
    };
    const ouvir = (e, id) => { if (e.sender === janela.webContents) fim(id); };
    ipcMain.on('picker:escolha', ouvir);
    janela.on('closed', () => fim(null));
    janela.webContents.once('did-finish-load', () => janela.webContents.send('picker:fontes', { lista, conta }));
    janela.loadFile(path.join(__dirname, 'renderer', 'picker.html'));
  });
}

function prepareSession(partition) {
  const s = session.fromPartition(partition);
  s.setUserAgent(UA);
  const conta = partition.replace(/^persist:/, '');
  s.setDisplayMediaRequestHandler(async (_req, callback) => {
    try {
      const fonte = await escolherFonte(conta);
      logDiag(`[${partition}] compartilhar tela: ${fonte ? fonte.name : 'cancelado'}`);
      callback(fonte ? { video: fonte } : {});
    } catch (err) {
      logDiag(`[${partition}] compartilhar tela falhou: ${err.message}`);
      callback({});
    }
  });
  const allowed = new Set(['notifications', 'media', 'display-capture', 'clipboard-read', 'clipboard-sanitized-write', 'fullscreen', 'screen-wake-lock']);
  s.setPermissionRequestHandler((_wc, perm, cb, det) => {
    if (!allowed.has(perm)) logDiag(`[${partition}] permissão PEDIDA e negada: ${perm} ${det?.requestingUrl || ''}`);
    cb(allowed.has(perm));
  });
  s.setPermissionCheckHandler((_wc, perm, origem) => {
    if (!allowed.has(perm)) logDiag(`[${partition}] permissão CONSULTADA e negada: ${perm} ${origem || ''}`);
    return allowed.has(perm);
  });
}

function createWindow() {
  win = new BrowserWindow({
    width: 1400,
    height: 900,
    title: 'Teams Hub',
    icon: path.join(__dirname, 'icons', 'icon.png'),
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      webviewTag: true,
      contextIsolation: true,
    },
  });
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  win.on('close', (e) => {
    if (!quitting) { e.preventDefault(); win.hide(); }
  });
}

app.on('web-contents-created', (_e, wc) => {
  if (wc.getType() !== 'webview') return;
  const quem = () => { try { return new URL(wc.getURL()).hostname; } catch { return '?'; } };
  wc.on('console-message', (ev) => {
    const nivel = ev.level ?? ev.params?.level;
    if (nivel === 'error' || nivel === 3) logDiag(`[${quem()}] console erro: ${String(ev.message ?? ev.params?.message).slice(0, 500)}`);
  });
  wc.on('render-process-gone', (_ev, det) => logDiag(`[${quem()}] processo da página caiu: ${det.reason}`));
  wc.setWindowOpenHandler(({ url }) => {
    let host = '';
    try { host = new URL(url).hostname; } catch { /* about:blank etc. */ }
    logDiag(`[${quem()}] abrir janela: ${url.slice(0, 200)}`);
    if (url === 'about:blank' || INTERNAL.test(host)) {
      return {
        action: 'allow',
        overrideBrowserWindowOptions: {
          autoHideMenuBar: true,
          webPreferences: { session: wc.session },
        },
      };
    }
    shell.openExternal(url);
    return { action: 'deny' };
  });
});

function updateTray(total, porConta) {
  if (!tray) return;
  tray.setImage(icon(total > 0 ? 'tray-unread.png' : 'tray.png'));
  const linha = (filtro) => Object.entries(porConta).filter(([k]) => filtro(k)).map(([id, n]) => `${id.replace(' email', '')}: ${n}`).join('  ');
  const teams = linha((k) => !k.endsWith(' email'));
  const email = linha((k) => k.endsWith(' email'));
  tray.setToolTip(`Teams Hub — ${total > 0 ? `${total} mensagem(ns) no Teams` : 'nada novo no Teams'}`
    + (teams ? `\nTeams  ${teams}` : '') + (email ? `\nEmail  ${email}` : ''));
  app.setBadgeCount(total);
}

function trayMenu() {
  const abas = [{ id: 'AGENDA', nome: 'Agenda' }, ...config.contas, ...config.extras, { id: 'CONTAS', nome: 'Contas…' }];
  tray.setContextMenu(Menu.buildFromTemplate([
    ...abas.map((c) => ({ label: c.nome, click: () => showWindow(c.id) })),
    { type: 'separator' },
    { label: 'Sair', click: () => { quitting = true; app.quit(); } },
  ]));
}

function buildTray() {
  tray = new Tray(icon('tray.png'));
  trayMenu();
  tray.on('click', () => (win.isVisible() && win.isFocused() ? win.hide() : showWindow()));
  updateTray(0, {});
}

// ---------- terras-agenda (agendas Google) ----------

function agendaCli(args, timeout = 90_000) {
  return new Promise((resolve) => {
    const script = config.agenda.script;
    if (!fs.existsSync(script)) return resolve({ ok: false, saida: `terras-agenda não encontrado em ${script}` });
    execFile('python3', [script, ...args], { timeout, maxBuffer: 10 * 1024 * 1024 }, (err, stdout, stderr) => {
      resolve({ ok: !err, saida: `${stdout || ''}${stderr || ''}`.trim() || err?.message || '' });
    });
  });
}

function agendaConfig() {
  try { return JSON.parse(fs.readFileSync(AGENDA_CONFIG, 'utf8')); } catch { return {}; }
}

// Lista do "google cals": "   * Nome do calendário   id". O id nunca tem espaço.
function parseCalendarios(saida) {
  const cals = [];
  let dentro = false;
  for (const linha of saida.split('\n')) {
    if (/^\s*calendários de /.test(linha)) { dentro = true; continue; }
    if (dentro && /^\s*Para trocar:/.test(linha)) break;
    const m = dentro && /^\s{2,}([* ])\s(.+?)\s{2,}(\S+)\s*$/.exec(linha);
    if (m) cals.push({ principal: m[1] === '*', nome: m[2].trim(), id: m[3] });
  }
  return cals;
}

ipcMain.handle('hub:config', () => ({
  ...config,
  // A aba Agenda depende do terras-agenda; sem ele, o Hub funciona só com as contas.
  agendaDisponivel: fs.existsSync(config.agenda.script),
  webviewPreload: 'file://' + path.join(__dirname, 'webview-preload.js'),
}));

ipcMain.handle('hub:agenda', (_e, dias) => new Promise((resolve) => {
  const script = config.agenda.script;
  if (!fs.existsSync(script)) return resolve({ erro: `terras-agenda não encontrado em ${script}` });
  execFile('python3', [script, 'agenda', '--dias', String(dias || config.agenda.dias), '--json'],
    { timeout: 90_000, maxBuffer: 10 * 1024 * 1024 },
    (err, stdout, stderr) => {
      try { resolve(JSON.parse(stdout)); } catch {
        resolve({ erro: (stderr || err?.message || 'saída inválida do terras-agenda').slice(0, 800) });
      }
    });
}));

// Cores por origem da agenda: contas e abas do hub + o que estiver no terras-agenda.
ipcMain.handle('hub:cores', () => {
  const ta = agendaConfig();
  const cores = {};
  for (const f of [...(ta.contas || []), ...(ta.feeds || [])]) if (f.cor) cores[f.nome] = f.cor;
  for (const g of ta.google || []) cores[g.nome] = g.cor || '#db4437';
  for (const c of [...config.contas, ...config.extras]) cores[c.id] = c.cor;
  return cores;
});

// Contas do Teams: a interface manda a lista inteira; quem saiu perde o login salvo.
ipcMain.handle('hub:contas:salvar', async (_e, contas) => {
  const novos = new Set(contas.map((c) => c.id));
  const removidos = config.contas.filter((c) => !novos.has(c.id)).map((c) => c.id);
  for (const c of contas) prepareSession(`persist:${c.id}`);
  config.contas = contas;
  saveConfig();
  trayMenu();
  for (const id of removidos) await session.fromPartition(`persist:${id}`).clearStorageData();
  return { ok: true, removidos };
});

ipcMain.handle('hub:google:listar', async () => {
  const ta = agendaConfig();
  const tokens = path.join(os.homedir(), '.config', 'terras-agenda');
  const slug = (s) => s.normalize('NFKD').replace(/[^\x00-\x7f]/g, '').replace(/[^a-zA-Z0-9]+/g, '-').replace(/^-|-$/g, '').toLowerCase();
  return (ta.google || []).map((g) => ({
    ...g,
    autorizada: fs.existsSync(path.join(tokens, `token-${slug('google-' + g.nome)}.json`)),
  }));
});

ipcMain.handle('hub:google:salvar', async (_e, { nome, conta, cor, calendarios }) => {
  const args = ['google', 'add', '--nome', nome, '--conta', conta];
  if (cor) args.push('--cor', cor);
  if (calendarios?.length) args.push('--calendarios', calendarios.join(','));
  return agendaCli(args);
});

// Abre o consentimento no navegador do sistema e espera até 6 min pelo retorno.
ipcMain.handle('hub:google:autorizar', async (_e, nome, escrita = false) => {
  const r = await agendaCli(['google', 'auth', '--nome', nome, ...(escrita ? ['--escrita'] : [])], 6 * 60_000);
  return { ...r, calendarios: parseCalendarios(r.saida) };
});

// Copia as reuniões das contas Microsoft para o calendário "Teams (cópia)" do Google.
ipcMain.handle('hub:espelhar', async () => {
  const r = await agendaCli(['espelhar', '--json'], 5 * 60_000);
  try { return { ok: true, ...JSON.parse(r.saida.slice(r.saida.indexOf('{'))) }; } catch {
    return { ok: false, precisaEscrita: /--escrita/.test(r.saida), saida: r.saida.slice(-400) };
  }
});

ipcMain.handle('hub:google:calendarios', async (_e, nome) => {
  const r = await agendaCli(['google', 'cals', '--nome', nome]);
  return { ...r, calendarios: parseCalendarios(r.saida) };
});

ipcMain.handle('hub:google:remover', (_e, nome) => agendaCli(['google', 'rm', '--nome', nome]));

// O ponto vermelho da bandeja é só do Teams (quem chamou): e-mail não lido
// acumula às centenas e deixaria o ícone sempre aceso. E-mail vai na dica.
ipcMain.on('hub:badges', (_e, porAba) => {
  const total = Object.entries(porAba).filter(([k]) => !k.endsWith(' email')).reduce((a, [, n]) => a + n, 0);
  updateTray(total, porAba);
});

ipcMain.on('hub:notify', (_e, { tab, title, body }) => {
  if (!Notification.isSupported()) return;
  const n = new Notification({ title, body, icon: path.join(__dirname, 'icons', 'icon.png') });
  n.on('click', () => showWindow(tab));
  n.show();
});

ipcMain.on('hub:open-external', (_e, url) => {
  if (/^https:\/\//.test(url)) shell.openExternal(url);
});

app.whenReady().then(() => {
  for (const c of [...config.contas, ...config.extras]) prepareSession(`persist:${c.id}`);
  createWindow();
  buildTray();
  if (process.argv.includes('--hidden')) win.hide();
});

app.on('before-quit', () => { quitting = true; });
