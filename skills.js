// Central de skills: lista as skills do Claude Code e roda cada uma com
// `claude -p "/skill pedido"`, na conta escolhida (cada conta é um
// CLAUDE_CONFIG_DIR próprio). Sem janela o Claude não consegue pedir
// aprovação, então o que precisaria de aprovação é recusado e a conversa
// pode continuar no terminal com --resume.
const { app, ipcMain, dialog } = require('electron');
const { execFile, spawn } = require('child_process');
const crypto = require('crypto');
const fs = require('fs');
const os = require('os');
const path = require('path');

const SKILLS_DIR = path.join(os.homedir(), '.claude', 'skills');
const CONTAS_DIR = path.join(os.homedir(), '.config', 'teams-hub', 'claude');
const HIST_FILE = path.join(os.homedir(), '.config', 'teams-hub', 'skills-historico.json');
const HIST_MAX = 50;
const EVENTOS_MAX = 400;
// "dontAsk" recusa o que não estiver liberado nas configurações; bypass nunca.
const MODOS = new Set(['plan', 'dontAsk', 'acceptEdits']);

// Só name e description do frontmatter; aceita valor entre aspas e bloco | ou >.
function frontmatter(txt) {
  const m = /^---\r?\n([\s\S]*?)\r?\n---/.exec(txt);
  const out = {};
  if (!m) return out;
  const linhas = m[1].split(/\r?\n/);
  for (let i = 0; i < linhas.length; i++) {
    const mm = /^(name|description):\s*(.*)$/.exec(linhas[i]);
    if (!mm) continue;
    let v = mm[2].trim();
    if (/^[|>][-+]?$/.test(v)) {
      const partes = [];
      while (i + 1 < linhas.length && /^(\s+\S|\s*$)/.test(linhas[i + 1])) partes.push(linhas[++i].trim());
      v = partes.join(' ').trim();
    } else v = v.replace(/^(['"])(.*)\1$/, '$2');
    out[mm[1]] = v;
  }
  return out;
}

function listarSkills() {
  let nomes = [];
  try { nomes = fs.readdirSync(SKILLS_DIR); } catch { return []; }
  const skills = [];
  for (const dir of nomes) {
    try {
      const fm = frontmatter(fs.readFileSync(path.join(SKILLS_DIR, dir, 'SKILL.md'), 'utf8'));
      skills.push({ nome: fm.name || dir, descricao: fm.description || '', origem: fs.realpathSync(path.join(SKILLS_DIR, dir)) });
    } catch { /* link quebrado ou pasta sem SKILL.md */ }
  }
  return skills.sort((a, b) => a.nome.localeCompare(b.nome));
}

// Repositório git de cada skill (a maioria é link para um clone).
function raizGit(dir) {
  for (let d = dir; d !== path.dirname(d); d = path.dirname(d)) if (fs.existsSync(path.join(d, '.git'))) return d;
  return null;
}

function gitPull(repo) {
  return new Promise((resolve) => {
    execFile('git', ['-C', repo, 'pull', '--ff-only', '--quiet'], { timeout: 90_000, env: { ...process.env, GIT_TERMINAL_PROMPT: '0' } },
      (err, stdout, stderr) => resolve({ repo, ok: !err, msg: (stderr || stdout || err?.message || '').trim().slice(0, 300) }));
  });
}

const HORAS = [1, 6, 12, 24];

function lerHistorico() {
  try { return JSON.parse(fs.readFileSync(HIST_FILE, 'utf8')); } catch { return []; }
}

function gravarHistorico(lista) {
  const tmp = HIST_FILE + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(lista.slice(0, HIST_MAX)) + '\n', { mode: 0o600 });
  fs.renameSync(tmp, HIST_FILE);
}

function resumoEntrada(input = {}) {
  const v = input.command || input.file_path || input.path || input.pattern || input.url || input.query
    || input.skill || input.description || JSON.stringify(input);
  return String(v).replace(/\s+/g, ' ').slice(0, 160);
}

// Uma linha do stream-json vira zero ou mais eventos simples para a interface.
function traduzir(d) {
  if (d.type === 'system' && d.subtype === 'init') return [{ tipo: 'inicio', sessao: d.session_id, modelo: d.model }];
  if (d.type === 'assistant') {
    return (d.message?.content || []).flatMap((c) => {
      if (c.type === 'text' && c.text.trim()) return [{ tipo: 'texto', texto: c.text }];
      if (c.type === 'tool_use') return [{ tipo: 'ferramenta', nome: c.name, resumo: resumoEntrada(c.input) }];
      return [];
    });
  }
  if (d.type === 'user') {
    return (d.message?.content || []).filter((c) => c.type === 'tool_result' && c.is_error).map((c) => {
      const t = Array.isArray(c.content) ? c.content.map((x) => x.text || '').join(' ') : String(c.content || '');
      return { tipo: 'aviso', texto: t.replace(/\s+/g, ' ').slice(0, 300) };
    });
  }
  if (d.type === 'result') {
    return [{
      tipo: 'fim', ok: !d.is_error, texto: d.result || '', custo: d.total_cost_usd, duracao: d.duration_ms,
      turnos: d.num_turns, recusas: d.permission_denials?.length || 0,
    }];
  }
  return [];
}

function registrar({ config, saveConfig, logDiag, getWin, notificar }) {
  const cfg = () => config.claude;
  const rodando = new Map(); // id -> { filho, registro }
  let historico = lerHistorico();
  // Execução que estava em andamento quando o app fechou não terminou.
  for (const r of historico) if (r.status === 'rodando') r.status = 'interrompida';

  const enviar = (canal, dados) => {
    const w = getWin();
    if (w && !w.isDestroyed()) w.webContents.send(canal, dados);
  };

  const contaPorId = (id) => cfg().contas.find((c) => c.id === id);

  function ambiente(conta) {
    const env = { ...process.env, PATH: `${path.dirname(cfg().bin)}:${process.env.PATH || ''}` };
    delete env.ELECTRON_RUN_AS_NODE;
    if (conta?.dir) env.CLAUDE_CONFIG_DIR = conta.dir;
    else delete env.CLAUDE_CONFIG_DIR;
    return env;
  }

  function statusConta(conta) {
    return new Promise((resolve) => {
      execFile(cfg().bin, ['auth', 'status', '--json'], { env: ambiente(conta), timeout: 20_000 }, (err, stdout) => {
        try {
          const s = JSON.parse(stdout);
          resolve({ id: conta.id, logado: !!s.loggedIn, email: s.email || '', plano: s.subscriptionType || s.authMethod || '' });
        } catch { resolve({ id: conta.id, logado: false, erro: (err?.message || 'sem resposta').slice(0, 200) }); }
      });
    });
  }

  // Abre um terminal de verdade: login e continuação de conversa precisam dele.
  function abrirTerminal(conta, pasta, args) {
    const env = conta?.dir ? ['env', `CLAUDE_CONFIG_DIR=${conta.dir}`] : [];
    const cmd = [...env, cfg().bin, ...args];
    // Ptyxis (padrão aqui) abre na pasta pedida; outro terminal herda o cwd.
    const [prog, pre] = fs.existsSync('/usr/bin/ptyxis')
      ? ['ptyxis', ['--new-window', '-d', pasta, '--']]
      : ['x-terminal-emulator', ['-e']];
    const filho = spawn(prog, [...pre, ...cmd], {
      cwd: pasta, detached: true, stdio: 'ignore', env: ambiente(null),
    });
    filho.on('error', (err) => logDiag(`[skills] terminal não abriu: ${err.message}`));
    filho.unref();
  }

  // Atualização da lista: relê as pastas e, se ligado, faz git pull (só
  // avanço rápido) nos repositórios de onde as skills vêm. No máximo 24 h.
  let atualizando = null;
  function atualizarLista() {
    if (atualizando) return atualizando;
    atualizando = (async () => {
      const antes = new Set(listarSkills().map((s) => s.nome));
      const repos = [...new Set(listarSkills().map((s) => raizGit(s.origem)).filter(Boolean))];
      const resultados = cfg().atualizar.repos ? await Promise.all(repos.map(gitPull)) : [];
      const depois = listarSkills().map((s) => s.nome);
      const novas = depois.filter((n) => !antes.has(n));
      const r = { em: new Date().toISOString(), total: depois.length, novas, repos: resultados };
      for (const x of resultados) if (!x.ok) logDiag(`[skills] git pull falhou em ${x.repo}: ${x.msg}`);
      cfg().atualizar.ultima = r;
      saveConfig();
      enviar('skills:atualizada', r);
      if (novas.length) notificar({ tab: 'SKILLS', title: `${novas.length} skill(s) nova(s)`, body: novas.slice(0, 5).join(', ') });
      return r;
    })().finally(() => { atualizando = null; });
    return atualizando;
  }
  const venceu = () => {
    const ultima = Date.parse(cfg().atualizar.ultima?.em || 0) || 0;
    return Date.now() - ultima >= cfg().atualizar.horas * 3600_000;
  };
  app.whenReady().then(() => {
    setTimeout(() => venceu() && atualizarLista(), 60_000);
    setInterval(() => venceu() && atualizarLista(), 15 * 60_000);
  });

  ipcMain.handle('skills:atualizar', () => atualizarLista());

  ipcMain.handle('skills:ocultas', (_e, ocultas) => {
    cfg().ocultas = [...new Set(ocultas.map(String))].sort();
    saveConfig();
    return cfg().ocultas;
  });

  ipcMain.handle('skills:listar', () => ({
    skills: listarSkills(),
    ocultas: cfg().ocultas,
    atualizar: cfg().atualizar,
    favoritos: cfg().favoritos,
    sensiveis: cfg().sensiveis,
    contas: cfg().contas,
    pasta: cfg().pasta,
    bin: cfg().bin,
    modo: cfg().modo,
    claudeOk: fs.existsSync(cfg().bin),
  }));

  ipcMain.handle('skills:config:salvar', (_e, { bin, modo, sensiveis, horas, repos }) => {
    bin = String(bin || '').trim().replace(/^~(?=\/)/, os.homedir());
    if (!path.isAbsolute(bin) || !fs.existsSync(bin)) return { erro: `Claude Code não encontrado em ${bin || '(vazio)'}.` };
    if (!MODOS.has(modo)) return { erro: 'Modo de permissão inválido.' };
    if (!HORAS.includes(Number(horas))) return { erro: 'Intervalo de atualização inválido.' };
    cfg().atualizar = { ...cfg().atualizar, horas: Number(horas), repos: !!repos };
    const nomes = [...new Set((sensiveis || []).map((n) => String(n).trim()).filter(Boolean))].sort();
    Object.assign(cfg(), { bin, modo, sensiveis: nomes });
    saveConfig();
    return { ok: true, sensiveis: nomes };
  });

  ipcMain.handle('skills:favorito', (_e, nome, ligado) => {
    const fav = new Set(cfg().favoritos);
    if (ligado) fav.add(nome); else fav.delete(nome);
    cfg().favoritos = [...fav].sort();
    saveConfig();
    return cfg().favoritos;
  });

  ipcMain.handle('skills:escolher-pasta', async (_e, atual) => {
    const r = await dialog.showOpenDialog(getWin(), { defaultPath: atual || cfg().pasta, properties: ['openDirectory'] });
    if (r.canceled || !r.filePaths[0]) return null;
    cfg().pasta = r.filePaths[0];
    saveConfig();
    return cfg().pasta;
  });

  ipcMain.handle('skills:historico', () => historico.map(({ eventos, ...resto }) => resto));
  ipcMain.handle('skills:execucao', (_e, id) => historico.find((r) => r.id === id) || null);

  ipcMain.handle('skills:rodar', (_e, { skill, pedido, contaId, modo, pasta }) => {
    const conta = contaPorId(contaId);
    if (!conta) return { erro: 'Conta não encontrada.' };
    if (!listarSkills().some((s) => s.nome === skill)) return { erro: `Skill ${skill} não encontrada.` };
    if (!fs.existsSync(pasta || '') || !fs.statSync(pasta).isDirectory()) return { erro: 'Pasta de trabalho não existe.' };
    if (!MODOS.has(modo)) modo = 'plan';
    const forcado = cfg().sensiveis.includes(skill) && modo !== 'plan';
    if (forcado) modo = 'plan';

    const id = crypto.randomUUID();
    const prompt = `/${skill}${pedido?.trim() ? ' ' + pedido.trim() : ''}`;
    const registro = {
      id, skill, pedido: pedido || '', conta: conta.id, modo, pasta, inicio: new Date().toISOString(),
      status: 'rodando', eventos: [],
    };
    const evento = (ev) => {
      if (registro.eventos.length < EVENTOS_MAX) registro.eventos.push(ev);
      else if (registro.eventos.length === EVENTOS_MAX) registro.eventos.push({ tipo: 'aviso', texto: 'Saída longa: o resto não foi guardado no histórico.' });
      enviar('skills:evento', { id, ev });
    };
    if (forcado) evento({ tipo: 'aviso', texto: 'Esta skill publica ou envia algo: rodou em modo plano. Para executar, continue no terminal.' });

    const filho = spawn(cfg().bin, ['-p', prompt, '--output-format', 'stream-json', '--verbose', '--permission-mode', modo], {
      cwd: pasta, env: ambiente(conta), stdio: ['ignore', 'pipe', 'pipe'],
    });
    rodando.set(id, { filho, registro });
    historico.unshift(registro);
    gravarHistorico(historico);
    logDiag(`[skills] rodar ${prompt.slice(0, 120)} conta=${conta.id} modo=${modo} pasta=${pasta}`);

    let resto = '';
    let stderr = '';
    filho.stdout.on('data', (buf) => {
      resto += buf.toString();
      const linhas = resto.split('\n');
      resto = linhas.pop();
      for (const l of linhas) {
        if (!l.trim()) continue;
        let d;
        try { d = JSON.parse(l); } catch { continue; }
        for (const ev of traduzir(d)) {
          if (ev.tipo === 'inicio') registro.sessao = ev.sessao;
          if (ev.tipo === 'fim') Object.assign(registro, { resultado: ev.texto, custo: ev.custo, duracao: ev.duracao, ok: ev.ok });
          evento(ev);
        }
      }
    });
    filho.stderr.on('data', (buf) => { stderr = (stderr + buf.toString()).slice(-4000); });
    const terminar = (status, erro) => {
      if (!rodando.has(id)) return;
      rodando.delete(id);
      registro.status = status;
      registro.fim = new Date().toISOString();
      if (erro) evento({ tipo: 'erro', texto: erro });
      gravarHistorico(historico);
      enviar('skills:fim', { id, status });
      const titulo = status === 'ok' ? `Skill concluída · ${skill}` : `Skill ${status === 'cancelada' ? 'cancelada' : 'com erro'} · ${skill}`;
      notificar({ tab: 'SKILLS', title: titulo, body: (registro.resultado || erro || '').slice(0, 180) });
    };
    filho.on('error', (err) => terminar('erro', `Não consegui rodar o claude: ${err.message}`));
    filho.on('close', (code, sinal) => {
      if (registro.status === 'interrompida') return;
      if (registro.status === 'cancelando') return terminar('cancelada');
      if (registro.ok === undefined) return terminar('erro', (stderr.trim() || `claude saiu com código ${code ?? sinal}`).slice(-600));
      terminar(registro.ok ? 'ok' : 'erro');
    });
    return { id, modo, forcado };
  });

  ipcMain.handle('skills:cancelar', (_e, id) => {
    const r = rodando.get(id);
    if (!r) return false;
    r.registro.status = 'cancelando';
    r.filho.kill('SIGTERM');
    setTimeout(() => { if (rodando.has(id)) r.filho.kill('SIGKILL'); }, 5000);
    return true;
  });

  ipcMain.handle('skills:continuar', (_e, id) => {
    const r = historico.find((x) => x.id === id);
    if (!r?.sessao) return { erro: 'Esta execução não tem conversa para continuar.' };
    const conta = contaPorId(r.conta);
    if (!conta) return { erro: `A conta ${r.conta} não existe mais.` };
    abrirTerminal(conta, fs.existsSync(r.pasta) ? r.pasta : os.homedir(), ['--resume', r.sessao]);
    return { ok: true };
  });

  // ---------- contas do Claude Code ----------

  ipcMain.handle('claude:contas', () => Promise.all(cfg().contas.map((c) => statusConta(c).then((s) => ({ ...c, ...s })))));

  ipcMain.handle('claude:conta:adicionar', (_e, { id, nome }) => {
    id = String(id || '').trim().toUpperCase();
    nome = String(nome || '').trim();
    if (!/^[A-Z0-9]{2,12}$/.test(id)) return { erro: 'A sigla precisa ter de 2 a 12 letras ou números, sem espaço.' };
    if (contaPorId(id)) return { erro: `Já existe uma conta ${id}.` };
    if (!nome) return { erro: 'Informe um nome.' };
    const dir = path.join(CONTAS_DIR, id.toLowerCase());
    fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
    // A conta nova enxerga as mesmas skills e configurações da principal.
    for (const item of ['skills', 'settings.json']) {
      const alvo = path.join(os.homedir(), '.claude', item);
      const link = path.join(dir, item);
      if (fs.existsSync(alvo) && !fs.existsSync(link)) fs.symlinkSync(alvo, link);
    }
    cfg().contas.push({ id, nome, dir });
    saveConfig();
    abrirTerminal({ dir }, os.homedir(), ['auth', 'login']);
    return { ok: true };
  });

  ipcMain.handle('claude:conta:login', (_e, id) => {
    const conta = contaPorId(id);
    if (!conta) return { erro: 'Conta não encontrada.' };
    abrirTerminal(conta, os.homedir(), ['auth', 'login']);
    return { ok: true };
  });

  // Tira a conta da lista; a pasta com o login fica, para não apagar nada sem querer.
  ipcMain.handle('claude:conta:remover', (_e, id) => {
    const conta = contaPorId(id);
    if (!conta || !conta.dir) return { erro: 'A conta principal não sai da lista.' };
    cfg().contas = cfg().contas.filter((c) => c.id !== id);
    saveConfig();
    return { ok: true, dir: conta.dir };
  });

  app.on('before-quit', () => {
    for (const { filho, registro } of rodando.values()) {
      registro.status = 'interrompida';
      filho.kill('SIGTERM');
    }
    if (rodando.size) gravarHistorico(historico);
  });
}

module.exports = { registrar, frontmatter, traduzir };
