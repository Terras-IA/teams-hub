'use strict';

const $ = (sel) => document.querySelector(sel);
const state = { config: null, ativo: 'AGENDA', badges: {}, eventos: [], avisados: new Set() };
let cores = {};

function el(tag, attrs = {}, ...kids) {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') n.className = v;
    else if (k === 'style') n.style.cssText = v;
    else if (k.startsWith('on')) n.addEventListener(k.slice(2), v);
    else n.setAttribute(k, v);
  }
  for (const k of kids) if (k != null) n.append(k);
  return n;
}

// ---------- abas ----------

function addTab(id, nome, cor, sigla, onclick = () => activate(id), extraClass = '') {
  const n = state.badges[id] || 0;
  const btn = el('div', { class: `tab ${extraClass}`.trim(), 'data-tab': id, title: nome, style: `background:${cor}`, onclick },
    sigla, el('span', { class: 'badge' }, n > 0 ? (n > 99 ? '99+' : String(n)) : ''));
  if (state.ativo === id) btn.classList.add('active');
  $('#sidebar').append(btn, el('div', { class: 'tab-label' }, nome));
}

function renderSidebar() {
  $('#sidebar').replaceChildren();
  if (state.config.agendaDisponivel) {
    addTab('AGENDA', 'Agenda', '#4b5263', '📅');
  }
  addTab('SKILLS', 'Skills', '#c96442', '✦');
  $('#sidebar').append(el('div', { class: 'sep' }));
  for (const c of state.config.contas) {
    addTab(c.id, c.nome, c.cor, c.id);
    if (c.email !== false) addTab(idEmail(c.id), `${c.nome} · Email`, c.cor, '✉', undefined, 'mail');
  }
  // Botão direito recarrega a aba: destrava o Teams/Outlook quando a sessão da página envelhece.
  for (const pn of paineis()) {
    document.querySelector(`.tab[data-tab="${pn.id}"]`)?.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      if (!confirm(`Recarregar ${pn.rotulo}? O que estiver sendo digitado nessa aba se perde.`)) return;
      document.querySelector(`#content > webview[data-tab="${pn.id}"]`)?.reload();
      activate(pn.id);
    });
  }
  if (state.config.extras.length && $('#sidebar').children.length) $('#sidebar').append(el('div', { class: 'sep' }));
  for (const c of state.config.extras) {
    const sigla = c.sigla || (c.id === 'GCAL' ? 'G' : c.id);
    // Aba externa abre no navegador do sistema (o Google bloqueia login embutido).
    if (c.externo) addTab(c.id, c.nome + ' ↗', c.cor, sigla, () => window.hub.openExternal(c.url));
    else addTab(c.id, c.nome, c.cor, sigla);
  }
  addTab('CONTAS', 'Configurações', '#3a4050', '⚙', undefined, 'gear');
}

const URL_OUTLOOK = 'https://outlook.office.com/mail/';
const idEmail = (id) => `${id}-MAIL`;

// Cada painel é um webview. O Email usa a mesma sessão (partição) do Teams da
// conta, então aproveita o login da Microsoft sem pedir de novo.
function paineis() {
  const lista = [];
  for (const c of state.config.contas) {
    lista.push({ id: c.id, conta: c.id, rotulo: c.nome, url: c.url });
    if (c.email !== false) lista.push({ id: idEmail(c.id), conta: c.id, rotulo: `${c.nome} · Email`, url: c.email_url || URL_OUTLOOK });
  }
  for (const c of state.config.extras.filter((x) => !x.externo)) lista.push({ id: c.id, conta: c.id, rotulo: c.nome, url: c.url });
  return lista;
}

// Cria o webview de quem ainda não tem e tira o de quem saiu da config.
function syncWebviews() {
  const quer = paineis();
  const ids = new Set(quer.map((pn) => pn.id));
  document.querySelectorAll('#content > webview').forEach((wv) => {
    if (!ids.has(wv.dataset.tab)) { wv.remove(); delete state.badges[wv.dataset.tab]; }
  });
  for (const pn of quer) if (!document.querySelector(`#content > webview[data-tab="${pn.id}"]`)) addWebview(pn);
  enviarBadges();
}

function activate(id) {
  if (!document.querySelector(`.tab[data-tab="${id}"]`)) id = abaInicial();
  state.ativo = id;
  document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.dataset.tab === id));
  document.querySelectorAll('#content > .pane, #content > webview').forEach((p) => p.classList.toggle('active', p.dataset.tab === id));
  if (id === 'AGENDA') renderAgenda();
  if (id === 'CONTAS') renderContas();
  if (id === 'SKILLS') carregarSkills();
}

// Sem agenda, abre na primeira conta; sem contas, na aba Contas para cadastrar.
function abaInicial() {
  if (state.config.agendaDisponivel) return 'AGENDA';
  return state.config.contas[0]?.id || 'CONTAS';
}

function setBadge(id, n) {
  state.badges[id] = n;
  const b = document.querySelector(`.tab[data-tab="${id}"] .badge`);
  if (b) b.textContent = n > 0 ? (n > 99 ? '99+' : String(n)) : '';
  enviarBadges();
}

function enviarBadges() {
  const porAba = Object.fromEntries(paineis().map((pn) => [pn.id.replace('-MAIL', ' email'), state.badges[pn.id] || 0]));
  window.hub.badges(porAba);
}

function addWebview(painel) {
  const wv = el('webview', {
    'data-tab': painel.id,
    src: painel.url,
    partition: `persist:${painel.conta}`,
    preload: state.config.webviewPreload,
    allowpopups: '',
  });
  // Teams e Outlook põem as não lidas no título: "(3) Chat | Microsoft Teams".
  wv.addEventListener('page-title-updated', (e) => {
    const m = /^\((\d+)\)/.exec(e.title || '');
    setBadge(painel.id, m ? Number(m[1]) : 0);
  });
  wv.addEventListener('ipc-message', (e) => {
    if (e.channel === 'nao-lidas') { setBadge(painel.id, Number(e.args[0]) || 0); return; }
    if (e.channel !== 'notif') return;
    if (document.hasFocus() && state.ativo === painel.id) return;
    const { title, body } = e.args[0] || {};
    // Nome atual, não o da criação: a conta pode ter sido renomeada.
    const nome = state.config.contas.find((c) => c.id === painel.conta)?.nome || painel.rotulo;
    const quem = painel.id === painel.conta ? nome : `${nome} · Email`;
    window.hub.notify({ tab: painel.id, title: `[${quem}] ${title}`, body });
  });
  if (state.ativo === painel.id) wv.classList.add('active');
  $('#content').append(wv);
}

// ---------- contas ----------

const URL_TEAMS = 'https://teams.microsoft.com/v2/';

function msg(sel, texto, erro = false) {
  const p = $(sel);
  p.textContent = texto;
  p.classList.toggle('erro', erro);
}

async function salvarTeams(contas, aviso) {
  const r = await window.hub.salvarContas(contas);
  state.config.contas = contas;
  cores = await window.hub.cores();
  renderSidebar();
  syncWebviews();
  renderContas();
  if (r.removidos?.length) msg('#teams-msg', `Removida(s): ${r.removidos.join(', ')}. O login salvo foi apagado.`);
  else if (aviso) msg('#teams-msg', aviso);
}

// Liga/desliga o Outlook da conta (✉ na barra lateral, mesma sessão do Teams).
function caixaEmail(c) {
  const cb = el('input', { type: 'checkbox', title: 'Mostrar o Outlook desta conta' });
  cb.checked = c.email !== false;
  cb.addEventListener('change', () => salvarTeams(
    state.config.contas.map((x) => (x.id === c.id ? { ...x, email: cb.checked } : x)),
    `${c.id}: Email ${cb.checked ? 'ligado' : 'desligado'}.`));
  return el('label', { class: 'check' }, cb, 'Email');
}

function linhaTeams(c) {
  const atualizar = (campo) => (e) => {
    const valor = e.target.value.trim();
    if (campo === 'nome' && !valor) { e.target.value = c.nome; return; }
    salvarTeams(state.config.contas.map((x) => (x.id === c.id ? { ...x, [campo]: valor } : x)), `${c.id}: salvo.`);
  };
  return el('div', { class: 'linha' },
    el('input', { type: 'color', value: c.cor, title: 'Cor', onchange: atualizar('cor') }),
    el('span', { class: 'sigla' }, c.id),
    el('input', { value: c.nome, placeholder: 'Nome', onchange: atualizar('nome') }),
    el('input', { type: 'email', value: c.email || '', placeholder: 'E-mail (opcional)', onchange: atualizar('email') }),
    caixaEmail(c),
    el('button', { type: 'button', onclick: () => activate(c.id) }, 'Abrir'),
    el('button', {
      type: 'button',
      class: 'perigo',
      onclick: () => {
        if (!confirm(`Remover a conta ${c.nome} (${c.id})?\n\nA aba sai do Teams Hub e o login salvo dela é apagado. A conta no Teams não é afetada.`)) return;
        salvarTeams(state.config.contas.filter((x) => x.id !== c.id));
      },
    }, 'Remover'));
}

function novoTeams(e) {
  e.preventDefault();
  const f = new FormData(e.target);
  const id = String(f.get('id')).trim().toUpperCase();
  const nome = String(f.get('nome')).trim();
  const ocupados = new Set([...state.config.contas, ...state.config.extras].map((c) => c.id).concat(['AGENDA', 'CONTAS']));
  if (!/^[A-Z0-9]{2,6}$/.test(id)) return msg('#teams-msg', 'A sigla precisa ter de 2 a 6 letras ou números, sem espaço.', true);
  if (ocupados.has(id)) return msg('#teams-msg', `Já existe uma aba com a sigla ${id}.`, true);
  if (!nome) return msg('#teams-msg', 'Informe um nome.', true);
  const conta = { id, nome, email: String(f.get('email')).trim(), cor: String(f.get('cor')), url: URL_TEAMS };
  e.target.reset();
  salvarTeams([...state.config.contas, conta], `${nome} adicionada. Abra a aba ${id} e entre com a conta da Microsoft.`);
}

async function renderGoogle() {
  if (!$('#google-lista')) return; // sem terras-agenda não há seção de agendas
  const lista = await window.hub.google.listar();
  $('#google-lista').replaceChildren(...lista.map(linhaGoogle));
  if (!lista.length) $('#google-lista').append(el('p', { class: 'muted' }, 'Nenhuma agenda Google conectada.'));
}

function linhaGoogle(g) {
  const linha = el('div', { class: 'linha' });
  const estado = el('span', { class: `estado ${g.autorizada ? 'ok' : ''}` },
    g.autorizada ? `autorizada · ${(g.calendarios || ['primary']).length} calendário(s)` : 'sem autorização');
  const salvar = async (mudanca, aviso) => {
    const r = await window.hub.google.salvar({ nome: g.nome, conta: g.conta, cor: g.cor, calendarios: g.calendarios, ...mudanca });
    msg('#google-msg', r.ok ? aviso : r.saida.slice(0, 300), !r.ok);
    cores = await window.hub.cores();
    renderGoogle();
    if (r.ok) carregarAgenda();
  };
  linha.append(
    el('input', { type: 'color', value: g.cor || '#db4437', title: 'Cor', onchange: (e) => salvar({ cor: e.target.value }, `${g.nome}: cor salva.`) }),
    el('span', { class: 'sigla' }, g.nome),
    el('span', {}, g.conta || ''),
    estado,
    el('button', { type: 'button', onclick: () => autorizarGoogle(g.nome) }, g.autorizada ? 'Reautorizar' : 'Autorizar'),
    el('button', { type: 'button', onclick: () => mostrarCalendarios(g, linha, salvar) }, 'Calendários'),
    el('button', {
      type: 'button',
      class: 'perigo',
      onclick: async () => {
        if (!confirm(`Desconectar a agenda ${g.nome} (${g.conta})?\n\nEla sai da aba Agenda e a autorização salva é apagada. Nada muda no Google.`)) return;
        const r = await window.hub.google.remover(g.nome);
        msg('#google-msg', r.ok ? `${g.nome} desconectada.` : r.saida.slice(0, 300), !r.ok);
        renderGoogle();
        carregarAgenda();
      },
    }, 'Remover'));
  return linha;
}

async function mostrarCalendarios(g, linha, salvar) {
  linha.querySelector('.cals')?.remove();
  const caixa = el('div', { class: 'cals' }, el('span', { class: 'muted' }, 'buscando calendários…'));
  linha.append(caixa);
  const r = await window.hub.google.calendarios(g.nome);
  if (!r.ok || !r.calendarios.length) {
    caixa.replaceChildren(el('span', { class: 'muted' }, 'Não consegui listar. Autorize a agenda primeiro.'));
    return;
  }
  const atuais = new Set(g.calendarios || ['primary']);
  const marcado = (c) => atuais.has(c.id) || (c.principal && atuais.has('primary'));
  // O calendário de cópia não entra na agenda: tudo apareceria duas vezes.
  const caixas = r.calendarios.filter((c) => c.id !== g.espelho_id).map((c) => {
    const cb = el('input', { type: 'checkbox', value: c.principal ? 'primary' : c.id });
    cb.checked = marcado(c);
    return el('label', {}, cb, c.nome + (c.principal ? ' (principal)' : ''));
  });
  caixa.replaceChildren(...caixas, el('div', { class: 'acoes' },
    el('button', {
      type: 'button',
      onclick: () => {
        const escolhidos = [...caixa.querySelectorAll('input:checked')].map((i) => i.value);
        if (!escolhidos.length) return msg('#google-msg', 'Escolha pelo menos um calendário.', true);
        salvar({ calendarios: escolhidos }, `${g.nome}: ${escolhidos.length} calendário(s) na agenda.`);
      },
    }, 'Salvar calendários')));
}

async function autorizarGoogle(nome) {
  msg('#google-msg', `${nome}: autorize na aba que abriu no seu navegador (até 6 minutos)…`);
  const r = await window.hub.google.autorizar(nome);
  msg('#google-msg', r.ok ? `${nome}: autorizada.` : `${nome}: não autorizada. ${r.saida.slice(-300)}`, !r.ok);
  renderGoogle();
  if (r.ok) carregarAgenda();
}

async function novoGoogle(e) {
  e.preventDefault();
  const f = new FormData(e.target);
  const nome = String(f.get('nome')).trim();
  const conta = String(f.get('conta')).trim();
  if (!/^[\p{L}\p{N} _-]{2,30}$/u.test(nome)) return msg('#google-msg', 'Nome com 2 a 30 letras, números, espaço, - ou _.', true);
  const existentes = await window.hub.google.listar();
  if (existentes.some((g) => g.nome.toLowerCase() === nome.toLowerCase())) return msg('#google-msg', `Já existe uma agenda chamada ${nome}.`, true);
  const r = await window.hub.google.salvar({ nome, conta, cor: String(f.get('cor')) });
  if (!r.ok) return msg('#google-msg', r.saida.slice(0, 300), true);
  e.target.reset();
  cores = await window.hub.cores();
  await renderGoogle();
  autorizarGoogle(nome);
}

function renderContas() {
  $('#teams-lista').replaceChildren(...state.config.contas.map(linhaTeams));
  renderExtras();
  renderSkillsConfig();
  renderClaude();
  renderAtualizacoes();
  renderGoogle();
}

// ---------- agenda ----------

const fmtHora = (d) => d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
const fmtDia = (d) => d.toLocaleDateString('pt-BR', { weekday: 'long', day: '2-digit', month: 'long' });
const corDe = (origem) => cores[origem] || '#7b8190';

function marcarConflitos(eventos) {
  const comHora = eventos.filter((e) => !e.dia_inteiro);
  for (let i = 0; i < comHora.length; i++) {
    for (let j = i + 1; j < comHora.length; j++) {
      const a = comHora[i]; const b = comHora[j];
      if (b.ini >= a.end) continue;
      if (a.ini < b.end && b.ini < a.end) {
        (a.choques ||= []).push(b);
        (b.choques ||= []).push(a);
      }
    }
  }
}

async function carregarAgenda() {
  const dias = Number($('#agenda-dias').value);
  $('#agenda-status').textContent = 'carregando…';
  const r = await window.hub.agenda(dias);
  if (r.erro) {
    state.eventos = [];
    state.avisos = [r.erro];
  } else {
    state.eventos = (r.eventos || []).map((e) => ({ ...e, ini: new Date(e.inicio), end: new Date(e.fim) }))
      .sort((a, b) => a.ini - b.ini);
    marcarConflitos(state.eventos);
    state.avisos = r.avisos || [];
  }
  state.atualizado = new Date();
  renderAgenda();
}

function avisoLegivel(txt) {
  if (/^Google .*<!doctype html/i.test(txt)) {
    return 'Google: o endereço do feed não é o "endereço secreto no formato iCal" (devolveu a página de login). '
      + 'Troque com: terras_agenda.py feeds add --nome Google --url "<endereço secreto>". Enquanto isso, a aba Google Agenda mostra o calendário direto.';
  }
  return txt.length > 300 ? txt.slice(0, 300) + '…' : txt;
}

function renderAgenda() {
  const agora = new Date();
  const evs = state.eventos;
  const nConf = evs.filter((e) => e.choques).length;
  $('#agenda-status').textContent = state.atualizado
    ? `${evs.length} compromisso(s)` + (nConf ? ` · ${nConf} em choque` : '') + ` · atualizado às ${fmtHora(state.atualizado)}`
    : 'carregando…';

  const avisos = $('#agenda-avisos');
  avisos.replaceChildren(...(state.avisos || []).map((a) => el('div', { class: 'aviso' }, avisoLegivel(a))));

  const prox = evs.find((e) => !e.dia_inteiro && e.end > agora);
  const next = $('#agenda-next');
  if (prox) {
    const min = Math.round((prox.ini - agora) / 60000);
    const quando = min <= 0 ? 'acontecendo agora' : min < 60 ? `em ${min} min` : `${fmtDia(prox.ini)} às ${fmtHora(prox.ini)}`;
    next.replaceChildren(el('div', { class: 'next', style: `border-left-color:${corDe(prox.origem.split('+')[0])}` },
      el('div', { class: 'when' }, `Próximo · ${quando} · ${prox.origem}`),
      el('div', { class: 'what' }, prox.titulo)));
  } else next.replaceChildren();

  const lista = $('#agenda-list');
  if (!evs.length) {
    lista.replaceChildren(el('p', { class: 'vazio' }, state.atualizado ? 'Nenhum compromisso no período.' : ''));
    return;
  }
  const porDia = new Map();
  for (const e of evs) {
    const k = e.ini.toDateString();
    if (!porDia.has(k)) porDia.set(k, []);
    porDia.get(k).push(e);
  }
  lista.replaceChildren(...[...porDia.values()].map((dia) => el('div', { class: 'day' },
    el('h2', {}, fmtDia(dia[0].ini)),
    ...dia.map((e) => {
      const origens = e.origem.split('+');
      const cls = ['ev'];
      if (e.end <= agora) cls.push('past');
      else if (e.ini <= agora) cls.push('now');
      if (e.choques) cls.push('conflict');
      return el('div', { class: cls.join(' '), style: `border-left-color:${corDe(origens[0])}` },
        el('div', { class: 'hora' }, e.dia_inteiro ? 'dia inteiro' : `${fmtHora(e.ini)}–${fmtHora(e.end)}`),
        el('div', {},
          el('div', { class: 'titulo' }, e.titulo || '(sem título)'),
          el('div', { class: 'meta' },
            ...origens.map((o) => el('span', { class: 'chip', style: `background:${corDe(o)}` }, o)),
            e.local || ''),
          e.choques ? el('div', { class: 'choque' }, '⚠ choca com: ' + e.choques.map((c) => `${c.titulo} (${c.origem})`).join(', ')) : null),
        e.url ? el('button', { type: 'button', onclick: () => window.hub.openExternal(e.url) }, e.online ? 'Abrir reunião' : 'Abrir') : el('span'));
    }))));
}

async function espelhar() {
  const botao = $('#agenda-espelhar');
  botao.disabled = true;
  msg('#espelho-msg', 'Copiando as reuniões para o Google…');
  try {
    let r = await window.hub.espelhar();
    if (!r.ok && r.precisaEscrita) {
      const g = (await window.hub.google.listar())[0];
      if (!g) return msg('#espelho-msg', 'Conecte uma agenda Google na aba Contas primeiro.', true);
      const ok = confirm(`Para copiar, o Teams Hub precisa de permissão para criar o calendário "Teams (cópia)" em ${g.conta} e gravar nele.\n\nOs outros calendários continuam só leitura. O consentimento abre no seu navegador. Continuar?`);
      if (!ok) return msg('#espelho-msg', 'Cópia cancelada.');
      msg('#espelho-msg', 'Autorize no navegador que abriu (até 6 minutos)…');
      const a = await window.hub.google.autorizar(g.nome, true);
      if (!a.ok) return msg('#espelho-msg', `Não autorizado. ${a.saida.slice(-300)}`, true);
      msg('#espelho-msg', 'Autorizado. Copiando…');
      r = await window.hub.espelhar();
    }
    if (!r.ok) return msg('#espelho-msg', `Não consegui copiar. ${r.saida}`, true);
    const feito = `${r.criados} nova(s), ${r.atualizados} atualizada(s), ${r.removidos} removida(s), ${r.iguais} sem mudança`;
    const fora = r.avisos?.length ? ` · ${r.avisos.join(' · ')}` : '';
    msg('#espelho-msg', `"${r.calendario}" no Google (${r.contas.join(', ') || 'nenhuma conta'}, ${r.dias} dias): ${feito}. ${fmtHora(new Date())}${fora}`);
  } finally {
    botao.disabled = false;
  }
}

function lembretes() {
  const agora = Date.now();
  const antes = (state.config.agenda.lembrete_min || 5) * 60000;
  for (const e of state.eventos) {
    if (e.dia_inteiro || state.avisados.has(e.inicio + e.titulo)) continue;
    const falta = e.ini - agora;
    if (falta > 0 && falta <= antes) {
      state.avisados.add(e.inicio + e.titulo);
      window.hub.notify({ tab: 'AGENDA', title: `Em ${Math.round(falta / 60000)} min · ${e.origem}`, body: e.titulo });
    }
  }
}

// ---------- skills ----------

const sk = { dados: null, atual: null, exibindo: null, rodando: new Set() };
const fmtCusto = (v) => (v == null ? '' : `US$ ${v.toFixed(2)}`);
const fmtDuracao = (ms) => (ms == null ? '' : ms < 60000 ? `${Math.round(ms / 1000)} s` : `${Math.round(ms / 60000)} min`);

async function carregarSkills() {
  const primeira = !sk.dados;
  sk.dados = await window.hub.skills.listar();
  if (primeira) $('#sk-modo').value = sk.dados.modo;
  const conta = $('#sk-conta').value;
  $('#sk-conta').replaceChildren(...sk.dados.contas.map((c) => el('option', { value: c.id }, c.nome)));
  if (sk.dados.contas.some((c) => c.id === conta)) $('#sk-conta').value = conta;
  $('#sk-pasta').textContent = sk.dados.pasta;
  renderListaSkills();
  renderHistorico();
  if (!sk.dados.claudeOk) msg('#sk-msg', 'Claude Code não encontrado em ~/.local/bin/claude.', true);
}

function renderListaSkills() {
  const termo = $('#sk-busca').value.trim().toLowerCase();
  const fav = new Set(sk.dados.favoritos);
  const ocultas = new Set(sk.dados.ocultas);
  const visiveis = sk.dados.skills.filter((s) => !ocultas.has(s.nome));
  const achadas = visiveis.filter((s) => !termo || `${s.nome} ${s.descricao}`.toLowerCase().includes(termo));
  const item = (s) => el('div', {
    class: `sk-item${s.nome === sk.atual ? ' ativo' : ''}`,
    title: s.descricao,
    onclick: () => selecionarSkill(s.nome),
  }, el('div', { class: 'sk-item-nome' }, (fav.has(s.nome) ? '★ ' : '') + s.nome),
  el('div', { class: 'sk-item-desc' }, s.descricao));
  const favoritas = achadas.filter((s) => fav.has(s.nome));
  const outras = achadas.filter((s) => !fav.has(s.nome));
  $('#sk-lista').replaceChildren(
    ...(favoritas.length ? [el('h3', {}, 'Favoritas'), ...favoritas.map(item), el('h3', {}, 'Todas')] : []),
    ...outras.map(item));
  $('#sk-total').textContent = (termo ? `${achadas.length} de ${visiveis.length}` : `${visiveis.length} skills`)
    + (ocultas.size ? ` · ${sk.dados.skills.length - visiveis.length} oculta(s)` : '');
}

function selecionarSkill(nome) {
  const s = sk.dados.skills.find((x) => x.nome === nome);
  if (!s) return;
  sk.atual = nome;
  $('#sk-vazio').hidden = true;
  $('#sk-form').hidden = false;
  $('#sk-nome').textContent = s.nome;
  $('#sk-desc').textContent = s.descricao;
  $('#sk-fav').textContent = sk.dados.favoritos.includes(nome) ? '★' : '☆';
  const sensivel = sk.dados.sensiveis.includes(nome);
  $('#sk-sensivel').hidden = !sensivel;
  $('#sk-modo').disabled = sensivel;
  if (sensivel) $('#sk-modo').value = 'plan';
  else if (sk.sensivelAntes) $('#sk-modo').value = sk.dados.modo;
  sk.sensivelAntes = sensivel;
  msg('#sk-msg', '');
  renderListaSkills();
  $('#sk-pedido').focus();
}

function botoesExecucao(registro) {
  const ativa = registro && sk.rodando.has(registro.id);
  $('#sk-cancelar').hidden = !ativa;
  $('#sk-continuar').hidden = !registro?.sessao || ativa;
}

function linhaEvento(ev) {
  switch (ev.tipo) {
    case 'inicio': return el('div', { class: 'sk-ev muted' }, `Sessão iniciada · ${ev.modelo || ''}`);
    case 'texto': return el('div', { class: 'sk-ev sk-texto' }, ev.texto);
    case 'ferramenta': return el('div', { class: 'sk-ev sk-ferr' }, el('b', {}, ev.nome), ' ', ev.resumo);
    case 'aviso': return el('div', { class: 'sk-ev aviso' }, ev.texto);
    case 'erro': return el('div', { class: 'sk-ev sk-erro' }, ev.texto);
    case 'fim': return el('div', { class: `sk-ev sk-fim ${ev.ok ? 'ok' : 'falhou'}` },
      [ev.ok ? 'Concluída' : 'Terminou com erro', fmtDuracao(ev.duracao), fmtCusto(ev.custo),
        ev.recusas ? `${ev.recusas} ação(ões) recusada(s) por falta de aprovação` : ''].filter(Boolean).join(' · '));
    default: return null;
  }
}

function mostrarExecucao(registro) {
  sk.exibindo = registro.id;
  const cab = el('div', { class: 'sk-cab' }, el('b', {}, `/${registro.skill}`), ` ${registro.pedido || ''}`,
    el('div', { class: 'muted' }, `${registro.conta} · ${registro.modo} · ${registro.pasta}`));
  $('#sk-saida').replaceChildren(cab, ...(registro.eventos || []).map(linhaEvento).filter(Boolean));
  if (registro.status === 'interrompida') $('#sk-saida').append(el('div', { class: 'sk-ev sk-erro' }, 'Interrompida: o Teams Hub fechou durante a execução.'));
  botoesExecucao(registro);
}

async function rodarSkill(e) {
  e.preventDefault();
  if (!sk.atual) return;
  msg('#sk-msg', '');
  const r = await window.hub.skills.rodar({
    skill: sk.atual, pedido: $('#sk-pedido').value, contaId: $('#sk-conta').value,
    modo: $('#sk-modo').value, pasta: sk.dados.pasta,
  });
  if (r.erro) return msg('#sk-msg', r.erro, true);
  sk.rodando.add(r.id);
  setBadge('SKILLS', sk.rodando.size);
  mostrarExecucao({ id: r.id, skill: sk.atual, pedido: $('#sk-pedido').value, conta: $('#sk-conta').value, modo: r.modo, pasta: sk.dados.pasta, eventos: [] });
  $('#sk-saida').append(el('div', { id: 'sk-rodando', class: 'sk-ev muted' }, 'Rodando… a resposta aparece aqui conforme chega.'));
  renderHistorico();
}

async function renderHistorico() {
  const lista = await window.hub.skills.historico();
  const icone = { ok: '✓', erro: '✕', cancelada: '⊘', interrompida: '⊘', rodando: '…', cancelando: '…' };
  $('#sk-historico').replaceChildren(...lista.map((r) => el('div', {
    class: `sk-hist ${r.status}${r.id === sk.exibindo ? ' ativo' : ''}`,
    onclick: async () => mostrarExecucao(await window.hub.skills.execucao(r.id)),
  },
  el('span', { class: 'sk-hist-st' }, icone[r.status] || '?'),
  el('span', { class: 'sk-hist-nome' }, r.skill),
  el('span', { class: 'sk-hist-pedido muted' }, r.pedido || ''),
  el('span', { class: 'muted' }, [r.conta, fmtCusto(r.custo), new Date(r.inicio).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' })].filter(Boolean).join(' · ')))));
  if (!lista.length) $('#sk-historico').append(el('p', { class: 'muted' }, 'Nenhuma execução ainda.'));
}

function iniciarSkills() {
  $('#sk-busca').addEventListener('input', renderListaSkills);
  $('#sk-form').addEventListener('submit', rodarSkill);
  $('#sk-fav').addEventListener('click', async () => {
    const ligado = !sk.dados.favoritos.includes(sk.atual);
    sk.dados.favoritos = await window.hub.skills.favorito(sk.atual, ligado);
    selecionarSkill(sk.atual);
  });
  $('#sk-pasta-btn').addEventListener('click', async () => {
    const p = await window.hub.skills.escolherPasta(sk.dados.pasta);
    if (p) { sk.dados.pasta = p; $('#sk-pasta').textContent = p; }
  });
  $('#sk-cancelar').addEventListener('click', () => sk.exibindo && window.hub.skills.cancelar(sk.exibindo));
  $('#sk-continuar').addEventListener('click', async () => {
    const r = await window.hub.skills.continuar(sk.exibindo);
    msg('#sk-msg', r.erro || 'Conversa aberta no terminal.', !!r.erro);
  });
  window.hub.skills.onEvento(({ id, ev }) => {
    if (id !== sk.exibindo) return;
    $('#sk-rodando')?.remove();
    const n = linhaEvento(ev);
    if (n) $('#sk-saida').append(n);
  });
  window.hub.skills.onFim(async ({ id }) => {
    sk.rodando.delete(id);
    setBadge('SKILLS', sk.rodando.size);
    if (id === sk.exibindo) botoesExecucao(await window.hub.skills.execucao(id));
    renderHistorico();
  });
  $('#claude-novo').addEventListener('submit', async (e) => {
    e.preventDefault();
    const f = new FormData(e.target);
    const r = await window.hub.claude.adicionar({ id: f.get('id'), nome: f.get('nome') });
    if (r.erro) return msg('#claude-msg', r.erro, true);
    e.target.reset();
    msg('#claude-msg', 'Conta adicionada. Termine o login no terminal que abriu e clique em Atualizar.');
    renderClaude();
  });
}

async function renderClaude() {
  $('#claude-lista').replaceChildren(el('p', { class: 'muted' }, 'verificando contas…'));
  const contas = await window.hub.claude.contas();
  $('#claude-lista').replaceChildren(...contas.map((c) => el('div', { class: 'linha' },
    el('span', { class: 'sigla' }, c.id),
    el('span', {}, c.nome),
    el('span', { class: `estado ${c.logado ? 'ok' : ''}` }, c.logado ? `${c.email} · ${c.plano}` : 'sem login'),
    el('button', { type: 'button', onclick: renderClaude }, 'Atualizar'),
    el('button', {
      type: 'button',
      onclick: async () => {
        const r = await window.hub.claude.login(c.id);
        msg('#claude-msg', r.erro || `${c.nome}: faça o login no terminal que abriu e clique em Atualizar.`, !!r.erro);
      },
    }, c.logado ? 'Trocar login' : 'Fazer login'),
    c.dir ? el('button', {
      type: 'button',
      class: 'perigo',
      onclick: async () => {
        if (!confirm(`Tirar a conta ${c.nome} da aba Skills?\n\nO login dela continua salvo em ${c.dir}.`)) return;
        const r = await window.hub.claude.remover(c.id);
        msg('#claude-msg', r.erro || `${c.nome} removida da lista. Login salvo em ${r.dir}.`, !!r.erro);
        renderClaude();
      },
    }, 'Remover') : null)));
}
// ---------- configurações: abas de site e skills ----------

async function salvarExtras(extras, aviso) {
  const r = await window.hub.salvarExtras(extras);
  if (r.erro) { msg('#extras-msg', r.erro, true); renderExtras(); return false; }
  state.config.extras = extras;
  cores = await window.hub.cores();
  renderSidebar();
  syncWebviews();
  renderExtras();
  msg('#extras-msg', r.removidos?.length ? `Removida(s): ${r.removidos.join(', ')}. O login salvo foi apagado.` : aviso || '');
  return true;
}

function linhaExtra(x) {
  const mudar = (campo, valor) => salvarExtras(state.config.extras.map((y) => (y.id === x.id ? { ...y, [campo]: valor } : y)), `${x.id}: salvo.`);
  const externo = el('input', { type: 'checkbox', title: 'Abrir no navegador do sistema' });
  externo.checked = !!x.externo;
  externo.addEventListener('change', () => mudar('externo', externo.checked));
  return el('div', { class: 'linha' },
    el('input', { type: 'color', value: x.cor, title: 'Cor', onchange: (e) => mudar('cor', e.target.value) }),
    el('span', { class: 'sigla' }, x.id),
    el('input', { value: x.nome, placeholder: 'Nome', onchange: (e) => (e.target.value.trim() ? mudar('nome', e.target.value.trim()) : (e.target.value = x.nome)) }),
    el('input', { type: 'url', value: x.url, placeholder: 'https://…', onchange: (e) => mudar('url', e.target.value.trim()) }),
    el('label', { class: 'check' }, externo, 'navegador'),
    el('button', { type: 'button', onclick: () => (x.externo ? window.hub.openExternal(x.url) : activate(x.id)) }, 'Abrir'),
    el('button', {
      type: 'button',
      class: 'perigo',
      onclick: () => {
        if (!confirm(`Remover a aba ${x.nome} (${x.id})?${x.externo ? '' : '\n\nO login salvo dela é apagado.'}`)) return;
        salvarExtras(state.config.extras.filter((y) => y.id !== x.id));
      },
    }, 'Remover'));
}

function renderExtras() {
  $('#extras-lista').replaceChildren(...state.config.extras.map(linhaExtra));
  if (!state.config.extras.length) $('#extras-lista').append(el('p', { class: 'muted' }, 'Nenhuma aba de site.'));
}

async function novoExtra(e) {
  e.preventDefault();
  const f = new FormData(e.target);
  const id = String(f.get('id')).trim().toUpperCase();
  if (['AGENDA', 'SKILLS', 'CONTAS'].includes(id) || state.config.extras.some((x) => x.id === id)) {
    return msg('#extras-msg', `Já existe uma aba com a sigla ${id}.`, true);
  }
  const nova = { id, nome: String(f.get('nome')).trim(), cor: String(f.get('cor')), url: String(f.get('url')).trim() };
  if (id.length > 3) nova.sigla = id[0];
  if (f.get('externo')) nova.externo = true;
  if (await salvarExtras([...state.config.extras, nova], `${nova.nome} adicionada.`)) e.target.reset();
}

async function renderSkillsConfig() {
  const d = await window.hub.skills.listar();
  const f = $('#skcfg');
  f.elements.bin.value = d.bin;
  f.elements.modo.value = d.modo;
  f.elements.sensiveis.value = d.sensiveis.join('\n');
  f.elements.horas.value = String(d.atualizar.horas);
  f.elements.repos.checked = d.atualizar.repos;
  $('#skcfg-pasta').textContent = d.pasta;
  mostrarUltimaAtualizacao(d.atualizar.ultima);
  skvis.dados = d;
  renderVisibilidade();
}

async function salvarSkillsConfig(e) {
  e.preventDefault();
  const f = e.target.elements;
  const sensiveis = f.sensiveis.value.split('\n').map((x) => x.trim()).filter(Boolean);
  const r = await window.hub.skills.salvarConfig({
    bin: f.bin.value, modo: f.modo.value, sensiveis, horas: f.horas.value, repos: f.repos.checked,
  });
  if (r.erro) return msg('#skcfg-msg', r.erro, true);
  const instaladas = new Set((await window.hub.skills.listar()).skills.map((s) => s.nome));
  const faltam = r.sensiveis.filter((n) => !instaladas.has(n));
  msg('#skcfg-msg', `Salvo.${faltam.length ? ` Não instaladas aqui (ficam na lista mesmo assim): ${faltam.join(', ')}.` : ''}`);
  sk.dados = null; // a aba Skills relê na próxima vez, com o modo padrão novo
  renderSkillsConfig();
}

function mostrarUltimaAtualizacao(u) {
  if (!u) { $('#skcfg-ultima').textContent = 'Lista ainda não atualizada.'; return; }
  const falhas = (u.repos || []).filter((x) => !x.ok);
  const quando = new Date(u.em).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' });
  $('#skcfg-ultima').textContent = `Última: ${quando} · ${u.total} skills`
    + (u.novas?.length ? ` · nova(s): ${u.novas.join(', ')}` : '')
    + (u.repos?.length ? ` · ${u.repos.length - falhas.length}/${u.repos.length} repositório(s) atualizados` : '')
    + (falhas.length ? ` · falhou: ${falhas.map((x) => `${x.repo.split('/').pop()} (${x.msg.split('\n')[0]})`).join('; ')}` : '');
}

// Lista de visibilidade: cada caixa liga/desliga a skill na aba Skills.
const skvis = { dados: null };

function renderVisibilidade() {
  const d = skvis.dados;
  if (!d) return;
  const termo = $('#skvis-busca').value.trim().toLowerCase();
  const ocultas = new Set(d.ocultas);
  const filtradas = d.skills.filter((s) => !termo || `${s.nome} ${s.descricao}`.toLowerCase().includes(termo));
  $('#skvis-total').textContent = `${d.skills.length - ocultas.size} de ${d.skills.length} exibidas`;
  $('#skvis-lista').replaceChildren(...filtradas.map((s) => {
    const cb = el('input', { type: 'checkbox' });
    cb.checked = !ocultas.has(s.nome);
    cb.addEventListener('change', () => salvarOcultas(cb.checked ? d.ocultas.filter((n) => n !== s.nome) : [...d.ocultas, s.nome]));
    return el('label', { class: 'skvis-item', title: s.descricao }, cb, s.nome);
  }));
}

async function salvarOcultas(lista) {
  skvis.dados.ocultas = await window.hub.skills.ocultas(lista);
  if (sk.dados) sk.dados.ocultas = skvis.dados.ocultas;
  renderVisibilidade();
  if (sk.dados) renderListaSkills();
}

function filtradasVis() {
  const termo = $('#skvis-busca').value.trim().toLowerCase();
  return skvis.dados.skills.filter((s) => !termo || `${s.nome} ${s.descricao}`.toLowerCase().includes(termo)).map((s) => s.nome);
}

// ---------- atualizações ----------

const upd = { estado: null };

function textoEstado(s) {
  switch (s.estado) {
    case 'dev': return 'rodando do código-fonte: atualize com git pull';
    case 'verificando': return 'verificando…';
    case 'em-dia': return 'em dia';
    case 'disponivel': return `versão ${s.versao} disponível`;
    case 'baixando': return `baixando ${s.progresso || 0}%`;
    case 'pronto': return `versão ${s.versao} pronta para instalar`;
    case 'erro': return `erro: ${s.erro}`;
    default: return '';
  }
}

function aplicarEstado(s) {
  upd.estado = { ...upd.estado, ...s };
  const e = upd.estado;
  $('#upd-versao').textContent = e.versaoAtual || state.config.versao;
  $('#upd-status').textContent = textoEstado(e);
  $('#upd-status').classList.toggle('ok', e.estado === 'em-dia');
  $('#upd-verificar').disabled = e.estado === 'dev' || e.estado === 'verificando' || e.estado === 'baixando';
  $('#upd-ver').hidden = !['disponivel', 'baixando', 'pronto'].includes(e.estado);
  if (e.perguntar) abrirDialogoUpd();
  if ($('#upd-dialogo').open) preencherDialogoUpd();
}

function preencherDialogoUpd() {
  const e = upd.estado;
  $('#upd-d-titulo').textContent = `Teams Hub ${e.versao}`;
  $('#upd-d-sub').textContent = `Você está na ${e.versaoAtual}.` + (e.data ? ` Publicada em ${new Date(e.data).toLocaleDateString('pt-BR')}.` : '');
  $('#upd-d-notas').textContent = e.notas || 'Sem notas para esta versão.';
  $('#upd-d-barra').hidden = e.estado !== 'baixando';
  $('#upd-d-barra > div').style.width = `${e.progresso || 0}%`;
  const ok = $('#upd-d-ok');
  ok.disabled = e.estado === 'baixando';
  ok.textContent = e.estado === 'pronto' ? 'Instalar e reiniciar' : e.estado === 'baixando' ? `Baixando ${e.progresso || 0}%…` : 'Atualizar agora';
  $('#upd-d-pular').hidden = e.estado !== 'disponivel';
  msg('#upd-d-msg', e.estado === 'pronto'
    ? 'O app vai fechar, o sistema vai pedir sua senha para instalar e o Teams Hub abre de novo. Seus logins continuam.'
    : e.estado === 'erro' ? `Não deu certo: ${e.erro}` : '', e.estado === 'erro');
}

function abrirDialogoUpd() {
  preencherDialogoUpd();
  if (!$('#upd-dialogo').open) $('#upd-dialogo').showModal();
}

async function renderAtualizacoes() {
  const s = await window.hub.upd.estado();
  $('#upd-auto').checked = s.automatico;
  aplicarEstado({ ...s, perguntar: false });
  const log = await window.hub.upd.changelog();
  $('#upd-changelog').textContent = log.replace(/^# .*\n+/, '');
}

function iniciarConfig() {
  $('#extras-novo').addEventListener('submit', novoExtra);
  $('#skcfg').addEventListener('submit', salvarSkillsConfig);
  $('#skvis-busca').addEventListener('input', renderVisibilidade);
  $('#skvis-todas').addEventListener('click', () => { const f = new Set(filtradasVis()); salvarOcultas(skvis.dados.ocultas.filter((n) => !f.has(n))); });
  $('#skvis-nenhuma').addEventListener('click', () => salvarOcultas([...skvis.dados.ocultas, ...filtradasVis()]));
  $('#skcfg-atualizar').addEventListener('click', async () => {
    $('#skcfg-atualizar').disabled = true;
    $('#skcfg-ultima').textContent = 'Atualizando…';
    try { mostrarUltimaAtualizacao(await window.hub.skills.atualizar()); } finally { $('#skcfg-atualizar').disabled = false; }
  });
  // Atualização automática: relê a aba Skills e a configuração com a lista nova.
  window.hub.skills.onAtualizada((u) => {
    mostrarUltimaAtualizacao(u);
    if (sk.dados) carregarSkills();
    if (skvis.dados) renderSkillsConfig();
  });
  $('#skcfg-pasta-btn').addEventListener('click', async () => {
    const p = await window.hub.skills.escolherPasta($('#skcfg-pasta').textContent);
    if (p) { $('#skcfg-pasta').textContent = p; sk.dados = null; }
  });
  $('#upd-auto').addEventListener('change', (e) => window.hub.upd.auto(e.target.checked));
  $('#upd-verificar').addEventListener('click', async () => aplicarEstado(await window.hub.upd.verificar()));
  $('#upd-ver').addEventListener('click', abrirDialogoUpd);
  $('#upd-d-depois').addEventListener('click', () => $('#upd-dialogo').close());
  $('#upd-d-pular').addEventListener('click', async () => { await window.hub.upd.pular(); $('#upd-dialogo').close(); });
  $('#upd-d-ok').addEventListener('click', async () => {
    if (upd.estado.estado === 'pronto') window.hub.upd.instalar();
    else aplicarEstado(await window.hub.upd.baixar());
  });
  window.hub.upd.onEstado(aplicarEstado);
}

// ---------- início ----------

(async () => {
  state.config = await window.hub.config();
  cores = await window.hub.cores();

  $('#agenda').dataset.tab = 'AGENDA';
  $('#skills').dataset.tab = 'SKILLS';
  $('#contas').dataset.tab = 'CONTAS';
  iniciarSkills();
  iniciarConfig();
  renderAtualizacoes();
  renderSidebar();
  syncWebviews();
  $('#teams-novo').addEventListener('submit', novoTeams);
  $('#google-novo').addEventListener('submit', novoGoogle);

  window.hub.onFocusTab(activate);
  if (!state.config.agendaDisponivel) {
    $('#secao-agendas').remove();
    activate(abaInicial());
    return;
  }
  $('#agenda-refresh').addEventListener('click', carregarAgenda);
  $('#agenda-espelhar').addEventListener('click', espelhar);
  $('#agenda-dias').addEventListener('change', carregarAgenda);
  activate('AGENDA');
  await carregarAgenda();
  setInterval(carregarAgenda, 5 * 60000);
  setInterval(() => { lembretes(); if (state.ativo === 'AGENDA') renderAgenda(); }, 30000);
})();
