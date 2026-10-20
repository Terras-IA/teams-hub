'use strict';

// Seletor do compartilhamento de tela: o Electron não tem um próprio no Linux.
const $ = (s) => document.querySelector(s);
let fontes = [];
let tipo = 'screen';
let escolhida = null;

function render() {
  const lista = fontes.filter((f) => f.tipo === tipo);
  $('#fontes').replaceChildren(...lista.map((f) => {
    const card = document.createElement('button');
    card.type = 'button';
    card.className = 'fonte' + (f.id === escolhida ? ' escolhida' : '');
    const img = document.createElement('img');
    img.src = f.miniatura;
    img.alt = '';
    const nome = document.createElement('span');
    nome.textContent = f.nome;
    nome.title = f.nome;
    card.append(img, nome);
    card.addEventListener('click', () => { escolhida = f.id; $('#compartilhar').disabled = false; render(); });
    card.addEventListener('dblclick', () => window.picker.escolher(f.id));
    return card;
  }));
  if (!lista.length) {
    const p = document.createElement('p');
    p.className = 'vazio';
    p.textContent = tipo === 'screen' ? 'Nenhuma tela encontrada.' : 'Nenhuma janela encontrada.';
    $('#fontes').append(p);
  }
}

document.querySelectorAll('nav button').forEach((b) => b.addEventListener('click', () => {
  tipo = b.dataset.tipo;
  document.querySelectorAll('nav button').forEach((x) => x.classList.toggle('ativo', x === b));
  render();
}));
$('#cancelar').addEventListener('click', () => window.picker.escolher(null));
$('#compartilhar').addEventListener('click', () => escolhida && window.picker.escolher(escolhida));
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') window.picker.escolher(null); });

window.picker.onFontes(({ lista, conta }) => {
  fontes = lista;
  $('#quem').textContent = conta ? `Pedido pela conta ${conta}` : '';
  // Com uma tela só, já deixa escolhida: é o caso mais comum.
  const telas = fontes.filter((f) => f.tipo === 'screen');
  if (telas.length === 1) { escolhida = telas[0].id; $('#compartilhar').disabled = false; }
  render();
});
