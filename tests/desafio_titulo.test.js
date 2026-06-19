const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const repoRoot = path.resolve(__dirname, '..');
const script = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'Script.html'), 'utf8');

function getFunctionSlice(name, nextName) {
  const start = script.indexOf(`function ${name}`);
  const end = script.indexOf(`function ${nextName}`, start + 1);
  assert.notEqual(start, -1, `Função ${name} não encontrada`);
  assert.notEqual(end, -1, `Limite ${nextName} não encontrado`);
  return script.slice(start, end);
}

const helperCode = getFunctionSlice('getDesafioTitulo', 'getActivityDateParts_') + '\n' +
  getFunctionSlice('getDesafioMesAnoPortugues_', 'parseDesafioPeriodoDate_') + '\n' +
  getFunctionSlice('parseDesafioPeriodoDate_', 'obterTodosDesafiosPainelV2_') + '\n' +
  getFunctionSlice('getDesafioMesAnoLabel', 'obterCampoAtividadePainel_') + '\n' +
  getFunctionSlice('getMesNome_', 'getMesLabel_');
const ctx = {};
vm.createContext(ctx);
vm.runInContext(helperCode, ctx);

test('getDesafioTitulo prioriza nomes amigáveis antes dos IDs técnicos', () => {
  assert.equal(ctx.getDesafioTitulo({ nome_exibicao: 'Giro de Verão', nome_desafio: 'Nome secundário', id_desafio: 128, id_item_estoque: 'GIRO_R_900' }), 'Giro de Verão');
  assert.equal(ctx.getDesafioTitulo({ nome_desafio: 'Desafio das Serras', id_desafio: 127, id_item_estoque: 'GIRO_O_200' }), 'Desafio das Serras');
  assert.equal(ctx.getDesafioTitulo({ desafio: 'Pedal Solidário', id_desafio: 111 }), 'Pedal Solidário');
  assert.equal(ctx.getDesafioTitulo({ titulo: 'Tour Mensal', nome: 'Nome genérico' }), 'Tour Mensal');
  assert.equal(ctx.getDesafioTitulo({ item_nome: 'Item amigável', id_item_estoque: 'GIRO_X_100' }), 'Item amigável');
  assert.equal(ctx.getDesafioTitulo({ nome_item_estoque: 'Nome do item', descricao_item: 'Descrição' }), 'Nome do item');
  assert.equal(ctx.getDesafioTitulo({ descricao_item: 'Descrição amigável' }), 'Descrição amigável');
});

test('getDesafioTitulo não mostra fallback técnico quando há nome amigável', () => {
  const titulo = ctx.getDesafioTitulo({ nome_exibicao: 'Rumo aos 900 km', id_desafio: 128, id_item_estoque: 'GIRO_R_900' });
  assert.equal(titulo, 'Rumo aos 900 km');
  assert.notEqual(titulo, 'Desafio 128 · Item GIRO_R_900');
  assert.doesNotMatch(titulo, /Desafio 128\s*[·-]\s*Item GIRO_R_900/);
});

test('getDesafioTitulo mantém fallback técnico sem nome amigável', () => {
  assert.equal(ctx.getDesafioTitulo({ id_desafio: 128, id_item_estoque: 'GIRO_R_900' }), 'Desafio 128 · Item GIRO_R_900');
  assert.equal(ctx.getDesafioTitulo({ id_item_estoque: 'GIRO_R_900' }), 'Item GIRO_R_900');
  assert.equal(ctx.getDesafioTitulo({ id_desafio: 128 }), 'Desafio 128');
});

test('cards em andamento e concluídos usam getDesafioTitulo', () => {
  const cardDesafio = getFunctionSlice('buildDesafioCardV2_', 'abrirDetalheDesafioV2_');
  const cardConquista = getFunctionSlice('buildConquistaCard_', 'getDesafioMesAnoPortugues_');
  assert.match(cardDesafio, /getDesafioTitulo\(desafio\)/);
  assert.match(cardConquista, /getDesafioTitulo\(desafio\)/);
});

test('ranking, desafio em foco e certificados administrativos usam getDesafioTitulo', () => {
  const foco = getFunctionSlice('renderDesafioEmFocoResumo', 'renderCardDesafioSelecionado');
  const ranking = getFunctionSlice('atualizarSeletorDesafioRanking', 'onRankingDesafioSelectorChange');
  const adminCert = getFunctionSlice('renderAdminValidacaoLista', 'formatAdminStatusValidacao_');
  assert.match(foco, /getDesafioTitulo\(desafioAtual\)/);
  assert.match(foco, /getDesafioTitulo\(item\)/);
  assert.match(ranking, /getDesafioTitulo\(item \|\| \{\}\)/);
  assert.match(adminCert, /getDesafioTitulo\(item\)/);
});


test('getDesafioMesAnoLabel prioriza periodo_desafio formatado como mês/ano amigável', () => {
  assert.equal(ctx.getDesafioMesAnoLabel({ periodo_desafio: 'junho/2026', periodo_inicio: '2026-05-01' }), 'Período: Junho/2026');
});

test('getDesafioMesAnoLabel usa periodo_inicio como fallback e periodo_fim como último fallback', () => {
  assert.equal(ctx.getDesafioMesAnoLabel({ periodo_inicio: '2026-06-01' }), 'Período: Junho/2026');
  assert.equal(ctx.getDesafioMesAnoLabel({ periodo_fim: '2026-04-30' }), 'Período: Abril/2026');
});

test('cards em andamento e conquistas não exibem hífen quando periodo_desafio existe', () => {
  assert.equal(ctx.getDesafioMesAnoLabel({ periodo_desafio: 'junho/2026' }), 'Período: Junho/2026');
  assert.equal(ctx.getDesafioMesAnoPortugues_({ periodo_desafio: 'abril/2026' }), 'Abril/2026');
  assert.notEqual(ctx.getDesafioMesAnoLabel({ periodo_desafio: 'junho/2026' }), 'Período: -');
  assert.notEqual(ctx.getDesafioMesAnoPortugues_({ periodo_desafio: 'abril/2026' }), '-');
});
