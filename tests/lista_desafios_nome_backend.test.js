const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const repoRoot = path.resolve(__dirname, '..');
const utils = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'Utils.gs'), 'utf8');

function getFunctionSlice(name, nextName) {
  const start = utils.indexOf(`function ${name}`);
  const end = utils.indexOf(`function ${nextName}`, start + 1);
  assert.notEqual(start, -1, `Função ${name} não encontrada`);
  assert.notEqual(end, -1, `Limite ${nextName} não encontrado`);
  return utils.slice(start, end);
}

function buildContext(rows) {
  const ctx = {
    SHEETS: { LISTA_DESAFIOS: 'ListaDesafios' },
    Logger: { log() {} },
    painelMG_incrementarAuditoriaCarregamentoInicial_() {},
    meuGiroPerfNow_() { return 0; },
    meuGiroPerfLog_() {},
    normalizeCell_(value) { return String(value == null ? '' : value).trim(); },
    normalizeText_(value) { return String(value == null ? '' : value).trim(); },
    normalizarPeriodoMensal_() { return { inicio: '', fim: '' }; }
  };
  ctx.buildHeaderMap_ = function(headerRow) {
    const map = {};
    headerRow.forEach((value, index) => {
      const key = ctx.normalizeCell_(value).toLowerCase().replace(/\s+/g, ' ').trim();
      if (key) map[key] = index;
    });
    return map;
  };
  ctx.getOptionalColumnIndex_ = function(map, candidates) {
    for (const candidate of candidates) {
      if (Object.prototype.hasOwnProperty.call(map, candidate)) return map[candidate];
    }
    return -1;
  };
  ctx.ss = {
    getSheetByName(name) {
      assert.equal(name, 'ListaDesafios');
      return { getDataRange: () => ({ getValues: () => rows }) };
    }
  };
  vm.createContext(ctx);
  vm.runInContext('var LISTA_DESAFIOS_CACHE_EXECUCAO_ = null;\n' + getFunctionSlice('buildListaDesafiosContexto_', 'buildPeriodoOficialPorAbaEId_'), ctx);
  return ctx;
}

test('ListaDesafios alimenta nome_desafio por id_Desafio_lista mesmo sem aba', () => {
  const rows = [
    ['id_Desafio_lista', 'Nome_Desafio', 'Status'],
    ['127', 'LETRA O', 'ativo'],
    ['128', 'LETRA R', 'ativo'],
    ['129', 'DESAFIO VOCÊ NA META', 'ativo']
  ];
  const ctx = buildContext(rows);

  const contexto = ctx.buildListaDesafiosContexto_(ctx.ss);

  assert.equal(contexto.periodos.byId['127'].nome_desafio, 'LETRA O');
  assert.equal(contexto.periodos.byId['128'].nome_desafio, 'LETRA R');
  assert.equal(contexto.periodos.byId['129'].nome_desafio, 'DESAFIO VOCÊ NA META');
});

test('objeto de desafio usa nome_desafio vindo do cache de ListaDesafios e preserva IDs', () => {
  const rows = [
    ['id_Desafio_lista', 'Nome_Desafio', 'Status'],
    ['128', 'LETRA R', 'ativo']
  ];
  const ctx = buildContext(rows);
  const contexto = ctx.buildListaDesafiosContexto_(ctx.ss);
  const periodoLista = contexto.periodos.byId['128'];
  const desafio = {
    id_desafio: '128',
    id_item_estoque: 'GIRO_R_900',
    nome_desafio: periodoLista.nome_desafio || ''
  };

  assert.deepEqual(desafio, {
    id_desafio: '128',
    id_item_estoque: 'GIRO_R_900',
    nome_desafio: 'LETRA R'
  });
});

test('obterMeuGiroResumoAtualizadoLeve consulta cache de ListaDesafios e injeta nome_desafio', () => {
  const light = getFunctionSlice('obterMeuGiroResumoAtualizadoLeve_', 'meuGiroResumoAgruparLinhasContiguas_');
  assert.match(light, /var periodosListaDesafios = buildListaDesafiosContexto_\(ss\)\.periodos;/);
  assert.match(light, /nome_desafio: obterNomeDesafioListaPorId_\(periodosListaDesafios, row\[idxDesafio\], ''\)/);
});

test('getPainelUsuario preserva nome_desafio em desafios_ativos e desafios_historico', () => {
  const painel = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'PainelService.gs'), 'utf8');
  const inicio = painel.indexOf('function getPainelUsuario(idDgmb)');
  const fim = painel.indexOf('function painelMG_criarAuditoriaCarregamentoInicial_', inicio);
  assert.ok(inicio >= 0 && fim > inicio, 'getPainelUsuario deve existir');
  const getPainelUsuario = painel.slice(inicio, fim);

  assert.match(getPainelUsuario, /Object\.keys\(item \|\| \{\}\)\.forEach[\s\S]*desafioPainel\[chave\] = item\[chave\]/);
  assert.match(getPainelUsuario, /desafios_ativos: desafiosAtivosPainel/);
  assert.match(getPainelUsuario, /desafios_historico: desafiosHistoricoPainel/);
  assert.match(getPainelUsuario, /desafio_em_foco: desafioPrincipalPainel/);
});
