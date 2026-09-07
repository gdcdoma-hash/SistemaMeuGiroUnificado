const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const utils = fs.readFileSync(path.join(root, 'Meu Giro', 'Utils.gs'), 'utf8');
const painel = fs.readFileSync(path.join(root, 'Meu Giro', 'PainelService.gs'), 'utf8');
const ranking = fs.readFileSync(path.join(root, 'Meu Giro', 'RankingService.gs'), 'utf8');
const rebuild = fs.readFileSync(path.join(root, 'Meu Giro', 'SimulacaoReconstrucaoResumo.gs'), 'utf8');

function trecho(source, name, nextName) {
  const start = source.indexOf(`function ${name}`);
  const end = nextName ? source.indexOf(`\nfunction ${nextName}`, start) : source.length;
  assert.ok(start >= 0, `${name} deve existir`);
  assert.ok(end > start, `fim de ${name} deve ser localizado`);
  return source.slice(start, end);
}

test('ListaDesafios mantém índice composto por ID e período', () => {
  const fonte = trecho(utils, 'buildListaDesafiosContexto_', 'buildPeriodoOficialPorAbaEId_');
  assert.match(fonte, /byIdPeriodo/);
  assert.match(fonte, /chaveListaDesafioIdPeriodo_/);
  assert.match(fonte, /contexto\.periodos\.byIdPeriodo\[chaveIdPeriodo\] = periodo/);
  assert.match(fonte, /contexto\.status\.byIdPeriodo\[chaveStatusPeriodo\] = statusLinha/);
  assert.match(fonte, /contexto\.tipoMeta\.byIdPeriodo\[chaveIdPeriodo\] = tipoMeta/);
});

test('resolvedor prefere ID + período e mantém fallback por ID para legado', () => {
  const periodo = trecho(utils, 'resolverPeriodoListaDesafio_', 'resolverStatusListaDesafio_');
  assert.match(periodo, /mapa\.byIdPeriodo/);
  assert.match(periodo, /mapa\.byId\[id\]/);
  const status = trecho(utils, 'resolverStatusListaDesafio_', 'resolverTipoMetaListaDesafio_');
  assert.match(status, /mapa\.byIdPeriodo/);
  assert.match(status, /mapa\.byId\[id\]/);
});

test('vínculo principal resolve catálogo com período da inscrição', () => {
  const fonte = trecho(utils, 'obterVinculosDesafioUsuario_', 'obterRegistrosKmUsuario_');
  assert.match(fonte, /periodoTextoLinha/);
  assert.match(fonte, /resolverStatusListaDesafio_/);
  assert.match(fonte, /resolverPeriodoListaDesafio_/);
  assert.match(fonte, /resolverTipoMetaListaDesafio_/);
});

test('painel leve, ranking e rebuild não usam apenas byId para catálogo moderno', () => {
  const p = trecho(painel, 'painelMG_obterInscricaoLevePorDesafio_', 'buscarInscricaoPainelMG_');
  assert.match(p, /resolverPeriodoListaDesafio_/);
  assert.match(p, /resolverTipoMetaListaDesafio_/);

  const r = trecho(ranking, 'rankingMG_criarIndiceCompetitivo_', 'rankingMG_resolverPeriodoCompetitivo_');
  assert.match(r, /resolverPeriodoListaDesafio_/);

  const b = trecho(rebuild, 'simularResumoBuildVinculos_', 'simularResumoMontarPeriodo_');
  assert.match(b, /resolverPeriodoListaDesafio_/);
});
