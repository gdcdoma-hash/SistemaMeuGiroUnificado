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

test('resumo legado sem ID_INSCRICAO força reconciliação quando há inscrição moderna apta', () => {
  const fonte = trecho(utils, 'meuGiroResumoPossuiInscricaoAusente_', 'obterMeuGiroResumoAtualizadoLeve_');
  assert.match(fonte, /var idsAptos = Object\.keys\(inscricoesAptas\)/);
  assert.match(fonte, /if \(idxInscricaoResumo < 0\) return true/);
});

test('ensure do resumo migra schema legado adicionando ID_INSCRICAO', () => {
  const fonte = trecho(utils, 'ensureMeuGiroResumoSheet_', 'meuGiroResumoBuildChave_');
  assert.match(fonte, /if \(!layoutAtual\.possuiIdInscricao\)/);
  assert.match(fonte, /setValue\('ID_INSCRICAO'\)/);
});

test('leitor leve do painel usa resolvedor central e Tipo_Meta do catálogo', () => {
  const fonte = trecho(painel, 'painelMG_obterInscricaoLevePorDesafio_', 'buscarInscricaoPainelMG_');
  assert.match(fonte, /var contextoLista = buildListaDesafiosContexto_\(getSpreadsheet_\(\)\)/);
  assert.match(fonte, /var tipoMeta = resolverTipoMetaListaDesafio_/);
  assert.match(fonte, /periodoLista\.tipo_meta/);
  assert.match(fonte, /montarPeriodoHistoricoVinculo_\(row,/);
  assert.match(fonte, /}, tipoMeta\)/);
});

test('ranking usa datas individuais para PRAZO_DIAS', () => {
  const fonte = trecho(ranking, 'rankingMG_resolverPeriodoCompetitivo_', 'rankingMG_resolverAtributosCompetitivos_');
  assert.match(fonte, /ehTipoMetaPrazoDias_\(tipoMeta\)/);
  assert.match(fonte, /return periodoCompletoValido_\(periodoDatas\) \? periodoDatas/);
});

test('rebuild propaga Tipo_Meta e usa datas individuais para PRAZO_DIAS', () => {
  const build = trecho(rebuild, 'simularResumoBuildPeriodos_', 'simularResumoBuildStatusLista_');
  const periodo = trecho(rebuild, 'simularResumoMontarPeriodo_', 'simularResumoComparar_');
  assert.match(build, /idxTipoMeta/);
  assert.match(build, /tipo_meta:/);
  assert.match(periodo, /ehTipoMetaPrazoDias_\(tipoMeta\)/);
  assert.match(periodo, /periodo = periodoDatasEspecificas/);
});


test('obterVinculosDesafioUsuario não usa periodoDetalhe/tipoMeta antes da declaração', () => {
  const fonte = trecho(utils, 'obterVinculosDesafioUsuario_', 'obterActivityIdRegistroKm_');
  const idxPeriodoLista = fonte.indexOf('var periodoLista = resolverPeriodoListaDesafio_');
  const idxTipoMeta = fonte.indexOf('var tipoMeta = resolverTipoMetaListaDesafio_');
  const idxPeriodoDetalheUso = fonte.indexOf('periodoDetalhe.nome_desafio');
  assert.ok(idxPeriodoLista >= 0 && idxTipoMeta > idxPeriodoLista);
  assert.equal(idxPeriodoDetalheUso, -1, 'obterVinculosDesafioUsuario_ não deve tocar periodoDetalhe inexistente');
});
