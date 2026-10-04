const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const repoRoot = path.resolve(__dirname, '..');
const utils = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'Utils.gs'), 'utf8');

function sliceFunction(source, name, nextName) {
  const start = source.indexOf(`function ${name}`);
  const end = source.indexOf(`\nfunction ${nextName}`, start);
  assert.ok(start >= 0, `${name} deve existir`);
  assert.ok(end > start, `${name} deve ter fim localizável`);
  return source.slice(start, end);
}

test('PRAZO_DIAS é informado ao resolvedor de janela individual no vínculo', () => {
  const fonte = sliceFunction(utils, 'obterVinculosDesafioUsuario_', 'obterActivityIdRegistroKm_');

  assert.match(fonte, /var tipoMetaListaDesafios = contextoLista\.tipoMeta/);
  assert.match(fonte, /var idxPrazoDias = getOptionalColumnIndex_\(map, \['prazo_dias', 'prazo dias'\]\)/);
  assert.match(fonte, /var tipoMeta = resolverTipoMetaListaDesafio_/);
  assert.match(fonte, /if \(prazoDias > 0\) tipoMeta = 'PRAZO_DIAS'/);
  assert.match(fonte, /montarPeriodoHistoricoVinculo_\([\s\S]*?, tipoMeta\);/);
});

test('janela individual atravessa setembro e outubro sem corte mensal', () => {
  const isData = sliceFunction(utils, 'isDataIsoValida_', 'atividadeDentroPeriodoOficial_');
  const dentro = sliceFunction(utils, 'atividadeDentroPeriodoOficial_', 'normalizarPeriodoMensal_');
  const ctx = {};
  vm.createContext(ctx);
  vm.runInContext(`${isData}\n${dentro}`, ctx);

  assert.equal(ctx.atividadeDentroPeriodoOficial_('2026-09-03', '2026-09-03', '2026-12-01'), true);
  assert.equal(ctx.atividadeDentroPeriodoOficial_('2026-10-15', '2026-09-03', '2026-12-01'), true);
  assert.equal(ctx.atividadeDentroPeriodoOficial_('2026-11-30', '2026-09-03', '2026-12-01'), true);
  assert.equal(ctx.atividadeDentroPeriodoOficial_('2026-09-02', '2026-09-03', '2026-12-01'), false);
  assert.equal(ctx.atividadeDentroPeriodoOficial_('2026-12-02', '2026-09-03', '2026-12-01'), false);
});

test('resumo antigo de PRAZO_DIAS é reconciliado uma vez após a correção', () => {
  const inicio = utils.indexOf('function meuGiroResumoReconciliarPrazoDiasUmaVez_');
  const fim = utils.indexOf('\nfunction obterMeuGiroResumoAtualizadoLeve_', inicio);
  assert.ok(inicio >= 0 && fim > inicio, 'helper de reconciliação deve existir');
  const fonte = utils.slice(inicio, fim);

  assert.match(fonte, /prazoIndividualPorResumoKey/);
  assert.match(fonte, /prazoIndividualPorDesafio/);
  assert.match(fonte, /PropertiesService\.getScriptProperties\(\)/);
  assert.match(fonte, /MEU_GIRO_FIX_PRAZO_ATRAVESSA_MESES_20261003_/);
  assert.match(fonte, /atualizarMeuGiroResumoComLockAdquirido_\(id\)/);
  assert.match(fonte, /propriedades\.setProperty\(chave, '1'\)/);
});
