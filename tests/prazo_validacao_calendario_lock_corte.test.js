const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const repoRoot = path.resolve(__dirname, '..');
const utils = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'Utils.gs'), 'utf8');
const corte = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'CorteControladoMeuGiroResumo.gs'), 'utf8');

function sliceFunction(source, name, nextName) {
  const start = source.indexOf(`function ${name}`);
  const end = source.indexOf(`\nfunction ${nextName}`, start);
  assert.ok(start >= 0, `${name} deve existir`);
  assert.ok(end > start, `${name} deve ter fim localizável`);
  return source.slice(start, end);
}

test('validação ISO rejeita datas impossíveis e preserva datas reais', () => {
  const isData = sliceFunction(utils, 'isDataIsoValida_', 'atividadeDentroPeriodoOficial_');
  const ctx = {};
  vm.createContext(ctx);
  vm.runInContext(isData, ctx);

  assert.equal(ctx.isDataIsoValida_('2026-02-31'), false);
  assert.equal(ctx.isDataIsoValida_('2026-04-31'), false);
  assert.equal(ctx.isDataIsoValida_('2025-02-29'), false);
  assert.equal(ctx.isDataIsoValida_('2024-02-29'), true);
  assert.equal(ctx.isDataIsoValida_('2026-09-03'), true);
});

test('período inválido no calendário não pode vencer os fallbacks válidos', () => {
  const isData = sliceFunction(utils, 'isDataIsoValida_', 'atividadeDentroPeriodoOficial_');
  const periodo = sliceFunction(utils, 'periodoCompletoValido_', 'bug03PeriodoDesafioLogBackend_');
  const ctx = {};
  vm.createContext(ctx);
  vm.runInContext(`${isData}\n${periodo}`, ctx);

  assert.equal(ctx.periodoCompletoValido_({ inicio: '2026-02-31', fim: '2026-03-10' }), false);
  assert.equal(ctx.periodoCompletoValido_({ inicio: '2026-03-01', fim: '2026-03-10' }), true);
  assert.equal(ctx.periodoCompletoValido_({ inicio: '2026-03-10', fim: '2026-03-01' }), false);
});

test('corte controlado serializa toda substituição com o mesmo ScriptLock', () => {
  const wrapper = sliceFunction(corte, 'substituirMeuGiroResumoPorRebuildTeste', 'substituirMeuGiroResumoPorRebuildTesteComLockAdquirido_');
  const interna = corte.slice(corte.indexOf('function substituirMeuGiroResumoPorRebuildTesteComLockAdquirido_'));

  assert.match(wrapper, /LockService\.getScriptLock\(\)/);
  assert.match(wrapper, /lock\.waitLock\(30000\)/);
  assert.match(wrapper, /return substituirMeuGiroResumoPorRebuildTesteComLockAdquirido_\(\)/);
  assert.match(wrapper, /finally[\s\S]*lock\.releaseLock\(\)/);
  assert.doesNotMatch(wrapper, /clearContents|setValues|copyTo/);
  assert.match(interna, /destino\.clearContents\(\)/);
  assert.match(interna, /\.setValues\(valoresOrigem\)/);
});
