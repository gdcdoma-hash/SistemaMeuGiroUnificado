const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const script = fs.readFileSync(path.join(root, 'Meu Giro', 'Script.html'), 'utf8');
const registro = fs.readFileSync(path.join(root, 'Meu Giro', 'RegistroService.gs'), 'utf8');

function trecho(source, name, nextName) {
  const start = source.indexOf('function ' + name);
  const end = nextName ? source.indexOf('\nfunction ' + nextName, start) : source.length;
  assert.ok(start >= 0, name + ' deve existir');
  assert.ok(end > start, 'fim de ' + name + ' deve existir');
  return source.slice(start, end);
}

test('backend bloqueia data futura antes de gravar REGISTRO_KM', () => {
  const helper = trecho(registro, 'validarDataAtividadeNaoFutura_', 'registrarAtividade');
  assert.match(helper, /DATA_FUTURA_NAO_PERMITIDA/);
  assert.match(helper, /dataIso > hojeIso/);

  const registrar = trecho(registro, 'registrarAtividade', 'gerarActivityId_');
  const idxValidacao = registrar.indexOf('validarDataAtividadeNaoFutura_');
  const idxSheet = registrar.indexOf('SpreadsheetApp.openById');
  assert.ok(idxValidacao >= 0 && idxSheet > idxValidacao);
});

test('backend também bloqueia edição para data futura', () => {
  const editar = trecho(registro, 'editarAtividade', 'parseKmInputSeguro_');
  assert.match(editar, /validarDataAtividadeNaoFutura_\(payload\.data_atividade\)/);
  assert.match(editar, /DATA_FUTURA_NAO_PERMITIDA|validacaoData\.code/);
});

test('frontend bloqueia data futura em cadastro e edição', () => {
  const helper = trecho(script, 'isFutureActivityDate', 'isValidActivityDate');
  assert.match(helper, /getTodayIsoLocal_/);

  const salvar = trecho(script, 'salvarAtividade', 'salvarEdicaoAtividade');
  const editar = trecho(script, 'salvarEdicaoAtividade', 'iniciarEdicaoAtividadeUI');
  assert.match(salvar, /isFutureActivityDate\(data\)/);
  assert.match(editar, /isFutureActivityDate\(data\)/);
  assert.match(salvar, /Não é possível registrar uma atividade com data futura/);
  assert.match(editar, /Não é possível registrar uma atividade com data futura/);
});
