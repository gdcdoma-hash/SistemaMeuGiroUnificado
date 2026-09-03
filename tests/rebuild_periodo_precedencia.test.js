const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const source = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'SimulacaoReconstrucaoResumo.gs'),
  'utf8'
);

test('rebuild usa a mesma precedência de período da operação normal', () => {
  const start = source.indexOf('function simularResumoMontarPeriodo_');
  const end = source.indexOf('\nfunction simularResumoComparar_', start);
  assert.ok(start >= 0 && end > start, 'simularResumoMontarPeriodo_ deve existir');
  const trecho = source.slice(start, end);
  const datas = trecho.indexOf('if (periodoCompletoValido_(periodoDatasEspecificas))');
  const texto = trecho.indexOf('else if (periodoCompletoValido_(periodoTextoEspecifico))');
  const catalogo = trecho.indexOf('else if (periodoCompletoValido_(periodoLista))');
  assert.ok(datas >= 0 && texto > datas && catalogo > texto);
});
