const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const source = fs.readFileSync(path.resolve(__dirname, '..', 'Meu Giro', 'Script.html'), 'utf8');

function trecho(name, nextName) {
  const start = source.indexOf('function ' + name);
  const end = nextName ? source.indexOf('\nfunction ' + nextName, start) : source.length;
  assert.ok(start >= 0, name + ' deve existir');
  assert.ok(end > start, 'fim de ' + name + ' deve existir');
  return source.slice(start, end);
}

test('total mensal usa ID próprio e não pode ser sobrescrito pelo feedback legado do desafio', () => {
  const render = trecho('renderAtividades', 'renderAtividadesState');
  assert.match(render, /id="atividades-total-mes"/);
  assert.doesNotMatch(render, /id="total-pedalado"/);
});

test('feedback legado continua separado do total mensal', () => {
  const feedback = trecho('playTotalPedaladoFeedback', 'playRankingPreviewFeedback');
  assert.match(feedback, /getElementById\('total-pedalado'\)/);
  assert.doesNotMatch(feedback, /atividades-total-mes/);
});
