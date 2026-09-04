const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const source = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'RankingService.gs'),
  'utf8'
);

test('ranking usa período textual mensal antes de datas individuais e catálogo', () => {
  const start = source.indexOf('function rankingMG_resolverPeriodoCompetitivo_');
  const end = source.indexOf('\nfunction rankingMG_resolverAtributosCompetitivos_', start);
  assert.ok(start >= 0 && end > start, 'rankingMG_resolverPeriodoCompetitivo_ deve existir');
  const trecho = source.slice(start, end);
  const texto = trecho.indexOf('if (periodoCompletoValido_(periodoHistorico)) return periodoHistorico;');
  const datas = trecho.indexOf('if (periodoCompletoValido_(periodoDatas)) return periodoDatas;');
  const catalogo = trecho.indexOf('if (periodoCompletoValido_(periodoLista))');
  assert.ok(texto >= 0 && datas > texto && catalogo > datas);
});
