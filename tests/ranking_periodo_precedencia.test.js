const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const source = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'RankingService.gs'),
  'utf8'
);

test('ranking usa texto, catálogo e só então datas individuais', () => {
  const start = source.indexOf('function rankingMG_resolverPeriodoCompetitivo_');
  const end = source.indexOf('\nfunction rankingMG_resolverAtributosCompetitivos_', start);
  assert.ok(start >= 0 && end > start, 'rankingMG_resolverPeriodoCompetitivo_ deve existir');
  const trecho = source.slice(start, end);
  const texto = trecho.indexOf('if (periodoCompletoValido_(periodoHistorico)) return periodoHistorico;');
  const catalogo = trecho.indexOf('if (periodoCompletoValido_(periodoLista))');
  const datas = trecho.indexOf('if (periodoCompletoValido_(periodoDatas)) return periodoDatas;');
  assert.ok(texto >= 0 && catalogo > texto && datas > catalogo);
});
