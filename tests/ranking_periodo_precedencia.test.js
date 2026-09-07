const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const source = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'RankingService.gs'),
  'utf8'
);

test('ranking usa janela individual para PRAZO_DIAS e mantém precedência mensal nos demais', () => {
  const start = source.indexOf('function rankingMG_resolverPeriodoCompetitivo_');
  const end = source.indexOf('\nfunction rankingMG_resolverAtributosCompetitivos_', start);
  assert.ok(start >= 0 && end > start, 'rankingMG_resolverPeriodoCompetitivo_ deve existir');
  const trecho = source.slice(start, end);
  assert.match(trecho, /ehTipoMetaPrazoDias_\(tipoMeta\)/);
  assert.match(trecho, /return periodoCompletoValido_\(periodoDatas\) \? periodoDatas/);
  const texto = trecho.indexOf('if (periodoCompletoValido_(periodoHistorico)) return periodoHistorico;');
  const catalogo = trecho.indexOf('if (periodoCompletoValido_(periodoLista))');
  assert.ok(texto >= 0 && catalogo > texto);
});
