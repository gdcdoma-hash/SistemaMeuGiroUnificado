const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const diagnostico = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'DiagnosticoMeuGiroAtleta.gs'),
  'utf8'
);
const utils = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'Utils.gs'),
  'utf8'
);
const painel = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'PainelService.gs'),
  'utf8'
);

test('diagnóstico usa resumo leve sem reconciliação e painel em modo somente-leitura', () => {
  assert.match(diagnostico, /obterMeuGiroResumoAtualizadoLeve_\(id, \{ reconciliar: false \}\)/);
  assert.match(diagnostico, /getPainelUsuario\(id, \{ somenteLeitura: true \}\)/);
  assert.match(utils, /function obterMeuGiroResumoAtualizadoLeve_\(idDgmb, opcoes\)/);
  assert.match(utils, /var reconciliarAusentes = !\(opcoes && opcoes\.reconciliar === false\)/);
  assert.match(utils, /if \(reconciliarAusentes && meuGiroResumoPossuiInscricaoAusente_/);
  assert.match(painel, /function getPainelUsuario\(idDgmb\)/);
  assert.match(painel, /var opcoes = arguments.length > 1 \? arguments\[1\] : null/);
  assert.match(painel, /var somenteLeitura = !!\(opcoes && opcoes\.somenteLeitura\)/);
  assert.match(painel, /obterMeuGiroResumoAtualizadoLeve_\(id, \{ reconciliar: !somenteLeitura \}\)/);
  assert.match(painel, /if \(!somenteLeitura && \(possuiPrazoDias \|\| !resumoDesafios\.length\)\)/);
});

test('diagnóstico usa o mesmo resolvedor central de período e Tipo_Meta', () => {
  const inicio = diagnostico.indexOf('function diagnosticoMeuGiroLerDgmbDesafios_');
  const fim = diagnostico.indexOf('\nfunction diagnosticoMeuGiroLerResumo_', inicio);
  const fonte = diagnostico.slice(inicio, fim);

  assert.match(fonte, /var contextoLista = buildListaDesafiosContexto_\(getSpreadsheet_\(\)\)/);
  assert.match(fonte, /resolverPeriodoListaDesafio_/);
  assert.match(fonte, /resolverTipoMetaListaDesafio_/);
  assert.match(fonte, /montarPeriodoHistoricoVinculo_\(row,/);
  assert.match(fonte, /}, tipoMeta\)/);
});
