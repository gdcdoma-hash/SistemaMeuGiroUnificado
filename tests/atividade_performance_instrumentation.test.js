const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const repoRoot = path.resolve(__dirname, '..');
const registroService = fs.readFileSync(
  path.join(repoRoot, 'Meu Giro', 'RegistroService.gs'),
  'utf8'
);
const utils = fs.readFileSync(
  path.join(repoRoot, 'Meu Giro', 'Utils.gs'),
  'utf8'
);
const fontes = registroService + '\n' + utils;

test('instrumenta as três operações de atividade e seus tempos totais', () => {
  [
    ['registrarAtividade', 'registrarAtividade_total'],
    ['editarAtividade', 'editarAtividade_total'],
    ['excluirAtividade', 'excluirAtividade_total']
  ].forEach(([operacao, etapaTotal]) => {
    assert.match(fontes, new RegExp(`MEU_GIRO_PERF_OPERACAO_ATUAL_ = '${operacao}'`));
    assert.match(fontes, new RegExp(`'${etapaTotal}'`));
  });
});

test('instrumenta todas as etapas obrigatórias da auditoria P25', () => {
  [
    'LockService',
    'leitura_REGISTRO_KM',
    'busca_duplicidade',
    'escrita_atividade_REGISTRO_KM',
    'edicao_atividade_REGISTRO_KM',
    'exclusao_atividade_REGISTRO_KM',
    'atualizarDistanciaRealizada_',
    'atualizarMeuGiroResumo_',
    'obterVinculosDesafioUsuario_',
    'obterRegistrosKmUsuario_',
    'leitura_MEU_GIRO_RESUMO',
    'escrita_MEU_GIRO_RESUMO',
    'leitura_ListaDesafios_contexto',
    'cache_hit_lista_desafios',
    'cache_miss_lista_desafios',
    'buildListaDesafiosContexto_total',
    'leitura_dgmbDesafios'
  ].forEach((etapa) => {
    assert.ok(fontes.includes(`'${etapa}'`), `Etapa sem instrumentação: ${etapa}`);
  });
});

test('correlaciona logs internos com a operação de atividade atual', () => {
  assert.match(
    utils,
    /payload\.operacao = MEU_GIRO_PERF_OPERACAO_ATUAL_/
  );
});

test('reaproveita uma única leitura de REGISTRO_KM nas sincronizações da P26', () => {
  assert.match(registroService, /atualizarDistanciaRealizada_\(idDgmb, opcoesRegistroKm\)/);
  assert.match(registroService, /atualizarMeuGiroResumo_\(idDgmb, opcoesRegistroKm\)/);
  assert.match(registroService, /dados\.push\(row\.slice\(\)\)/);
  assert.match(registroService, /dados\[linha - 1\]\[cols\.idxKm\] = novoKm/);
  assert.match(registroService, /dados\.splice\(linha - 1, 1\)/);
  assert.match(utils, /function obterRegistrosKmObjetosReaproveitados_\(idDgmb, opcoes\)/);
  assert.match(utils, /contextoId === id/);
  assert.match(utils, /reaproveitados: false/);
  assert.ok(fontes.includes("'leitura_REGISTRO_KM_reaproveitada'"));
});

test('reaproveita uma única leitura de dgmbDesafios ao atualizar a distância realizada', () => {
  const atualizarDistancia = registroService.match(
    /function atualizarDistanciaRealizada_\(idDgmb, opcoes\)\{[\s\S]*?\n\}/
  );

  assert.ok(atualizarDistancia, 'Função atualizarDistanciaRealizada_ não encontrada');
  assert.equal(
    (atualizarDistancia[0].match(/getDataRange\(\)\.getValues\(\)/g) || []).length,
    1
  );
  assert.match(
    atualizarDistancia[0],
    /obterDadosInscricaoUsuario_\(idDgmb, \{\s*abaDesafio: abaDesafio,\s*values: dados\s*\}\)/
  );
  assert.ok(atualizarDistancia[0].includes("'leitura_dgmbDesafios_unica'"));
  assert.ok(
    atualizarDistancia[0].includes('quantidade_linhas_dgmbDesafios: dados && dados.length ? dados.length - 1 : 0')
  );
});

test('reaproveita o contexto completo de ListaDesafios durante a execução', () => {
  assert.match(utils, /var LISTA_DESAFIOS_CACHE_EXECUCAO_ = null;/);
  assert.match(
    utils,
    /if \(LISTA_DESAFIOS_CACHE_EXECUCAO_ !== null\) \{[\s\S]*?return LISTA_DESAFIOS_CACHE_EXECUCAO_;/
  );
  assert.match(utils, /LISTA_DESAFIOS_CACHE_EXECUCAO_ = contexto;/);
  assert.match(
    utils,
    /function buildPeriodoOficialPorAbaEId_\(ss\) \{\s*return buildListaDesafiosContexto_\(ss\)\.periodos;/
  );
  assert.match(
    utils,
    /function buildMapaStatusDesafioListaPorId_\(ss\) \{\s*return buildListaDesafiosContexto_\(ss\)\.status;/
  );
  assert.match(
    utils,
    /function obterVinculosDesafioUsuario_\(idDgmb\)[\s\S]*?var contextoLista = buildListaDesafiosContexto_\(ss\);/
  );
});
