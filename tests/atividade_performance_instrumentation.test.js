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
