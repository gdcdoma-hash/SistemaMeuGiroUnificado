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

function extractFunction(source, name) {
  const match = source.match(new RegExp(`function ${name}\\([^]*?\\n\\}`));
  assert.ok(match, `Função ${name} não encontrada`);
  return match[0];
}

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

test('busca somente as linhas de dgmbDesafios pertencentes ao usuário', () => {
  const helper = extractFunction(utils, 'obterLinhasDgmbDesafiosUsuario_');
  const obterVinculos = extractFunction(utils, 'obterVinculosDesafioUsuario_');

  assert.doesNotMatch(obterVinculos, /getDataRange\(\)\.getValues\(\)/);
  assert.doesNotMatch(helper, /createTextFinder/);
  assert.match(helper, /var valoresId = sh\.getRange\([\s\S]*?\)\.getValues\(\)/);
  assert.match(helper, /var idLinha = normalizeText_\(\(valoresId\[i\] \|\| \[\]\)\[0\]\)/);
  assert.match(helper, /var numerosLinhas = \(indiceExecucao\.numerosLinhasPorId\[id\] \|\| \[\]\)\.slice\(\)/);
  assert.match(helper, /blocoAtual && numeroLinha === blocoAtual\.linhaFinal \+ 1/);
  assert.match(helper, /bloco\.quantidadeLinhas,\s*ultimaColuna\s*\)\.getValues\(\)/);
  assert.match(helper, /quantidade_linhas_total: indiceExecucao\.quantidadeLinhasTotal/);
  assert.match(helper, /quantidade_linhas_usuario: indiceExecucao\.porId\[id\]\.length/);
  assert.match(helper, /quantidade_blocos_lidos: indiceExecucao\.quantidadeBlocosPorId\[id\] \|\| 0/);
  assert.ok(utils.includes("'indice_dgmbDesafios_usuario'"));
  assert.ok(utils.includes("'obterVinculosDesafioUsuario_otimizado'"));
  assert.ok(utils.includes('quantidade_vinculos: vinculos.length'));
});

test('índice usa igualdade exata, agrupa blocos e reaproveita as linhas do atleta', () => {
  const helper = extractFunction(utils, 'obterLinhasDgmbDesafiosUsuario_');
  const factory = new Function(`
    var DGMB_DESAFIOS_INDICE_USUARIO_EXECUCAO_ = null;
    function normalizeText_(value) { return value == null ? '' : String(value).trim(); }
    function buildHeaderMap_(header) {
      var out = {};
      header.forEach(function(value, index) { out[normalizeText_(value).toLowerCase()] = index; });
      return out;
    }
    function getOptionalColumnIndex_(map, names) {
      for (var i = 0; i < names.length; i++) {
        if (Object.prototype.hasOwnProperty.call(map, names[i])) return map[names[i]];
      }
      return -1;
    }
    function meuGiroPerfNow_() { return 0; }
    function meuGiroPerfLog_() {}
    ${helper}
    return obterLinhasDgmbDesafiosUsuario_;
  `);
  const obterLinhas = factory();
  const rows = [['ID_DGMB', 'ID_DESAFIO']];
  for (let i = 1; i <= 383; i++) {
    rows.push([String(3000 + i), `D${i}`]);
  }
  [10, 11, 12, 60, 61, 100].forEach(numero => { rows[numero][0] = '1133'; });
  rows[150][0] = '21133';
  rows[151][0] = '11330';

  let leiturasCabecalho = 0;
  let leiturasColunaId = 0;
  const blocosLidos = [];
  const sheet = {
    getLastRow: () => rows.length,
    getLastColumn: () => rows[0].length,
    getRange(row, column, numRows, numColumns) {
      if (row === 1) leiturasCabecalho++;
      return {
        getValues() {
          if (row > 1 && numColumns === 1) leiturasColunaId++;
          if (row > 1 && numColumns > 1) {
            blocosLidos.push({ linhaInicial: row, quantidadeLinhas: numRows });
          }
          return rows.slice(row - 1, row - 1 + numRows)
            .map(values => values.slice(column - 1, column - 1 + numColumns));
        }
      };
    }
  };

  const primeira = obterLinhas(sheet, '1133');
  const segunda = obterLinhas(sheet, '1133');

  assert.equal(primeira.quantidadeLinhasTotal, 383);
  assert.equal(primeira.linhas.length, 6);
  assert.deepEqual(primeira.linhas.map(item => item.numeroLinha), [11, 12, 13, 61, 62, 101]);
  assert.ok(primeira.linhas.every(item => item.valores[0] === '1133'));
  assert.ok(!primeira.linhas.some(item => ['21133', '11330'].includes(item.valores[0])));
  assert.deepEqual(segunda, primeira);
  assert.equal(leiturasCabecalho, 1);
  assert.equal(leiturasColunaId, 1);
  assert.deepEqual(blocosLidos, [
    { linhaInicial: 11, quantidadeLinhas: 3 },
    { linhaInicial: 61, quantidadeLinhas: 2 },
    { linhaInicial: 101, quantidadeLinhas: 1 }
  ]);
});
