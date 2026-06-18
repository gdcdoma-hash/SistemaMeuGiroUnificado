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

test('usa e sincroniza o cache de dgmbDesafios ao atualizar a distância realizada', () => {
  const atualizarDistancia = registroService.match(
    /function atualizarDistanciaRealizada_\(idDgmb, opcoes\)\{[\s\S]*?\n\}/
  );

  assert.ok(atualizarDistancia, 'Função atualizarDistanciaRealizada_ não encontrada');
  assert.doesNotMatch(atualizarDistancia[0], /getDataRange\(\)\.getValues\(\)/);
  assert.match(atualizarDistancia[0], /var cacheDesafios = obterDgmbDesafiosCacheExecucao_\(\)/);
  assert.match(atualizarDistancia[0], /var sheet = cacheDesafios\.sheet/);
  assert.match(atualizarDistancia[0], /var dados = cacheDesafios\.values/);
  assert.match(
    atualizarDistancia[0],
    /obterDadosInscricaoUsuario_\(idDgmb, \{[\s\S]*?cache: cacheDesafios/
  );
  assert.match(atualizarDistancia[0], /dados\[i\]\[idxRealizado\] = total/);
  assert.ok(atualizarDistancia[0].includes("'leitura_dgmbDesafios_cache'"));
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

test('filtra somente as linhas do usuário sobre o cache único de dgmbDesafios', () => {
  const helper = extractFunction(utils, 'obterLinhasDgmbDesafiosUsuario_');
  const obterVinculos = extractFunction(utils, 'obterVinculosDesafioUsuario_');

  assert.doesNotMatch(obterVinculos, /getDataRange\(\)\.getValues\(\)/);
  assert.match(obterVinculos, /var cacheDesafios = obterDgmbDesafiosCacheExecucao_\(\)/);
  assert.doesNotMatch(helper, /createTextFinder/);
  assert.doesNotMatch(helper, /\.getRange\(/);
  assert.match(helper, /var values = cacheDesafios\.values \|\| \[\]/);
  assert.match(helper, /normalizeText_\(values\[i\]\[idxId\]\) === id/);
  assert.match(helper, /quantidade_linhas_usuario: linhasUsuario\.length/);
  assert.match(helper, /quantidade_blocos_lidos: 0/);
  assert.ok(utils.includes("'indice_dgmbDesafios_usuario'"));
  assert.ok(utils.includes("'obterVinculosDesafioUsuario_otimizado'"));
  assert.ok(utils.includes('quantidade_vinculos: vinculos.length'));
});

test('cache por execução centraliza a leitura e instrumenta hit e miss', () => {
  assert.match(utils, /var DGMB_DESAFIOS_CACHE_EXECUCAO_ = null;/);
  const cacheHelper = extractFunction(utils, 'obterDgmbDesafiosCacheExecucao_');
  assert.match(cacheHelper, /DGMB_DESAFIOS_CACHE_EXECUCAO_ !== null/);
  assert.match(cacheHelper, /'cache_hit_dgmbDesafios'/);
  assert.match(cacheHelper, /'cache_miss_dgmbDesafios'/);
  assert.match(cacheHelper, /'leitura_dgmbDesafios_cache'/);
  assert.equal((cacheHelper.match(/\.getValues\(\)/g) || []).length, 1);
  assert.match(cacheHelper, /header: header/);
  assert.match(cacheHelper, /map: buildHeaderMap_\(header\)/);
});

test('filtro em memória usa igualdade exata e preserva os números das linhas', () => {
  const helper = extractFunction(utils, 'obterLinhasDgmbDesafiosUsuario_');
  const factory = new Function(`
    function normalizeText_(value) { return value == null ? '' : String(value).trim(); }
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

  const cache = {
    values: rows,
    header: rows[0],
    map: { id_dgmb: 0, id_desafio: 1 },
    usouCache: true
  };

  const primeira = obterLinhas(cache, '1133');
  const segunda = obterLinhas(cache, '1133');

  assert.equal(primeira.quantidadeLinhasTotal, 383);
  assert.equal(primeira.linhas.length, 6);
  assert.deepEqual(primeira.linhas.map(item => item.numeroLinha), [11, 12, 13, 61, 62, 101]);
  assert.ok(primeira.linhas.every(item => item.valores[0] === '1133'));
  assert.ok(!primeira.linhas.some(item => ['21133', '11330'].includes(item.valores[0])));
  assert.deepEqual(segunda, primeira);
});

test('inscrição e painel leve reutilizam o cache de dgmbDesafios', () => {
  const obterDados = extractFunction(utils, 'obterDadosInscricaoUsuario_');
  const painel = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'PainelService.gs'), 'utf8');
  const inscricaoLeve = extractFunction(painel, 'painelMG_obterInscricaoLevePorDesafio_');

  assert.match(obterDados, /contextoDesafios && contextoDesafios\.cache/);
  assert.match(obterDados, /obterDgmbDesafiosCacheExecucao_\(\)/);
  assert.match(inscricaoLeve, /obterDgmbDesafiosCacheExecucao_\(\)/);
  assert.doesNotMatch(inscricaoLeve, /getDataRange\(\)\.getValues\(\)/);
});
