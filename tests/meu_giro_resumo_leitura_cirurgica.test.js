const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const utils = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'Utils.gs'),
  'utf8'
);

const inicio = utils.indexOf('function atualizarMeuGiroResumoComLockAdquirido_(idDgmb, opcoes)');
const fim = utils.indexOf('\nfunction atualizarMeuGiroResumoEmLote_', inicio);
assert.ok(inicio >= 0 && fim > inicio, 'atualizarMeuGiroResumo_ deve existir');
const atualizarResumo = utils.slice(inicio, fim);

test('remove a leitura integral do fluxo normal e mantém fallback explícito', () => {
  assert.doesNotMatch(atualizarResumo, /shResumo\.getDataRange\(\)\.getValues\(\)/);
  assert.match(utils, /function meuGiroResumoLeituraIntegralFallback_\(shResumo\)[\s\S]*?shResumo\.getDataRange\(\)\.getValues\(\)/);
  assert.match(atualizarResumo, /meuGiroResumoLeituraIntegralFallback_\(shResumo\)/);
  assert.match(atualizarResumo, /leitura_MEU_GIRO_RESUMO_fallback_integral/);
});

test('consulta os índices ID_DGMB e ID_INSCRICAO e seleciona ambos os tipos de candidato', () => {
  assert.match(utils, /getRange\(2, idxId \+ 1, quantidadeLinhasConsultadas, 1\)\.getValues\(\)/);
  assert.match(utils, /usouIdInscricao[\s\S]*?getRange\(2, idxInscricaoResumo \+ 1, quantidadeLinhasConsultadas, 1\)\.getValues\(\)/);
  assert.match(utils, /pertenceAoId \|\| \(idInscricaoLinha && inscricoesDoAtleta\[idInscricaoLinha\]\)/);
});

test('ordena, agrupa e lê somente blocos completos das linhas alvo', () => {
  assert.match(utils, /linhasCandidatas\.sort\(function\(a, b\) \{ return a - b; \}\)/);
  assert.match(utils, /meuGiroResumoAgruparLinhasContiguas_\(linhasCandidatas\)/);
  assert.match(
    utils,
    /getRange\(\s*bloco\.linhaInicial,\s*1,\s*bloco\.quantidadeLinhas,\s*totalColunasResumo\s*\)\.getValues\(\)/
  );
  assert.match(utils, /valoresResumo\[bloco\.linhaInicial \+ linhaBloco - 1\]/);
});

test('preserva chave, atualização física, colunas extras e appendRow', () => {
  assert.match(atualizarResumo, /meuGiroResumoBuildChave_\(/);
  assert.match(atualizarResumo, /var rowAtual = numeroLinha \? \(valoresResumo\[numeroLinha - 1\] \|\| \[\]\) : \[\]/);
  assert.match(atualizarResumo, /linha\[c\] = numeroLinha \? rowAtual\[c\] : ''/);
  assert.match(atualizarResumo, /shResumo\.getRange\(numeroLinha, 1, 1, totalColunasResumo\)\.setValues\(\[linha\]\)/);
  assert.match(atualizarResumo, /shResumo\.appendRow\(linha\)/);
});

test('mantém logs existentes e adiciona métricas da leitura cirúrgica', () => {
  [
    'leitura_MEU_GIRO_RESUMO_indices',
    'quantidade_linhas_consultadas',
    'usou_id_inscricao',
    'leitura_MEU_GIRO_RESUMO_linhas_alvo',
    'quantidade_linhas_completas_lidas',
    'quantidade_blocos_lidos',
    'quantidade_celulas_estimadas',
    'leitura_MEU_GIRO_RESUMO',
    'escrita_MEU_GIRO_RESUMO',
    'atualizarMeuGiroResumo_total'
  ].forEach((trecho) => assert.ok(utils.includes(trecho), `Instrumentação ausente: ${trecho}`));
});
