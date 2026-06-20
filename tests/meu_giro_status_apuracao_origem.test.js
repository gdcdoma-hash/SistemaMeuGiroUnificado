const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const utils = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'Utils.gs'),
  'utf8'
);
const simulacao = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'SimulacaoReconstrucaoResumo.gs'),
  'utf8'
);

const inicio = utils.indexOf('function atualizarMeuGiroResumo_(idDgmb, opcoes)');
const fim = utils.indexOf('\nfunction atualizarMeuGiroResumoEmLote_', inicio);
assert.ok(inicio >= 0 && fim > inicio, 'atualizarMeuGiroResumo_ deve existir');
const atualizarResumo = utils.slice(inicio, fim);

test('Status_Apuracao nasce do percentual concluído na geração do MEU_GIRO_RESUMO', () => {
  assert.match(utils, /function calcularStatusMeuGiroPorPercentual_\(percentualConcluido\)[\s\S]*parseLocalizedNumber_\(percentualConcluido\) >= 100[\s\S]*'CONCLUIDO'[\s\S]*'STATUS_EM_ANALISE'/);
  assert.match(atualizarResumo, /var percentualArredondado = Math\.round\(\(percentual \+ Number\.EPSILON\) \* 10\) \/ 10;/);
  assert.match(atualizarResumo, /var status = calcularStatusMeuGiroPorPercentual_\(percentualArredondado\);/);
  assert.match(atualizarResumo, /linha\[idxStatusResumo\] = status;/);
});

test('Status_Usuario_Desafio retornado pela geração acompanha o status calculado na origem', () => {
  assert.match(atualizarResumo, /var statusUsuarioDesafioCalculado = status;/);
  assert.match(atualizarResumo, /status_usuario_desafio: statusUsuarioDesafioCalculado/);
  assert.doesNotMatch(atualizarResumo, /status_usuario_desafio: normalizeText_\(vinculo\.status_usuario_desafio\)/);
});

test('simulação de reconstrução usa a mesma origem de cálculo de status', () => {
  assert.match(simulacao, /var status = calcularStatusMeuGiroPorPercentual_\(percentualArredondado\);/);
  assert.doesNotMatch(simulacao, /status = 'ATIVO'|status = 'EXPIRADO'|status = 'INAPTO'/);
});
