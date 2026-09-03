const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const utils = fs.readFileSync(path.resolve(__dirname, '..', 'Meu Giro', 'Utils.gs'), 'utf8');
const inicio = utils.indexOf('function meuGiroResumoBuildChave_(idDgmb, idDesafio, idItemEstoque, metaKm, idInscricao)');
const fim = utils.indexOf('\nfunction calcularStatusMeuGiroPorPercentual_', inicio);
assert.ok(inicio >= 0 && fim > inicio, 'meuGiroResumoBuildChave_ deve existir');
const buildChave = utils.slice(inicio, fim);

const inicioAtualizar = utils.indexOf('function atualizarMeuGiroResumoComLockAdquirido_(idDgmb, opcoes)');
const fimAtualizar = utils.indexOf('\nfunction atualizarMeuGiroResumoEmLote_', inicioAtualizar);
assert.ok(inicioAtualizar >= 0 && fimAtualizar > inicioAtualizar, 'atualizarMeuGiroResumo_ deve existir');
const atualizarResumo = utils.slice(inicioAtualizar, fimAtualizar);

test('mantém a chave original por ID_INSCRICAO sem incluir desafio/item', () => {
  assert.match(buildChave, /if \(inscricao\) return 'INSCRICAO\|' \+ inscricao;/);
  assert.doesNotMatch(buildChave, /\['INSCRICAO', inscricao, id, desafio/);
});

test('remove log temporário de debug do resumo Meu Giro', () => {
  assert.doesNotMatch(atualizarResumo, /\[MEU_GIRO_RESUMO\]\[DEBUG\]/);
});
