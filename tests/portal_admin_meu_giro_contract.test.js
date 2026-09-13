const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const code = fs.readFileSync(path.join(__dirname, '..', 'Meu Giro', 'PortalAdminMeuGiroService.gs'), 'utf8');

test('contrato admin usa painel oficial em somente leitura', () => {
  assert.match(code, /getPainelUsuario\(id, \{ somenteLeitura: true \}\)/);
  assert.match(code, /desafios_ativos/);
  assert.match(code, /desafios_historico/);
  assert.match(code, /somente_leitura: true/);
});

test('contrato admin exige ID_DGMB e nao implementa calculo paralelo', () => {
  assert.match(code, /ID_DGMB_OBRIGATORIO/);
  assert.doesNotMatch(code, /getSheetByName_|REGISTRO_KM|MEU_GIRO_RESUMO|dgmbDesafios/);
});
