const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const repoRoot = path.resolve(__dirname, '..');
const service = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'PortalAdminMeuGiroService.gs'), 'utf8');

test('Meu Giro Admin oferece busca por CPF, ID_DGMB ou nome sem expor CPF no retorno', () => {
  assert.match(service, /function portalAdminMeuGiroBuscarAtletas\(termo\)/);
  assert.match(service, /Utilities\.base64Encode\(digitos\)/);
  assert.match(service, /portalAdminMeuGiroNormalizarBusca_\(nome\)\.indexOf\(termoNormalizado\)/);
  assert.match(service, /id_dgmb:\s*id/);
  assert.match(service, /nome:\s*nome/);
  assert.match(service, /cidade_uf:\s*cidade/);
  assert.doesNotMatch(service, /resultados\.push\(\{[\s\S]*?cpf\s*:/);
});

test('consulta administrativa permanece somente leitura e reutiliza o painel oficial', () => {
  assert.match(service, /getPainelUsuario\(id, \{ somenteLeitura: true \}\)/);
  assert.match(service, /somente_leitura:\s*true/);
  assert.match(service, /desafios_ativos:\s*painel\.desafios_ativos \|\| \[\]/);
  assert.match(service, /desafios_historico:\s*painel\.desafios_historico \|\| \[\]/);
});
