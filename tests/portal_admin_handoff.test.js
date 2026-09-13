const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..', 'Meu Giro');
const handoff = fs.readFileSync(path.join(root, 'PortalHandoffService.gs'), 'utf8');
const code = fs.readFileSync(path.join(root, 'Code.gs'), 'utf8');
const html = fs.readFileSync(path.join(root, 'AdminMeuGiro.html'), 'utf8');

test('handoff administrativo exige token ativo, escopo ADMIN e validade', () => {
  assert.match(handoff, /status !== 'ATIVO'/);
  assert.match(handoff, /agora >= expira/);
  assert.match(handoff, /portalHandoffValidar_\(token, 'ADMIN'\)/);
});

test('rota admin nao abre sem handoff valido', () => {
  assert.match(code, /portalHandoffValidarAdmin_\(token\)/);
  assert.match(code, /Acesso administrativo indisponível/);
  assert.match(code, /setXFrameOptionsMode\(HtmlService\.XFrameOptionsMode\.ALLOWALL\)/);
});

test('tela envia o mesmo handoff em todas as consultas administrativas', () => {
  assert.match(html, /HANDOFF_TOKEN/);
  assert.match(html, /portalAdminMeuGiroBuscarAtletas\(HANDOFF_TOKEN,termo\)/);
  assert.match(html, /portalAdminMeuGiroConsultarAtleta\(HANDOFF_TOKEN,id\)/);
});
