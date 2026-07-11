const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const repoRoot = path.resolve(__dirname, '..');
const inscricao = fs.readFileSync(path.join(repoRoot, 'InscriçãoDesafio', 'codeDGMB.gs'), 'utf8');
const adminCert = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'AdminCertificadoService.gs'), 'utf8');

function sliceFunction(source, name, nextName) {
  const start = source.indexOf(`function ${name}`);
  assert.ok(start >= 0, `${name} deve existir`);
  const end = nextName ? source.indexOf(`function ${nextName}`, start) : source.indexOf('\nfunction ', start + 10);
  assert.ok(end > start, `${name} deve ter fim localizável`);
  return source.slice(start, end);
}

const processarUpload = sliceFunction(inscricao, 'processarUpload', 'verificarCPF');
const gravarInscricao = sliceFunction(inscricao, 'gravarInscricaoDesafio', 'getScriptUrl');
const atualizarStatus = sliceFunction(adminCert, 'atualizarStatusValidacaoCertificadoAdmin', 'adminCertificadoBuildMapaNomesPessoas_');

test('nova inscrição sincroniza MEU_GIRO_RESUMO após appendRow bem-sucedido', () => {
  assert.match(gravarInscricao, /s\.appendRow\(novaLinha\);[\s\S]*?try \{[\s\S]*?atualizarMeuGiroResumo_\(d\.id_dgmb\);/);
  assert.match(gravarInscricao, /\[MEU_GIRO_RESUMO\]\[ERRO_SINCRONIZACAO\] origem=gravarInscricaoDesafio id_dgmb=/);
  assert.ok(gravarInscricao.indexOf('s.appendRow(novaLinha);') < gravarInscricao.indexOf('atualizarMeuGiroResumo_(d.id_dgmb);'));
});

test('upload de avatar ou comprovante não sincroniza resumo sem evidência de impacto', () => {
  assert.match(processarUpload, /sheet\.getRange\(i \+ 1, col\)\.setValue\(urlArquivo\);[\s\S]*?sheet\.getRange\(i \+ 1, col \+ 1\)\.setValue\('Enviado'\);/);
  assert.doesNotMatch(processarUpload, /atualizarMeuGiroResumo_\(/);
  assert.doesNotMatch(processarUpload, /\[MEU_GIRO_RESUMO\]\[ERRO_SINCRONIZACAO\]/);
});

test('validação administrativa de certificado não sincroniza resumo sem evidência de impacto', () => {
  assert.match(atualizarStatus, /sh\.getRange\(linhaAtualizacao, idxStatusValidacao \+ 1\)\.setValue\(novoStatus\);/);
  assert.doesNotMatch(atualizarStatus, /atualizarMeuGiroResumo_\(/);
  assert.doesNotMatch(atualizarStatus, /\[MEU_GIRO_RESUMO\]\[ERRO_SINCRONIZACAO\]/);
});
