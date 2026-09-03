const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const repoRoot = path.resolve(__dirname, '..');
const code = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'Code.gs'), 'utf8');
const registro = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'RegistroService.gs'), 'utf8');
const adminCert = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'AdminCertificadoService.gs'), 'utf8');
const utils = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'Utils.gs'), 'utf8');
const readme = fs.readFileSync(path.join(repoRoot, 'README.md'), 'utf8');
const auditoria = fs.readFileSync(path.join(repoRoot, 'docs', 'auditoria-fluxo-inscricao.md'), 'utf8');

function sliceFunction(source, name, nextName) {
  const start = source.indexOf(`function ${name}`);
  assert.ok(start >= 0, `${name} deve existir`);
  const end = nextName ? source.indexOf(`function ${nextName}`, start) : source.indexOf('\nfunction ', start + 10);
  assert.ok(end > start, `${name} deve ter fim localizável`);
  return source.slice(start, end);
}

function listRepoFiles(dir) {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap(entry => {
    const fullPath = path.join(dir, entry.name);
    const rel = path.relative(repoRoot, fullPath);
    if (entry.isDirectory()) {
      if (['.git', 'node_modules'].includes(entry.name)) return [];
      return listRepoFiles(fullPath);
    }
    return [rel];
  });
}

const doGet = sliceFunction(code, 'doGet', 'include');
const registrarAtividade = sliceFunction(registro, 'registrarAtividade', 'gerarActivityId_');
const atualizarDistancia = sliceFunction(registro, 'atualizarDistanciaRealizada_', 'editarAtividade');
const atualizarStatus = sliceFunction(adminCert, 'atualizarStatusValidacaoCertificadoAdmin', 'adminCertificadoBuildMapaNomesPessoas_');
const atualizarResumo = sliceFunction(utils, 'atualizarMeuGiroResumoComLockAdquirido_', 'atualizarMeuGiroResumoEmLote_');

const arquivosOperacionais = listRepoFiles(repoRoot)
  .filter(file => !file.startsWith('docs/') && !file.startsWith('tests/') && file !== 'README.md');
const codigoOperacional = arquivosOperacionais
  .map(file => fs.readFileSync(path.join(repoRoot, file), 'utf8'))
  .join('\n');

test('código ativo não referencia a pasta/funções legadas de InscriçãoDesafio', () => {
  assert.ok(!fs.existsSync(path.join(repoRoot, 'InscriçãoDesafio')), 'pasta legada não deve existir na árvore ativa');
  assert.doesNotMatch(doGet, /renderInscricaoDGMB/);
  assert.doesNotMatch(doGet, /view\s*===\s*['"]inscricao['"]/);
  assert.doesNotMatch(codigoOperacional, /InscriçãoDesafio|renderInscricaoDGMB|gravarInscricaoDesafio|formDGMB|processarUpload/);
});

test('repositório declara ausência do fluxo real de nova inscrição e aponta correção externa', () => {
  assert.match(readme, /não foi localizado o fluxo atual que cria nova inscrição/);
  assert.match(auditoria, /não encontrou o fluxo atual que cria novas inscrições em `dgmbDesafios`/);
  assert.match(auditoria, /correção funcional deve ser aplicada no repositório que contém o fluxo atual de inscrição/);
  assert.doesNotMatch(codigoOperacional, /gravarInscricao|nova inscrição|repescagem|reativar|adminAtualizarStatus/i);
});

test('não há teste falso de inscrição: fluxo de atividade sincroniza apenas após gravação própria', () => {
  assert.match(registrarAtividade, /sheet\.getRange\(linhaInserida, 1, 1, rowLength\)\.setValues\(\[row\]\);[\s\S]*?atualizarDistanciaRealizada_\(idDgmb, opcoesRegistroKm\);[\s\S]*?atualizarMeuGiroResumoComLockAdquirido_\(idDgmb, opcoesRegistroKm\);/);
  assert.match(atualizarDistancia, /sheet\.getRange\(i \+ 1, idxRealizado \+ 1\)\.setValue\(total\);/);
  assert.doesNotMatch(registrarAtividade, /gravarInscricao|ID_INSCRICAO|Status_Usuario_Desafio/);
});

test('validação administrativa de certificado não sincroniza resumo sem evidência de impacto em inscrição', () => {
  assert.match(atualizarStatus, /sh\.getRange\(linhaAtualizacao, idxStatusValidacao \+ 1\)\.setValue\(novoStatus\);/);
  assert.doesNotMatch(atualizarStatus, /atualizarMeuGiroResumo_\(/);
  assert.doesNotMatch(atualizarStatus, /\[MEU_GIRO_RESUMO\]\[ERRO_SINCRONIZACAO\]/);
});

test('implementação interna permanece responsável por gravar MEU_GIRO_RESUMO a partir de vínculos dgmbDesafios', () => {
  assert.match(atualizarResumo, /var vinculos = obterVinculosDesafioUsuario_\(id\);/);
  assert.match(atualizarResumo, /shResumo\.getRange\(numeroLinha, 1, 1, totalColunasResumo\)\.setValues\(\[linha\]\);/);
  assert.match(atualizarResumo, /shResumo\.appendRow\(linha\);/);
});
