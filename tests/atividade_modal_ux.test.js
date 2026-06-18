const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const repoRoot = path.resolve(__dirname, '..');
const script = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'Script.html'), 'utf8');
const index = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'Index.html'), 'utf8');
const styles = fs.readFileSync(path.join(repoRoot, 'Meu Giro', 'Styles.html'), 'utf8');

function getFunctionSlice(name, nextName) {
  const start = script.indexOf(`function ${name}`);
  const end = script.indexOf(`function ${nextName}`, start + 1);
  assert.notEqual(start, -1, `Função ${name} não encontrada`);
  assert.notEqual(end, -1, `Limite ${nextName} não encontrado`);
  return script.slice(start, end);
}

test('fluxos de atividade não usam confirmações ou alertas nativos', () => {
  const activityFlows = script.slice(
    script.indexOf('async function salvarAtividade'),
    script.indexOf('function novaFrase')
  );

  assert.doesNotMatch(activityFlows, /\b(?:window\.)?confirm\s*\(/);
  assert.doesNotMatch(activityFlows, /\b(?:window\.)?alert\s*\(/);
});

test('modal Meu Giro oferece helpers reutilizáveis e marcação acessível', () => {
  assert.match(script, /function mostrarModalMeuGiro\(opcoes\)/);
  assert.match(script, /function atualizarModalMeuGiro\(opcoes\)/);
  assert.match(script, /function fecharModalMeuGiro\(resolverPendente\)/);
  assert.match(script, /function resolverModalMeuGiro\(valor\)/);
  assert.match(index, /id="global-loading-actions"/);
  assert.match(index, /aria-modal="true"/);
});

test('modal sem botões não cria Promise pendente', () => {
  const helper = getFunctionSlice('mostrarModalMeuGiro', 'atualizarModalMeuGiro');
  assert.match(helper, /if \(!botoes\.length\) \{\s*return Promise\.resolve\(null\);\s*\}/);
  assert.ok(
    helper.indexOf('if (!botoes.length)') < helper.indexOf('return new Promise'),
    'Modal informativo deve retornar antes de registrar resolver'
  );
});

test('resolver centralizado garante que cada decisão seja concluída uma vez', () => {
  const resolver = getFunctionSlice('resolverModalMeuGiro', 'fecharModalMeuGiro');
  assert.match(resolver, /if \(!modalMeuGiroResolver\) return;/);
  assert.match(resolver, /modalMeuGiroResolver = null;\s*resolver\(valor\);/);
  assert.doesNotMatch(script, /modalMeuGiroResolver\([^)]*\)/);
});

test('modal bloqueia clique fora e controla ESC por tipo', () => {
  assert.match(styles, /body\.modal-meu-giro-aberto\s*\{[^}]*overflow:\s*hidden/);
  assert.match(script, /document\.body\.classList\.add\('modal-meu-giro-aberto'\)/);
  assert.match(script, /if \(modalMeuGiroBloqueante \|\| !modalMeuGiroResolver\) return;/);
  assert.match(script, /resolverModalMeuGiro\(escapeValue\)/);
  assert.match(script, /botoes\.length === 1[\s\S]*botoes\[0\]\.valor/);
  assert.match(script, /botoes\.length > 1 \? botoes\[0\]\.valor/);
  assert.match(script, /event\.target === overlay[\s\S]*event\.preventDefault\(\)[\s\S]*event\.stopPropagation\(\)/);
});

test('salvar, editar e excluir mantêm processamento até a atualização visual', () => {
  assert.match(script, /titulo: 'Salvando atividade\.\.\.'/);
  assert.match(script, /titulo: 'Atualizando atividade\.\.\.'/);
  assert.match(script, /titulo: 'Excluindo atividade\.\.\.'/);
  assert.match(script, /function atualizarPainelAposAtividade\(\)/);
  assert.equal((script.match(/await atualizarPainelAposAtividade\(\)/g) || []).length, 3);
});

test('trava de mutação não é liberada antes da atualização do painel', () => {
  [
    ['salvarAtividade', 'salvarEdicaoAtividade'],
    ['salvarEdicaoAtividade', 'iniciarEdicaoAtividadeUI'],
    ['excluirAtividadeUI', 'novaFrase']
  ].forEach(([name, nextName]) => {
    const flow = getFunctionSlice(name, nextName);
    const beforeRefresh = flow.slice(0, flow.indexOf('await atualizarPainelAposAtividade()'));
    assert.doesNotMatch(
      beforeRefresh,
      /activityMutationInProgress = false/,
      `${name} libera activityMutationInProgress antes da atualização`
    );
  });
  assert.match(script, /carregarPainelPosSalvarLeve\(true\)/);
  assert.match(script, /carregarPainel\('pos-salvar-fallback', permitirDuranteMutacao\)/);
});

test('duplicidade permite repetir cadastro e edição com force', () => {
  assert.match(script, /res\.code === 'DUPLICIDADE'/);
  assert.match(script, /salvarAtividade\(true\)/);
  assert.match(script, /res\.code === 'DUPLICIDADE_EDICAO'/);
  assert.match(script, /salvarEdicaoAtividade\(true\)/);
  assert.match(script, /force: !!forcarDuplicidade/);
});

test('fluxo normal não solicita segunda recarga após atualização bem-sucedida', () => {
  [
    ['salvarAtividade', 'salvarEdicaoAtividade'],
    ['salvarEdicaoAtividade', 'iniciarEdicaoAtividadeUI'],
    ['excluirAtividadeUI', 'novaFrase']
  ].forEach(([name, nextName]) => {
    const flow = getFunctionSlice(name, nextName);
    const afterRefresh = flow.slice(flow.indexOf('await atualizarPainelAposAtividade()'));
    assert.doesNotMatch(
      afterRefresh,
      /carregarPainelPosSalvarLeve\(/,
      `${name} agenda uma segunda atualização do painel`
    );
  });
});

test('falhas das três mutações convertem o processamento em modal de erro', () => {
  assert.ok(
    (script.match(/await mostrarErroAcaoAtividade\(\)/g) || []).length >= 6,
    'Respostas inválidas e falhas de conexão devem liberar a UI com mensagem clara'
  );
  assert.match(script, /titulo: 'Não foi possível concluir a ação'/);
  assert.match(script, /pendingDeletionKeys\.delete\(deletionKey\)/);
  assert.match(script, /setActivityFormLocked\(false\)/);
});
