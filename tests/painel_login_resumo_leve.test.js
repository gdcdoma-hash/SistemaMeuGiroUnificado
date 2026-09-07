const assert = require('assert');
const fs = require('fs');
const path = require('path');

const painel = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'PainelService.gs'),
  'utf8'
);

const inicio = painel.indexOf('function getPainelUsuario(idDgmb)');
const fim = painel.indexOf('\nvar PAINEL_MG_AUDITORIA_CARREGAMENTO_INICIAL_', inicio);
assert.ok(inicio >= 0 && fim > inicio, 'getPainelUsuario deve existir');

const getPainelUsuario = painel.slice(inicio, fim);

assert.match(
  getPainelUsuario,
  /var resumoDesafios = obterMeuGiroResumoAtualizadoLeve_\(id, \{ reconciliar: !somenteLeitura \}\) \|\| \[\];/,
  'login deve ler MEU_GIRO_RESUMO pela função leve e respeitar modo somente-leitura'
);
assert.doesNotMatch(
  getPainelUsuario,
  /var resumoDesafios = atualizarMeuGiroResumo_\(id\) \|\| \[\];/,
  'login não deve recalcular o resumo no fluxo normal'
);
assert.match(
  getPainelUsuario,
  /if \(!somenteLeitura && \(possuiPrazoDias \|\| !resumoDesafios\.length\)\) \{[\s\S]*?resumoDesafios = atualizarMeuGiroResumo_\(id\) \|\| \[\];/,
  'login deve reconciliar PRAZO_DIAS e manter fallback quando o resumo estiver vazio'
);
assert.match(
  getPainelUsuario,
  /var desafio = buscarInscricaoPainelMGLeve_\(id, resumoDesafios\);/,
  'login deve usar a leitura leve da inscrição'
);
assert.match(
  getPainelUsuario,
  /lerMeuGiroResumoAtualizadoLogin_/,
  'login deve medir a leitura leve do resumo'
);
assert.match(
  getPainelUsuario,
  /var rankingInfo = \{ posicao: 0, total: 0 \};[\s\S]*?var rankingPorDesafio = \{\};/,
  'ranking deve permanecer neutro no carregamento inicial'
);

console.log('OK: login usa resumo e inscrição leves, com fallback explícito e ranking neutro.');
