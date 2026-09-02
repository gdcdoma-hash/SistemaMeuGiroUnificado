const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const utils = fs.readFileSync(
  path.resolve(__dirname, '..', 'Meu Giro', 'Utils.gs'),
  'utf8'
);

function trecho(nome, proximo) {
  const inicio = utils.indexOf(`function ${nome}`);
  const fim = proximo ? utils.indexOf(`\nfunction ${proximo}`, inicio) : utils.length;
  assert.ok(inicio >= 0, `${nome} deve existir`);
  assert.ok(fim > inicio, `fim de ${nome} deve ser localizado`);
  return utils.slice(inicio, fim);
}

test('datas específicas da inscrição têm precedência sobre período textual e catálogo', () => {
  const fonte = trecho('montarPeriodoHistoricoVinculo_', 'obterLinhasDgmbDesafiosUsuario_');
  const datas = fonte.indexOf('if (periodoCompletoValido_(periodoDatasEspecificas))');
  const texto = fonte.indexOf('else if (periodoCompletoValido_(periodoTextoEspecifico))');
  const catalogo = fonte.indexOf('else if (periodoCompletoValido_(periodoLista))');

  assert.ok(datas >= 0 && texto > datas && catalogo > texto);
  assert.match(fonte, /origemPeriodo = 'dgmbDesafios\.data_inicio_desafio\/data_fim_desafio'/);
});

test('índice leve preserva janela por ID_INSCRICAO e restringe fallback por ID_DESAFIO ao legado', () => {
  const fonte = trecho('buildPeriodosDgmbDesafiosPorChave_', 'meuGiroResumoPossuiInscricaoAusente_');

  assert.match(fonte, /data_inicio_desafio/);
  assert.match(fonte, /data_fim_desafio/);
  assert.match(fonte, /detalhePorResumoKey/);
  assert.match(fonte, /inscricoesAptas/);
  assert.match(fonte, /if \(idDesafio && !idInscricao\)/);
});

test('resumo leve reconcilia vínculo ausente uma única vez sob lock', () => {
  const fonte = trecho('obterMeuGiroResumoAtualizadoLeve_', 'meuGiroResumoAgruparLinhasContiguas_');

  assert.match(fonte, /meuGiroResumoPossuiInscricaoAusente_/);
  assert.match(fonte, /LockService\.getScriptLock\(\)/);
  assert.match(fonte, /lock\.tryLock\(5000\)/);
  assert.match(fonte, /var resumoSobLock = shResumo\.getDataRange\(\)\.getValues\(\)/);
  assert.match(fonte, /atualizarMeuGiroResumo_\(id\)/);
});

test('inscrição identificada não usa período nem status de outra inscrição do mesmo desafio', () => {
  const fonte = trecho('obterMeuGiroResumoAtualizadoLeve_', 'meuGiroResumoAgruparLinhasContiguas_');

  assert.match(fonte, /var usarFallbackDesafio = !idInscricaoResumo;/);
  assert.match(fonte, /usarFallbackDesafio \? periodosDgmbDesafios\.byDesafio\[idDesafioResumo\] : ''/);
  assert.match(fonte, /usarFallbackDesafio \? periodosDgmbDesafios\.statusPorDesafio\[idDesafioResumo\] : null/);
  assert.match(fonte, /periodo_inicio: periodoInicioLeve/);
  assert.match(fonte, /periodo_fim: periodoFimLeve/);
});

test('cada atividade continua sendo avaliada independentemente contra o período de cada vínculo', () => {
  const fonte = trecho('atualizarMeuGiroResumo_', 'atualizarMeuGiroResumoEmLote_');

  assert.match(fonte, /for \(var v = 0; v < vinculos\.length; v\+\+\)/);
  assert.match(fonte, /for \(var r = 0; r < registros\.length; r\+\+\)/);
  assert.match(fonte, /atividadeDentroPeriodoOficial_\(reg\.data_atividade, inicio, fim\)/);
});
