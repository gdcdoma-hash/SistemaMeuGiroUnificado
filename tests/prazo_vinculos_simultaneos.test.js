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
  assert.match(fonte, /atualizarMeuGiroResumoComLockAdquirido_\(id\)/);
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
  const fonte = trecho('atualizarMeuGiroResumoComLockAdquirido_', 'atualizarMeuGiroResumoEmLote_');

  assert.match(fonte, /for \(var v = 0; v < vinculos\.length; v\+\+\)/);
  assert.match(fonte, /for \(var r = 0; r < registros\.length; r\+\+\)/);
  assert.match(fonte, /atividadeDentroPeriodoOficial_\(reg\.data_atividade, inicio, fim\)/);
});

test('toda escrita do resumo passa pelo wrapper de lock e fluxos já travados usam a implementação interna', () => {
  const wrapper = trecho('atualizarMeuGiroResumo_', 'atualizarMeuGiroResumoComLockAdquirido_');
  const interna = trecho('atualizarMeuGiroResumoComLockAdquirido_', 'atualizarMeuGiroResumoEmLote_');
  const registro = fs.readFileSync(path.resolve(__dirname, '..', 'Meu Giro', 'RegistroService.gs'), 'utf8');

  assert.match(wrapper, /LockService\.getScriptLock\(\)/);
  assert.match(wrapper, /lock\.waitLock\(30000\)/);
  assert.match(wrapper, /return atualizarMeuGiroResumoComLockAdquirido_\(idDgmb, opcoes\)/);
  assert.match(interna, /shResumo\.appendRow\(linha\)/);
  assert.doesNotMatch(wrapper, /appendRow|setValues/);
  assert.doesNotMatch(registro, /atualizarMeuGiroResumo_\(idDgmb, opcoesRegistroKm\)/);
  assert.match(registro, /atualizarMeuGiroResumoComLockAdquirido_\(idDgmb, opcoesRegistroKm\)/);
});

test('caminho pesado seleciona inscrição exata e só usa desafio mais item para vínculo legado', () => {
  const painel = fs.readFileSync(path.resolve(__dirname, '..', 'Meu Giro', 'PainelService.gs'), 'utf8');
  const inicio = painel.indexOf('function painelMG_buscarVinculoPrincipal_');
  const fim = painel.indexOf('\nfunction painelMG_parseDataISO_', inicio);
  const fonte = painel.slice(inicio, fim);

  assert.match(fonte, /var idInscricaoPrincipal = painelMG_norm_\(desafioPrincipal && desafioPrincipal\.id_inscricao\)/);
  assert.match(fonte, /idInscricaoPrincipal\s*\? painelMG_norm_\(v\.id_inscricao\) === idInscricaoPrincipal\s*: painelMG_norm_\(v\.id_desafio\) === idDesafioPrincipal/);
});

test('leitor leve do painel prefere o par de datas individuais ao período textual', () => {
  const painel = fs.readFileSync(path.resolve(__dirname, '..', 'Meu Giro', 'PainelService.gs'), 'utf8');
  const inicio = painel.indexOf('function painelMG_obterInscricaoLevePorDesafio_');
  const fim = painel.indexOf('\nfunction buscarInscricaoPainelMG_', inicio);
  const fonte = painel.slice(inicio, fim);

  assert.match(fonte, /var periodoSelecionado = periodoCompletoValido_\(periodoDatas\) \? periodoDatas : periodoTexto/);
});
