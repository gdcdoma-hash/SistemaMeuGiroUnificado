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

test('período mensal usa texto, depois catálogo, e só então datas herdadas', () => {
  const fonte = trecho('montarPeriodoHistoricoVinculo_', 'obterLinhasDgmbDesafiosUsuario_');
  const texto = fonte.indexOf('if (periodoCompletoValido_(periodoTextoEspecifico))');
  const catalogo = fonte.indexOf('else if (periodoCompletoValido_(periodoLista))');
  const datas = fonte.indexOf('else if (periodoCompletoValido_(periodoDatasEspecificas))');

  assert.ok(texto >= 0 && catalogo > texto && datas > catalogo);
  assert.match(fonte, /origemPeriodo = 'dgmbDesafios\.periodo_desafio'/);
  assert.match(fonte, /origemPeriodo = 'ListaDesafios\.Periodo'/);
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

test('índice leve usa catálogo mensal antes das datas herdadas', () => {
  const fonte = trecho('buildPeriodosDgmbDesafiosPorChave_', 'meuGiroResumoPossuiInscricaoAusente_');

  assert.match(fonte, /var periodosLista = buildListaDesafiosContexto_\(getSpreadsheet_\(\)\)\.periodos/);
  assert.match(fonte, /var periodoLista = \(idDesafio && periodosLista\.byId\[idDesafio\]\)/);
  const texto = fonte.indexOf('var periodoDetalhe = periodoCompletoValido_(periodoTextoNormalizado)');
  const catalogo = fonte.indexOf(': periodoCompletoValido_(periodoLista)');
  const datas = fonte.indexOf(': periodoCompletoValido_(periodoDatas)');
  assert.ok(texto >= 0 && catalogo > texto && datas > catalogo);
});

test('índice leve cria alias composto apenas quando a chave identifica um único vínculo distinto', () => {
  const fonte = trecho('buildPeriodosDgmbDesafiosPorChave_', 'meuGiroResumoPossuiInscricaoAusente_');

  assert.match(fonte, /var aliasLegadoVinculos = \{\}/);
  assert.match(fonte, /var chaveAlias = meuGiroResumoBuildChave_\(id, idDesafioAlias, idItemAlias, metaAlias, ''\)/);
  assert.doesNotMatch(fonte, /if \(!idInscricaoAlias\) continue/);
  assert.match(fonte, /'INSCRICAO\|' \+ idInscricaoAlias/);
  assert.match(fonte, /'LEGADO'/);
  assert.match(fonte, /aliasLegadoVinculos\[chaveAlias\]\[assinaturaAlias\] = true/);
  assert.match(fonte, /aliasLegadoContagem\[chaveAliasContagem\] = Object\.keys\(aliasLegadoVinculos\[chaveAliasContagem\]\)\.length/);
  assert.match(fonte, /if \(chaveLegadaUnica && aliasLegadoContagem\[chaveLegadaUnica\] === 1\)/);
  assert.match(fonte, /periodos\.detalhePorResumoKey\[chaveLegadaUnica\] = periodoDetalhe/);
  assert.match(fonte, /periodos\.statusPorResumoKey\[chaveLegadaUnica\] = statusDgmb/);
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

test('leitor leve do painel usa texto, catálogo e só então datas individuais', () => {
  const painel = fs.readFileSync(path.resolve(__dirname, '..', 'Meu Giro', 'PainelService.gs'), 'utf8');
  const inicio = painel.indexOf('function painelMG_obterInscricaoLevePorDesafio_');
  const fim = painel.indexOf('\nfunction buscarInscricaoPainelMG_', inicio);
  const fonte = painel.slice(inicio, fim);

  assert.match(fonte, /var periodosLista = buildListaDesafiosContexto_\(getSpreadsheet_\(\)\)\.periodos/);
  assert.match(fonte, /var periodoLista = \(idDesafio && periodosLista\.byId\[idDesafio\]\) \|\| \{ inicio: '', fim: '' \}/);
  const texto = fonte.indexOf('var periodoSelecionado = periodoCompletoValido_(periodoTexto)');
  const catalogo = fonte.indexOf('? periodoTexto');
  const datas = fonte.indexOf(': periodoDatas;');
  assert.ok(texto >= 0 && catalogo > texto && datas > catalogo);
});

test('atualizador migra linha legada única para ID_INSCRICAO em vez de anexar duplicata', () => {
  const fonte = trecho('atualizarMeuGiroResumoComLockAdquirido_', 'atualizarMeuGiroResumoEmLote_');

  assert.match(fonte, /var quantidadeLinhasPorChave = \{\}/);
  assert.match(fonte, /quantidadeLinhasPorChave\[chaveExistente\] = \(quantidadeLinhasPorChave\[chaveExistente\] \|\| 0\) \+ 1/);
  assert.match(fonte, /var quantidadeVinculosPorChaveLegada = \{\}/);
  assert.match(fonte, /quantidadeVinculosPorChaveLegada\[chaveLegadaVinculo\] === 1/);
  assert.match(fonte, /quantidadeLinhasPorChave\[chaveLegadaVinculo\] === 1/);
  assert.match(fonte, /if \(linhaLegadaUnica\) numeroLinha = linhasPorChave\[chaveLegadaVinculo\] \|\| 0/);
  assert.match(fonte, /if \(idxInscricaoResumo > -1\) linha\[idxInscricaoResumo\] = idInscricao/);
});
