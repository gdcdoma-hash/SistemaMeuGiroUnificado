from pathlib import Path


def replace(path, old, new, count=1):
    p = Path(path)
    text = p.read_text(encoding='utf-8')
    actual = text.count(old)
    if actual != count:
        raise SystemExit(f'{path}: expected {count} occurrence(s), found {actual}: {old[:100]!r}')
    p.write_text(text.replace(old, new, count), encoding='utf-8')


replace('Meu Giro/PainelService.gs',
"""    var periodoTexto = idxPeriodo > -1 ? extrairPeriodoDesafioTexto_(row[idxPeriodo]) : { inicio: '', fim: '' };
    var inicio = periodoCompletoValido_(periodoTexto) ? periodoTexto.inicio : normalizarDataISO_(idxInicio > -1 ? row[idxInicio] : '');
    var fim = periodoCompletoValido_(periodoTexto) ? periodoTexto.fim : normalizarDataISO_(idxFim > -1 ? row[idxFim] : '');
""",
"""    var periodoDatas = {
      inicio: normalizarDataISO_(idxInicio > -1 ? row[idxInicio] : ''),
      fim: normalizarDataISO_(idxFim > -1 ? row[idxFim] : '')
    };
    var periodoTexto = idxPeriodo > -1 ? extrairPeriodoDesafioTexto_(row[idxPeriodo]) : { inicio: '', fim: '' };
    var periodoSelecionado = periodoCompletoValido_(periodoDatas) ? periodoDatas : periodoTexto;
    var inicio = periodoCompletoValido_(periodoSelecionado) ? periodoSelecionado.inicio : '';
    var fim = periodoCompletoValido_(periodoSelecionado) ? periodoSelecionado.fim : '';
""")

replace('Meu Giro/PainelService.gs',
"""    var vinculos = obterVinculosDesafioUsuario_(idDgmb) || [];
    var idDesafioPrincipal = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_desafio);
    var idItemPrincipal = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_item_estoque);

    for (var i = 0; i < vinculos.length; i++) {
      var v = vinculos[i] || {};
      if (
        painelMG_norm_(v.id_desafio) === idDesafioPrincipal &&
        painelMG_norm_(v.id_item_estoque) === idItemPrincipal
      ) {
""",
"""    var vinculos = obterVinculosDesafioUsuario_(idDgmb) || [];
    var idInscricaoPrincipal = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_inscricao);
    var idDesafioPrincipal = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_desafio);
    var idItemPrincipal = painelMG_norm_(desafioPrincipal && desafioPrincipal.id_item_estoque);

    for (var i = 0; i < vinculos.length; i++) {
      var v = vinculos[i] || {};
      var corresponde = idInscricaoPrincipal
        ? painelMG_norm_(v.id_inscricao) === idInscricaoPrincipal
        : painelMG_norm_(v.id_desafio) === idDesafioPrincipal &&
          painelMG_norm_(v.id_item_estoque) === idItemPrincipal;
      if (corresponde) {
""")

replace('Meu Giro/RegistroService.gs',
        '      atualizarMeuGiroResumo_(idDgmb, opcoesRegistroKm);\n',
        '      atualizarMeuGiroResumoComLockAdquirido_(idDgmb, opcoesRegistroKm);\n', 3)

replace('Meu Giro/Utils.gs',
        '          atualizarMeuGiroResumo_(id);\n          return obterMeuGiroResumoAtualizadoLeve_(id);\n',
        '          atualizarMeuGiroResumoComLockAdquirido_(id);\n          return obterMeuGiroResumoAtualizadoLeve_(id);\n')

replace('Meu Giro/Utils.gs',
        'function atualizarMeuGiroResumo_(idDgmb, opcoes) {\n  var perfTotalInicio = meuGiroPerfNow_();\n',
"""function atualizarMeuGiroResumo_(idDgmb, opcoes) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return atualizarMeuGiroResumoComLockAdquirido_(idDgmb, opcoes);
  } finally {
    lock.releaseLock();
  }
}

function atualizarMeuGiroResumoComLockAdquirido_(idDgmb, opcoes) {
  var perfTotalInicio = meuGiroPerfNow_();
""")

replace('tests/atividade_performance_instrumentation.test.js',
        '  assert.match(registroService, /atualizarMeuGiroResumo_\\(idDgmb, opcoesRegistroKm\\)/);\n',
        '  assert.match(registroService, /atualizarMeuGiroResumoComLockAdquirido_\\(idDgmb, opcoesRegistroKm\\)/);\n')

replace('tests/lista_desafios_nome_backend.test.js',
        '  assert.match(light, /var statusDgmbResumo = periodosDgmbDesafios\\.statusPorResumoKey\\[chaveResumo\\] \\|\\| periodosDgmbDesafios\\.statusPorDesafio\\[idDesafioResumo\\] \\|\\| \\{\\}/);\n',
        '  assert.match(light, /var usarFallbackDesafio = !idInscricaoResumo/);\n  assert.match(light, /var statusDgmbResumo = periodosDgmbDesafios\\.statusPorResumoKey\\[chaveResumo\\] \\|\\| \\(usarFallbackDesafio \\? periodosDgmbDesafios\\.statusPorDesafio\\[idDesafioResumo\\] : null\\) \\|\\| \\{\\}/);\n')

replace('tests/lista_desafios_nome_backend.test.js',
"""    getIdDesafioColumnIndex_(map) { return map.id_desafio ?? -1; },
    obterIdDesafioRegistro_(row, idxIdDesafio) { return idxIdDesafio > -1 ? ctx.normalizeText_(row[idxIdDesafio]) : ''; }
""",
"""    getIdDesafioColumnIndex_(map) { return map.id_desafio ?? -1; },
    obterIdDesafioRegistro_(row, idxIdDesafio) { return idxIdDesafio > -1 ? ctx.normalizeText_(row[idxIdDesafio]) : ''; },
    extrairPeriodoDesafioTexto_() { return { inicio: '2026-06-01', fim: '2026-06-30' }; },
    periodoCompletoValido_(periodo) { return !!(periodo && periodo.inicio && periodo.fim); },
    normalizarDataISO_() { return ''; },
    validarInscricaoMinima_() { return { valida: true }; },
    inscricaoTemBloqueioMinimo_() { return false; }
""")

replace('tests/meu_giro_resumo_chave_legada.test.js',
        "const inicioAtualizar = utils.indexOf('function atualizarMeuGiroResumo_(idDgmb, opcoes)');\n",
        "const inicioAtualizar = utils.indexOf('function atualizarMeuGiroResumoComLockAdquirido_(idDgmb, opcoes)');\n")
replace('tests/meu_giro_resumo_chave_legada.test.js',
        "const fim = utils.indexOf('\\n}\\n\\n\\n\\nfunction calcularStatusMeuGiroPorPercentual_', inicio) + 3;\n",
        "const fim = utils.indexOf('\\nfunction calcularStatusMeuGiroPorPercentual_', inicio);\n")

replace('tests/meu_giro_resumo_leitura_cirurgica.test.js',
        "const inicio = utils.indexOf('function atualizarMeuGiroResumo_(idDgmb, opcoes)');\n",
        "const inicio = utils.indexOf('function atualizarMeuGiroResumoComLockAdquirido_(idDgmb, opcoes)');\n")

replace('tests/meu_giro_status_apuracao_origem.test.js',
        "const inicio = utils.indexOf('function atualizarMeuGiroResumo_(idDgmb, opcoes)');\n",
        "const inicio = utils.indexOf('function atualizarMeuGiroResumoComLockAdquirido_(idDgmb, opcoes)');\n")

replace('tests/meu_giro_resumo_sincronizacao_fluxos.test.js',
        "const atualizarResumo = sliceFunction(utils, 'atualizarMeuGiroResumo_', 'atualizarMeuGiroResumoEmLote_');\n",
        "const atualizarResumo = sliceFunction(utils, 'atualizarMeuGiroResumoComLockAdquirido_', 'atualizarMeuGiroResumoEmLote_');\n")
replace('tests/meu_giro_resumo_sincronizacao_fluxos.test.js',
        'atualizarMeuGiroResumo_\\(idDgmb, opcoesRegistroKm\\);/',
        'atualizarMeuGiroResumoComLockAdquirido_\\(idDgmb, opcoesRegistroKm\\);/')
replace('tests/meu_giro_resumo_sincronizacao_fluxos.test.js',
        "test('atualizarMeuGiroResumo_ permanece responsável por gravar MEU_GIRO_RESUMO a partir de vínculos dgmbDesafios', () => {\n",
        "test('implementação interna permanece responsável por gravar MEU_GIRO_RESUMO a partir de vínculos dgmbDesafios', () => {\n")

p = Path('tests/prazo_vinculos_simultaneos.test.js')
text = p.read_text(encoding='utf-8')
old = '/atualizarMeuGiroResumo_\\(id\\)/'
if text.count(old) != 1:
    raise SystemExit('prazo_vinculos_simultaneos: updater reconciliation expectation mismatch')
text = text.replace(old, '/atualizarMeuGiroResumoComLockAdquirido_\\(id\\)/', 1)
old = "trecho('atualizarMeuGiroResumo_', 'atualizarMeuGiroResumoEmLote_')"
if text.count(old) != 1:
    raise SystemExit('prazo_vinculos_simultaneos: updater slice expectation mismatch')
text = text.replace(old, "trecho('atualizarMeuGiroResumoComLockAdquirido_', 'atualizarMeuGiroResumoEmLote_')", 1)
addition = r'''

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
'''
if 'toda escrita do resumo passa pelo wrapper de lock' in text:
    raise SystemExit('prazo_vinculos_simultaneos: focused additions already present')
p.write_text(text.rstrip() + addition + '\n', encoding='utf-8')
