function portalAdminMeuGiroNormalizarBusca_(valor) {
  return String(valor || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function portalAdminMeuGiroAutorizar_(token) {
  var acesso = portalHandoffValidarAdmin_(token);
  if (!acesso || acesso.ok !== true) {
    return acesso || { ok: false, code: 'HANDOFF_NEGADO', msg: 'Acesso administrativo negado.' };
  }
  return acesso;
}

function portalAdminMeuGiroBuscarAtletas(token, termo) {
  var acesso = portalAdminMeuGiroAutorizar_(token);
  if (!acesso.ok) return acesso;

  var buscaBruta = String(termo || '').trim();
  if (!buscaBruta) {
    return { ok: false, code: 'TERMO_OBRIGATORIO', msg: 'Informe Cód, CPF, ID_DGMB ou nome.' };
  }

  var sh = getSheetByName_(SHEETS.PESSOAS);
  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { ok: true, data: [] };
  }

  var map = buildHeaderMap_(values[0]);
  var idxId = getRequiredColumnIndex_(map, ['id_dgmb'], SHEETS.PESSOAS);
  var idxCpf = getRequiredColumnIndex_(map, ['cpf'], SHEETS.PESSOAS);
  var idxNome = getRequiredColumnIndex_(map, ['nome'], SHEETS.PESSOAS);
  var idxCidade = getOptionalColumnIndex_(map, ['cidade-uf', 'cidade_uf', 'cidade uf']);
  var idxCod = getOptionalColumnIndex_(map, ['cod', 'cód', 'codigo', 'código', 'cod_atleta', 'codigo_atleta']);

  var digitos = onlyDigits_(buscaBruta);
  var cpfBase64 = digitos.length === 11 ? Utilities.base64Encode(digitos) : '';
  var termoNormalizado = portalAdminMeuGiroNormalizarBusca_(buscaBruta);
  var resultados = [];

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var id = normalizeText_(row[idxId]);
    var nome = normalizeText_(row[idxNome]);
    var cpfSalvo = normalizeText_(row[idxCpf]);
    var cidade = idxCidade > -1 ? normalizeText_(row[idxCidade]) : '';
    var cod = idxCod > -1 ? normalizeText_(row[idxCod]) : '';

    var encontrou = false;
    if (cpfBase64 && cpfSalvo === cpfBase64) encontrou = true;
    if (!encontrou && cod && portalAdminMeuGiroNormalizarBusca_(cod) === termoNormalizado) encontrou = true;
    if (!encontrou && id && portalAdminMeuGiroNormalizarBusca_(id) === termoNormalizado) encontrou = true;
    if (!encontrou && nome && portalAdminMeuGiroNormalizarBusca_(nome).indexOf(termoNormalizado) !== -1) encontrou = true;

    if (!encontrou) continue;

    resultados.push({
      id_dgmb: id,
      cod: cod,
      nome: nome,
      cidade_uf: cidade
    });

    if (resultados.length >= 20) break;
  }

  return {
    ok: true,
    data: resultados,
    total_retornado: resultados.length,
    limite: 20
  };
}

function portalAdminMeuGiroConsultarAtleta(token, idDgmb) {
  var acesso = portalAdminMeuGiroAutorizar_(token);
  if (!acesso.ok) return acesso;

  var id = normalizeText_(idDgmb);
  if (!id) {
    return { ok: false, code: 'ID_DGMB_OBRIGATORIO', msg: 'ID_DGMB obrigatório.' };
  }

  var painel = getPainelUsuario(id, { somenteLeitura: true });
  if (!painel || painel.ok === false) {
    return painel || { ok: false, code: 'PAINEL_NAO_DISPONIVEL', msg: 'Painel não disponível.' };
  }

  return {
    ok: true,
    id_dgmb: id,
    atleta: painel.atleta || {},
    desafio_em_foco: painel.desafio_em_foco || null,
    desafios: painel.desafios || [],
    desafios_ativos: painel.desafios_ativos || [],
    desafios_historico: painel.desafios_historico || [],
    atividades: painel.atividades || [],
    total_pedalado: painel.total_pedalado || painel.totalPedalado || 0,
    somente_leitura: true,
    administrador_id_dgmb: acesso.id_dgmb
  };
}
