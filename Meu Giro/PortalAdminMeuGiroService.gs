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

  var dados = painel.data || painel;
  var todosDesafios = Array.isArray(dados.desafios) ? dados.desafios : [];
  var desafiosAtivos = Array.isArray(dados.desafios_ativos) ? dados.desafios_ativos : [];
  var desafiosHistorico = Array.isArray(dados.desafios_historico) ? dados.desafios_historico.slice() : [];
  var chavesAtivas = {};
  var chavesHistorico = {};

  desafiosAtivos.forEach(function(item) {
    chavesAtivas[painelMG_chaveDesafioPainel_(item)] = true;
  });
  desafiosHistorico.forEach(function(item) {
    chavesHistorico[painelMG_chaveDesafioPainel_(item)] = true;
  });
  todosDesafios.forEach(function(item) {
    var chave = painelMG_chaveDesafioPainel_(item);
    if (!chavesAtivas[chave] && !chavesHistorico[chave]) {
      desafiosHistorico.push(item);
      chavesHistorico[chave] = true;
    }
  });
  desafiosHistorico.sort(painelMG_compareHistoricoDesafios_);

  return {
    ok: true,
    id_dgmb: dados.id_dgmb || id,
    atleta: {
      nome: dados.nome || '',
      cidade_uf: dados.cidade_uf || '',
      id_dgmb: dados.id_dgmb || id
    },
    desafio_em_foco: dados.desafio_em_foco || null,
    desafios: dados.desafios || [],
    desafios_ativos: desafiosAtivos,
    desafios_historico: desafiosHistorico,
    atividades: dados.atividades || [],
    auditorias: portalAdminMeuGiroListarAuditorias_(dados.id_dgmb || id),
    total_pedalado: dados.total_pedalado || dados.totalPedalado || dados.realizado || 0,
    somente_leitura: true,
    administrador_id_dgmb: acesso.id_dgmb
  };
}


function portalAdminMeuGiroAuditoriaSheet_() {
  var ss = getSpreadsheet_();
  var nomeAba = '_MEU_GIRO_AUDITORIA_ADMIN';
  var sh = ss.getSheetByName(nomeAba);
  if (!sh) {
    sh = ss.insertSheet(nomeAba);
    sh.getRange(1, 1, 1, 12).setValues([[
      'ID_AUDITORIA', 'DATA_HORA', 'ADMIN_ID_DGMB', 'ATLETA_ID_DGMB',
      'ACAO', 'ACTIVITY_ID', 'CHAVE_EDICAO', 'DATA_ANTERIOR',
      'KM_ANTERIOR', 'DATA_NOVA', 'KM_NOVO', 'MOTIVO'
    ]]);
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}

function portalAdminMeuGiroLocalizarAtividade_(idDgmb, activityId, chaveEdicao) {
  var atividades = buscarAtividadesUsuario_(idDgmb) || [];
  for (var i = 0; i < atividades.length; i++) {
    var item = atividades[i] || {};
    if (activityId && String(item.activity_id || '').trim() === activityId) return item;
    if (!activityId && chaveEdicao && String(item.chave_edicao || '').trim() === chaveEdicao) return item;
  }
  return null;
}

function portalAdminMeuGiroEditarAtividade(token, payload) {
  var acesso = portalAdminMeuGiroAutorizar_(token);
  if (!acesso.ok) return acesso;

  var dados = payload && typeof payload === 'object' ? payload : {};
  var idDgmb = normalizeText_(dados.id_dgmb);
  var activityId = normalizeText_(dados.activity_id);
  var chaveEdicao = normalizeText_(dados.chave_edicao);
  var motivo = String(dados.motivo || '').trim();
  var novaData = normalizarDataISO_(dados.data_atividade);
  var novoKm = painelMG_toNumber_(dados.km);

  if (!idDgmb) return { ok: false, code: 'ID_DGMB_OBRIGATORIO', msg: 'Atleta não informado.' };
  if (!activityId && !chaveEdicao) return { ok: false, code: 'ATIVIDADE_SEM_IDENTIFICADOR', msg: 'Atividade sem identificador de edição.' };
  if (!novaData) return { ok: false, code: 'DATA_INVALIDA', msg: 'Informe uma data válida.' };
  if (!(novoKm > 0)) return { ok: false, code: 'KM_INVALIDO', msg: 'Informe um KM maior que zero.' };
  if (motivo.length < 5) return { ok: false, code: 'MOTIVO_OBRIGATORIO', msg: 'Informe o motivo da correção com pelo menos 5 caracteres.' };

  var anterior = portalAdminMeuGiroLocalizarAtividade_(idDgmb, activityId, chaveEdicao);
  if (!anterior) return { ok: false, code: 'ATIVIDADE_NAO_ENCONTRADA', msg: 'Atividade não encontrada para correção.' };

  var resultado = editarAtividade({
    id_dgmb: idDgmb,
    activity_id: activityId,
    chave_edicao: chaveEdicao,
    data_atividade: novaData,
    km: novoKm
  });
  if (!resultado || resultado.ok !== true) return resultado || { ok: false, code: 'EDICAO_FALHOU', msg: 'Não foi possível corrigir a atividade.' };

  var auditoriaId = Utilities.getUuid();
  portalAdminMeuGiroAuditoriaSheet_().appendRow([
    auditoriaId,
    new Date(),
    acesso.id_dgmb || '',
    idDgmb,
    'CORRECAO_ATIVIDADE',
    activityId,
    chaveEdicao,
    anterior.data || '',
    anterior.km || 0,
    novaData,
    novoKm,
    motivo
  ]);

  var painelAtualizado = portalAdminMeuGiroConsultarAtleta(token, idDgmb);
  painelAtualizado.auditoria_id = auditoriaId;
  painelAtualizado.msg = 'Atividade corrigida e registrada na auditoria.';
  return painelAtualizado;
}


function portalAdminMeuGiroListarAuditorias_(idDgmb) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName('_MEU_GIRO_AUDITORIA_ADMIN');
  if (!sh || sh.getLastRow() < 2) return [];

  var values = sh.getDataRange().getDisplayValues();
  var id = normalizeText_(idDgmb);
  var out = [];

  for (var i = values.length - 1; i >= 1; i--) {
    var row = values[i] || [];
    if (normalizeText_(row[3]) !== id) continue;
    out.push({
      id_auditoria: row[0] || '',
      data_hora: row[1] || '',
      admin_id_dgmb: row[2] || '',
      acao: row[4] || '',
      activity_id: row[5] || '',
      data_anterior: row[7] || '',
      km_anterior: row[8] || '',
      data_nova: row[9] || '',
      km_novo: row[10] || '',
      motivo: row[11] || ''
    });
    if (out.length >= 50) break;
  }
  return out;
}
