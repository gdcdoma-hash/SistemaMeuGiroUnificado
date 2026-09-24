var PORTAL_HANDOFF_SHEET_ = '_PORTAL_HANDOFF';

function portalHandoffNormalizar_(valor) {
  return String(valor || '').trim();
}

function portalHandoffObterSheet_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(PORTAL_HANDOFF_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(PORTAL_HANDOFF_SHEET_);
    sh.getRange(1, 1, 1, 7).setValues([[
      'TOKEN', 'ID_DGMB', 'ESCOPO', 'CRIADO_EM', 'EXPIRA_EM', 'STATUS', 'ORIGEM'
    ]]);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}

function portalHandoffValidar_(token, escopo) {
  var alvo = portalHandoffNormalizar_(token);
  var scope = portalHandoffNormalizar_(escopo).toUpperCase();
  if (!alvo) return { ok: false, code: 'HANDOFF_AUSENTE', msg: 'Acesso integrado ausente.' };

  var sh = portalHandoffObterSheet_();
  var values = sh.getDataRange().getValues();
  var agora = Date.now();

  for (var i = values.length - 1; i >= 1; i--) {
    var row = values[i] || [];
    if (portalHandoffNormalizar_(row[0]) !== alvo) continue;

    var id = normalizeText_(row[1]);
    var escopoLinha = portalHandoffNormalizar_(row[2]).toUpperCase();
    var expira = row[4] instanceof Date ? row[4].getTime() : Number(row[4]);
    var status = portalHandoffNormalizar_(row[5]).toUpperCase();

    if (status !== 'ATIVO') return { ok: false, code: 'HANDOFF_INATIVO', msg: 'Acesso integrado encerrado.' };
    if (!expira || agora >= expira) return { ok: false, code: 'HANDOFF_EXPIRADO', msg: 'Acesso integrado expirado.' };
    if (scope && escopoLinha !== scope) return { ok: false, code: 'HANDOFF_ESCOPO', msg: 'Acesso integrado não autorizado.' };
    if (!id) return { ok: false, code: 'HANDOFF_SEM_ID', msg: 'Identidade do acesso integrado ausente.' };

    return { ok: true, id_dgmb: id, escopo: escopoLinha, expira_em: expira };
  }

  return { ok: false, code: 'HANDOFF_INVALIDO', msg: 'Acesso integrado inválido.' };
}

function portalHandoffValidarAdmin_(token) {
  return portalHandoffValidar_(token, 'ADMIN');
}
