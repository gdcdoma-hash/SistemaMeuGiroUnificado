var ATLETA_SESSION_SHEET_ = '_MEU_GIRO_SESSOES';
var ATLETA_SESSION_TTL_MS_ = 6 * 60 * 60 * 1000;
var ATLETA_SESSION_RETENCAO_MS_ = 24 * 60 * 60 * 1000;

function atletaSessionNormalizar_(valor) { return String(valor || '').trim(); }

function atletaSessionBuscarPessoaPorId_(idDgmb) {
  var id = normalizeText_(idDgmb);
  if (!id) return null;
  try {
    var sh = getSheetByName_(SHEETS.PESSOAS);
    var values = sh.getDataRange().getValues();
    if (!values || values.length < 2) return null;
    var map = buildHeaderMap_(values[0]);
    var idxId = getRequiredColumnIndex_(map, ['id_dgmb'], SHEETS.PESSOAS);
    var idxNome = getRequiredColumnIndex_(map, ['nome'], SHEETS.PESSOAS);
    var idxCidade = getOptionalColumnIndex_(map, ['cidade-uf', 'cidade_uf', 'cidade uf']);
    for (var i = 1; i < values.length; i++) {
      if (normalizeText_(values[i][idxId]) !== id) continue;
      return { id_dgmb: id, nome: normalizeText_(values[i][idxNome]), cidade_uf: idxCidade > -1 ? normalizeText_(values[i][idxCidade]) : '' };
    }
  } catch (e) { Logger.log('[ATLETA_SESSION] falha ao buscar nome: ' + (e && e.message ? e.message : e)); }
  return null;
}

function atletaSessionSheet_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(ATLETA_SESSION_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(ATLETA_SESSION_SHEET_);
    sh.getRange(1, 1, 1, 7).setValues([['SESSION_TOKEN', 'ID_DGMB', 'CRIADO_EM', 'EXPIRA_EM', 'STATUS', 'ORIGEM', 'HANDOFF_TOKEN']]);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}

function atletaSessionLimparAntigas_(sh, agora) {
  var last = sh.getLastRow(); if (last < 2) return;
  var rows = sh.getRange(2, 1, last - 1, 7).getValues(); var limite = agora - ATLETA_SESSION_RETENCAO_MS_;
  for (var i = rows.length - 1; i >= 0; i--) { var expira = rows[i][3] instanceof Date ? rows[i][3].getTime() : Number(rows[i][3]); if (expira && expira < limite) sh.deleteRow(i + 2); }
}

function atletaTrocarHandoffPorSessao(handoffToken) {
  var token = atletaSessionNormalizar_(handoffToken);
  if (!token) return { ok: false, code: 'HANDOFF_AUSENTE', msg: 'Acesso integrado ausente.' };
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var handoffSh = portalHandoffObterSheet_(); var values = handoffSh.getDataRange().getValues(); var agora = Date.now(); var linha = -1; var idDgmb = '';
    for (var i = values.length - 1; i >= 1; i--) {
      var row = values[i] || []; if (atletaSessionNormalizar_(row[0]) !== token) continue;
      linha = i + 1; idDgmb = normalizeText_(row[1]); var escopo = atletaSessionNormalizar_(row[2]).toUpperCase(); var expira = row[4] instanceof Date ? row[4].getTime() : Number(row[4]); var status = atletaSessionNormalizar_(row[5]).toUpperCase();
      if (escopo !== 'ATLETA') return { ok: false, code: 'HANDOFF_ESCOPO', msg: 'Acesso integrado não autorizado.' };
      if (status !== 'ATIVO') return { ok: false, code: 'HANDOFF_REUTILIZADO', msg: 'Este acesso integrado já foi utilizado.' };
      if (!expira || agora >= expira) { handoffSh.getRange(linha, 6).setValue('EXPIRADO'); return { ok: false, code: 'HANDOFF_EXPIRADO', msg: 'Acesso integrado expirado.' }; }
      if (!idDgmb) return { ok: false, code: 'HANDOFF_SEM_ID', msg: 'Identidade ausente.' }; break;
    }
    if (linha < 0) return { ok: false, code: 'HANDOFF_INVALIDO', msg: 'Acesso integrado inválido.' };
    var sessionToken = Utilities.getUuid() + '-' + Utilities.getUuid(); var sessionExpira = agora + ATLETA_SESSION_TTL_MS_; var sessionSh = atletaSessionSheet_(); atletaSessionLimparAntigas_(sessionSh, agora);
    sessionSh.appendRow([sessionToken,idDgmb,new Date(agora),new Date(sessionExpira),'ATIVA','PORTAL_ATLETA',token]); handoffSh.getRange(linha, 6).setValue('CONSUMIDO');
    var pessoa = atletaSessionBuscarPessoaPorId_(idDgmb) || {id_dgmb:idDgmb,nome:'',cidade_uf:''};
    return { ok:true, session_token:sessionToken, expira_em:sessionExpira, usuario:{id_dgmb:idDgmb,nome:pessoa.nome||'',cidade_uf:pessoa.cidade_uf||''} };
  } finally { lock.releaseLock(); }
}

function atletaValidarSessao(sessionToken) {
  var token = atletaSessionNormalizar_(sessionToken); if (!token) return { ok:false, code:'SESSAO_AUSENTE', msg:'Sessão ausente.' };
  var sh=atletaSessionSheet_(),values=sh.getDataRange().getValues(),agora=Date.now();
  for(var i=values.length-1;i>=1;i--){var row=values[i]||[];if(atletaSessionNormalizar_(row[0])!==token)continue;var id=normalizeText_(row[1]),expira=row[3] instanceof Date?row[3].getTime():Number(row[3]),status=atletaSessionNormalizar_(row[4]).toUpperCase();if(status!=='ATIVA')return{ok:false,code:'SESSAO_INATIVA',msg:'Sessão encerrada.'};if(!expira||agora>=expira){sh.getRange(i+1,5).setValue('EXPIRADA');return{ok:false,code:'SESSAO_EXPIRADA',msg:'Sessão expirada.'}}if(!id)return{ok:false,code:'SESSAO_SEM_ID',msg:'Identidade da sessão ausente.'};return{ok:true,id_dgmb:id,expira_em:expira}}
  return {ok:false,code:'SESSAO_INVALIDA',msg:'Sessão inválida.'};
}

function atletaEncerrarSessao(sessionToken) {
  var token=atletaSessionNormalizar_(sessionToken);if(!token)return{ok:true};var sh=atletaSessionSheet_(),values=sh.getDataRange().getValues();for(var i=values.length-1;i>=1;i--){if(atletaSessionNormalizar_(values[i][0])===token){sh.getRange(i+1,5).setValue('ENCERRADA');break}}return{ok:true};
}
