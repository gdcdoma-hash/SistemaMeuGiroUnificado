function getRanking(idDgmb, idDesafio, idItemEstoque) {
  try {
    var idUsuario = rankingMG_norm_(idDgmb);
    var desafioSolicitado = rankingMG_norm_(idDesafio);
    var itemSolicitado = rankingMG_norm_(idItemEstoque);
    if (!idUsuario) {
      return { ok: false, data: [], total: 0, msg: 'ID do usuário não informado.' };
    }

    var pessoas = getAllObjects_(SHEETS.PESSOAS);
    var resumo = getAllObjects_(SHEETS.MEU_GIRO_RESUMO) || [];

    if (!resumo.length) {
      return {
        ok: true,
        data: [],
        total: 0,
        msg: 'Nenhum atleta encontrado no ranking.'
      };
    }

    var statusValidos = { ATIVO: true, CONCLUIDO: true };
    var referencia = null;

    for (var r = 0; r < resumo.length; r++) {
      var rowRef = resumo[r] || {};
      var rowIdRef = rankingMG_norm_(rankingMG_firstFilled_(rowRef, ['ID_DGMB', 'id_dgmb']));
      if (rowIdRef !== idUsuario) continue;

      var statusRef = rankingMG_norm_(rankingMG_firstFilled_(rowRef, ['Status_Apuracao', 'status_apuracao'])).toUpperCase();
      if (!statusValidos[statusRef]) continue;

      var rowDesafioRef = rankingMG_norm_(rankingMG_firstFilled_(rowRef, ['ID_DESAFIO', 'id_desafio']));
      var rowItemRef = rankingMG_norm_(rankingMG_firstFilled_(rowRef, ['id_item_estoque', 'id item estoque']));
      if (desafioSolicitado && rowDesafioRef !== desafioSolicitado) continue;
      if (desafioSolicitado && itemSolicitado && rowItemRef !== itemSolicitado) continue;

      if (statusRef === 'ATIVO') {
        referencia = rowRef;
        break;
      }

      if (!referencia) referencia = rowRef;
    }

    if (!referencia) {
      return {
        ok: true,
        data: [],
        total: 0,
        msg: 'Usuário sem desafio elegível para o ranking.'
      };
    }

    var desafioPrincipal = rankingMG_norm_(rankingMG_firstFilled_(referencia, ['ID_DESAFIO', 'id_desafio']));
    var itemPrincipal = rankingMG_norm_(rankingMG_firstFilled_(referencia, ['id_item_estoque', 'id item estoque']));
    var grupoBasePrincipal = rankingMG_extrairGrupoBaseDesafio_(itemPrincipal);

    if (!desafioPrincipal) {
      return {
        ok: true,
        data: [],
        total: 0,
        msg: 'Desafio-base não identificado para o ranking.'
      };
    }

    var pessoasMap = rankingMG_buildPessoasMap_(pessoas);
    var ranking = [];

    for (var i = 0; i < resumo.length; i++) {
      var row = resumo[i] || {};

      var idDgmb = rankingMG_norm_(rankingMG_firstFilled_(row, ['ID_DGMB', 'id_dgmb']));
      if (!idDgmb) continue;

      var rowDesafio = rankingMG_norm_(rankingMG_firstFilled_(row, ['ID_DESAFIO', 'id_desafio']));
      var rowItem = rankingMG_norm_(rankingMG_firstFilled_(row, ['id_item_estoque', 'id item estoque']));
      var rowGrupoBase = rankingMG_extrairGrupoBaseDesafio_(rowItem);
      var rowStatus = rankingMG_norm_(rankingMG_firstFilled_(row, ['Status_Apuracao', 'status_apuracao'])).toUpperCase();

      if (rowDesafio !== desafioPrincipal) continue;
      if (grupoBasePrincipal && rowGrupoBase !== grupoBasePrincipal) continue;
      if (!statusValidos[rowStatus]) continue;

      var meta = rankingMG_round1_(rankingMG_toNumber_(rankingMG_firstFilled_(row, [
        'Distancia_KM', 'distancia_km', 'Distancia KM', 'Meta_KM', 'meta_km', 'meta'
      ])));

      var realizado = rankingMG_round1_(rankingMG_toNumber_(rankingMG_firstFilled_(row, [
        'Distancia_Realizada', 'distancia_realizada', 'Distancia Realizada'
      ])));

      var restante = rankingMG_round1_(Math.max(meta - realizado, 0));
      var percentual = rankingMG_round1_(rankingMG_toNumber_(rankingMG_firstFilled_(row, [
        'Percentual_Concluido', 'percentual_concluido'
      ])));
      if (percentual <= 0 && meta > 0) {
        percentual = rankingMG_round1_((realizado / meta) * 100);
      }

      var pessoa = pessoasMap[idDgmb] || {};

      ranking.push({
        id_dgmb: idDgmb,
        nome: pessoa.nome || rankingMG_norm_(rankingMG_firstFilled_(row, ['Nome_Avatar', 'nome_avatar'])) || 'Participante',
        cidade_uf: pessoa.cidade_uf || '',
        distancia_realizada: realizado,
        percentual_concluido: percentual,
        meta: meta,
        realizado: realizado,
        restante: restante,
        percentual: percentual
      });
    }

    ranking.sort(function(a, b) {
      if (b.distancia_realizada !== a.distancia_realizada) return b.distancia_realizada - a.distancia_realizada;
      if (b.percentual_concluido !== a.percentual_concluido) return b.percentual_concluido - a.percentual_concluido;
      return String(a.id_dgmb || '').localeCompare(String(b.id_dgmb || ''));
    });

    for (var p = 0; p < ranking.length; p++) {
      ranking[p].posicao = p + 1;
      ranking[p].posicao_ranking = p + 1;
    }

    return {
      ok: true,
      data: ranking,
      total: ranking.length
    };
  } catch (err) {
    return {
      ok: false,
      data: [],
      total: 0,
      msg: err && err.message ? err.message : 'Erro ao carregar ranking.'
    };
  }
}

function rankingMG_resolverAbaDesafio_(idDgmb) {
  var id = rankingMG_norm_(idDgmb);

  if (id) {
    try {
      var localizacao = localizarAbaDesafioUsuario_(id) || {};
      var abaUsuario = rankingMG_norm_(localizacao.abaDesafio);
      if (abaUsuario) return abaUsuario;
    } catch (e) {}
  }

  try {
    var ss = getSpreadsheet_();
    var lista = ss.getSheetByName(SHEETS.LISTA_DESAFIOS || 'ListaDesafios');
    if (lista) {
      var rows = lista.getDataRange().getDisplayValues();
      if (rows && rows.length > 1) {
        var map = buildHeaderMap_(rows[0]);
        var idxAba = getOptionalColumnIndex_(map, ['aba', 'aba desafio', 'abadesafio']);
        var idxStatus = getOptionalColumnIndex_(map, ['status', 'situacao', 'situação']);

        if (idxAba === -1) idxAba = 1;
        if (idxStatus === -1) idxStatus = 3;

        for (var i = 1; i < rows.length; i++) {
          var aba = rankingMG_norm_(rows[i][idxAba]);
          var status = normalizeText_(rows[i][idxStatus]).toLowerCase();

          if (!aba || status !== 'ativo') continue;
          if (ss.getSheetByName(aba)) return aba;
        }
      }
    }
  } catch (e) {}

  return SHEETS.DESAFIO;
}

function rankingMG_buildPessoasMap_(pessoas) {
  var map = {};

  for (var i = 0; i < pessoas.length; i++) {
    var row = pessoas[i];

    var idDgmb = rankingMG_norm_(rankingMG_firstFilled_(row, ['ID_DGMB', 'id_dgmb']));
    if (!idDgmb) continue;

    map[idDgmb] = {
      id_dgmb: idDgmb,
      nome: rankingMG_norm_(rankingMG_firstFilled_(row, ['nome', 'Nome'])),
      cidade_uf: rankingMG_norm_(rankingMG_firstFilled_(row, ['Cidade-UF', 'Cidade_UF', 'cidade_uf', 'cidade-uf']))
    };
  }

  return map;
}

function rankingMG_firstFilled_(obj, keys) {
  for (var i = 0; i < keys.length; i++) {
    if (Object.prototype.hasOwnProperty.call(obj, keys[i])) {
      var value = obj[keys[i]];
      if (value !== '' && value !== null && value !== undefined) {
        return value;
      }
    }
  }
  return '';
}

function rankingMG_norm_(value) {
  return String(value || '').trim();
}

function rankingMG_toNumber_(value) {
  if (typeof value === 'number') {
    return isNaN(value) ? 0 : value;
  }

  var s = String(value == null ? '' : value).trim();
  if (!s) return 0;

  s = s.replace(/\s+/g, '');

  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.indexOf(',') >= 0) {
    s = s.replace(',', '.');
  }

  var n = Number(s);
  return isNaN(n) ? 0 : n;
}

function rankingMG_round1_(n) {
  return Math.round((Number(n || 0) + Number.EPSILON) * 10) / 10;
}

function rankingMG_extrairGrupoBaseDesafio_(idItemEstoque) {
  var item = rankingMG_norm_(idItemEstoque);
  if (!item) return '';

  var semKm = item.replace(/_[0-9]+(?:[.,][0-9]+)?$/g, '');
  return rankingMG_norm_(semKm || item);
}
