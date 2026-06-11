function getRanking(idDgmb, idDesafio, idItemEstoque, idInscricao) {
  try {
    var idUsuario = rankingMG_norm_(idDgmb);
    var desafioSolicitado = rankingMG_norm_(idDesafio);
    var itemSolicitado = rankingMG_norm_(idItemEstoque);
    var inscricaoSolicitada = rankingMG_norm_(idInscricao);
    var diagnosticos = rankingMG_criarDiagnosticos_(inscricaoSolicitada);

    if (!idUsuario) {
      return { ok: false, data: [], total: 0, msg: 'ID do usuário não informado.', diagnosticos: diagnosticos };
    }

    var pessoas = getAllObjects_(SHEETS.PESSOAS);
    var resumo = getAllObjects_(SHEETS.MEU_GIRO_RESUMO) || [];

    if (!resumo.length) {
      return {
        ok: true,
        data: [],
        total: 0,
        msg: 'Nenhum atleta encontrado no ranking.',
        diagnosticos: diagnosticos
      };
    }

    var statusValidos = { ATIVO: true, CONCLUIDO: true };
    var referencia = null;

    if (inscricaoSolicitada) {
      for (var ri = 0; ri < resumo.length; ri++) {
        var rowInscricao = resumo[ri] || {};
        var idInscricaoRef = rankingMG_obterIdInscricao_(rowInscricao);
        var idDgmbInscricaoRef = rankingMG_norm_(rankingMG_firstFilled_(rowInscricao, ['ID_DGMB', 'id_dgmb']));
        if (idInscricaoRef === inscricaoSolicitada && idDgmbInscricaoRef === idUsuario) {
          referencia = rowInscricao;
          diagnosticos.referencia_por = 'ID_INSCRICAO';
          break;
        }
      }

      if (!referencia) {
        diagnosticos.id_inscricao_nao_encontrado = true;
        rankingMG_logDiagnostico_('ID_INSCRICAO recebido, mas não encontrado no MEU_GIRO_RESUMO.', {
          id_dgmb: idUsuario,
          id_inscricao: inscricaoSolicitada,
          id_desafio: desafioSolicitado,
          id_item_estoque: itemSolicitado
        });
      }
    }

    // Compatibilidade com chamadas antigas e com inscrições legadas sem ID_INSCRICAO.
    if (!referencia) {
      diagnosticos.referencia_por = 'LEGADO';
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
    }

    if (!referencia) {
      return {
        ok: true,
        data: [],
        total: 0,
        msg: 'Usuário sem desafio elegível para o ranking.',
        diagnosticos: diagnosticos
      };
    }

    var idInscricaoReferencia = rankingMG_obterIdInscricao_(referencia);
    var desafioPrincipal = rankingMG_norm_(rankingMG_firstFilled_(referencia, ['ID_DESAFIO', 'id_desafio']));
    var itemPrincipal = rankingMG_norm_(rankingMG_firstFilled_(referencia, ['id_item_estoque', 'id item estoque']));
    var grupoBasePrincipal = rankingMG_extrairGrupoBaseDesafio_(itemPrincipal);

    diagnosticos.id_inscricao_referencia = idInscricaoReferencia;
    diagnosticos.id_desafio_referencia = desafioPrincipal;
    diagnosticos.grupo_base_referencia = grupoBasePrincipal;

    if (!desafioPrincipal) {
      return {
        ok: true,
        data: [],
        total: 0,
        msg: 'Desafio-base não identificado para o ranking.',
        diagnosticos: diagnosticos
      };
    }

    var pessoasMap = rankingMG_buildPessoasMap_(pessoas);
    var ranking = [];
    var ocorrenciasPorDgmb = {};
    var metasMap = {};

    for (var i = 0; i < resumo.length; i++) {
      var row = resumo[i] || {};

      var idDgmbRanking = rankingMG_norm_(rankingMG_firstFilled_(row, ['ID_DGMB', 'id_dgmb']));
      if (!idDgmbRanking) continue;

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

      var pessoa = pessoasMap[idDgmbRanking] || {};
      var rowIdInscricao = rankingMG_obterIdInscricao_(row);
      ocorrenciasPorDgmb[idDgmbRanking] = (ocorrenciasPorDgmb[idDgmbRanking] || 0) + 1;
      metasMap[String(meta)] = meta;

      ranking.push({
        id_inscricao: rowIdInscricao,
        id_dgmb: idDgmbRanking,
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

    var posicaoUsuario = 0;
    for (var p = 0; p < ranking.length; p++) {
      ranking[p].posicao = p + 1;
      ranking[p].posicao_ranking = p + 1;

      if (!posicaoUsuario && idInscricaoReferencia && ranking[p].id_inscricao === idInscricaoReferencia) {
        posicaoUsuario = p + 1;
      }
    }

    if (!posicaoUsuario && !idInscricaoReferencia) {
      for (var pu = 0; pu < ranking.length; pu++) {
        if (ranking[pu].id_dgmb === idUsuario) {
          posicaoUsuario = pu + 1;
          break;
        }
      }
    }

    rankingMG_finalizarDiagnosticos_(diagnosticos, ocorrenciasPorDgmb, metasMap);

    return {
      ok: true,
      data: ranking,
      total: ranking.length,
      posicao_usuario: posicaoUsuario,
      id_inscricao_referencia: idInscricaoReferencia,
      diagnosticos: diagnosticos
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

function rankingMG_obterIdInscricao_(row) {
  return rankingMG_norm_(rankingMG_firstFilled_(row || {}, [
    'ID_INSCRICAO', 'id_inscricao', 'ID Inscricao', 'ID Inscrição', 'id inscricao', 'id inscrição'
  ]));
}

function rankingMG_criarDiagnosticos_(idInscricaoSolicitada) {
  return {
    id_inscricao_recebido: rankingMG_norm_(idInscricaoSolicitada),
    id_inscricao_nao_encontrado: false,
    referencia_por: '',
    id_inscricao_referencia: '',
    id_desafio_referencia: '',
    grupo_base_referencia: '',
    dgmb_duplicados: [],
    metas_distintas: [],
    possui_dgmb_duplicado: false,
    possui_multiplas_metas: false,
    metas_diferentes_no_ranking: false,
    grupo_competitivo_multiplas_metas: false
  };
}

function rankingMG_finalizarDiagnosticos_(diagnosticos, ocorrenciasPorDgmb, metasMap) {
  var duplicados = [];
  var metas = [];
  var id;

  for (id in ocorrenciasPorDgmb) {
    if (Object.prototype.hasOwnProperty.call(ocorrenciasPorDgmb, id) && ocorrenciasPorDgmb[id] > 1) {
      duplicados.push({ id_dgmb: id, quantidade: ocorrenciasPorDgmb[id] });
    }
  }

  for (var metaKey in metasMap) {
    if (Object.prototype.hasOwnProperty.call(metasMap, metaKey)) metas.push(metasMap[metaKey]);
  }

  duplicados.sort(function(a, b) { return String(a.id_dgmb).localeCompare(String(b.id_dgmb)); });
  metas.sort(function(a, b) { return a - b; });

  diagnosticos.dgmb_duplicados = duplicados;
  diagnosticos.metas_distintas = metas;
  diagnosticos.possui_dgmb_duplicado = duplicados.length > 0;
  diagnosticos.possui_multiplas_metas = metas.length > 1;
  diagnosticos.metas_diferentes_no_ranking = metas.length > 1;
  diagnosticos.grupo_competitivo_multiplas_metas = metas.length > 1;

  if (diagnosticos.possui_dgmb_duplicado) {
    rankingMG_logDiagnostico_('Mais de uma linha do mesmo ID_DGMB no ranking.', duplicados);
  }
  if (diagnosticos.possui_multiplas_metas) {
    rankingMG_logDiagnostico_('Grupo competitivo formado por mais de uma meta.', metas);
  }
}

function rankingMG_logDiagnostico_(mensagem, dados) {
  try {
    Logger.log('[Meu Giro][Ranking] ' + mensagem + ' ' + JSON.stringify(dados || {}));
  } catch (e) {}
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
