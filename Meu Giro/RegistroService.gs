function registrarAtividade(idDgmb, dataAtividade, km, force) {
  var perfTotalInicio = meuGiroPerfNow_();
  var perfOperacaoAnterior = MEU_GIRO_PERF_OPERACAO_ATUAL_;
  MEU_GIRO_PERF_OPERACAO_ATUAL_ = 'registrarAtividade';
  var lock = LockService.getScriptLock();
  try {
    var perfEtapaInicio = meuGiroPerfNow_();
    lock.waitLock(30000);
    meuGiroPerfLog_('registrar-atividade', 'LockService', perfEtapaInicio);

    idDgmb = String(idDgmb || '').trim();
    dataAtividade = normalizarDataISO_(dataAtividade);
    km = parseKmInputSeguro_(km);

    if (!idDgmb) {
      return { ok:false, code:'ID_OBRIGATORIO', msg:'ID do atleta é obrigatório.' };
    }

    if (!dataAtividade) {
      return { ok:false, code:'DATA_OBRIGATORIA', msg:'Informe o dia da atividade.' };
    }

    if (!km || km <= 0) {
      return { ok:false, code:'KM_INVALIDO', msg:'Informe um valor de KM maior que zero.' };
    }

    perfEtapaInicio = meuGiroPerfNow_();
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
      .getSheetByName(SHEETS.REGISTRO_KM);

    var dados = sheet.getDataRange().getValues();
    meuGiroPerfLog_('registrar-atividade', 'leitura_REGISTRO_KM', perfEtapaInicio, {
      quantidade_linhas_registro_km: dados && dados.length ? dados.length - 1 : 0
    });
    var cols = getRegistroKmColumnIndexes_(dados);

    cols = ensureRegistroKmActivityIdColumn_(sheet, dados, cols);
    if (!dados[0]) dados[0] = [];
    dados[0][cols.idxActivityId] = dados[0][cols.idxActivityId] || 'activity_id';
    var activityId = gerarActivityId_();

    var perfBuscaDuplicidadeInicio = meuGiroPerfNow_();
    var perfNormalizacaoDuplicidadeInicio = meuGiroPerfNow_();
    var quantidadeLinhasVerificadas = 0;
    var totalComparacoes = 0;
    var duracaoComparacoesMs = 0;
    var duracaoNormalizacaoMs = 0;
    var duplicidadeEncontrada = false;
    duracaoNormalizacaoMs += meuGiroPerfNow_() - perfNormalizacaoDuplicidadeInicio;

    var perfLoopDuplicidadeInicio = meuGiroPerfNow_();
    for (var i = 1; i < dados.length; i++) {
      quantidadeLinhasVerificadas++;
      perfNormalizacaoDuplicidadeInicio = meuGiroPerfNow_();
      var rowId = String(dados[i][cols.idxId] || '').trim();
      var rowData = normalizarDataISO_(dados[i][cols.idxData]);
      var rowKm = normalizarKmEdicao_(dados[i][cols.idxKm]);
      duracaoNormalizacaoMs += meuGiroPerfNow_() - perfNormalizacaoDuplicidadeInicio;

      var perfComparacaoInicio = meuGiroPerfNow_();
      totalComparacoes++;
      var mesmoId = rowId === idDgmb;
      var mesmaData = false;
      var mesmoKm = false;
      if (mesmoId) {
        totalComparacoes++;
        mesmaData = rowData === dataAtividade;
        if (mesmaData) {
          totalComparacoes++;
          mesmoKm = kmsIguaisEdicao_(rowKm, km);
        }
      }
      duracaoComparacoesMs += meuGiroPerfNow_() - perfComparacaoInicio;

      if (mesmoId && mesmaData && mesmoKm) {
        duplicidadeEncontrada = true;
        meuGiroPerfLog_('registrar-atividade', 'duplicidade_encontrada', perfBuscaDuplicidadeInicio, {
          linha_encontrada: i + 1
        });
        if (!force) {
          meuGiroPerfLog_('registrar-atividade', 'normalizacao_dados_duplicidade',
            meuGiroPerfNow_() - duracaoNormalizacaoMs, {
              campos_preparados: ['data_atividade', 'km', 'id_dgmb', 'activity_id'],
              campos_comparados: ['id_dgmb', 'data_atividade', 'km'],
              data_atividade: dataAtividade,
              km: km,
              id_dgmb: idDgmb,
              activity_id: activityId
            });
          meuGiroPerfLog_('registrar-atividade', 'loop_busca_duplicidade', perfLoopDuplicidadeInicio, {
            quantidade_linhas_verificadas: quantidadeLinhasVerificadas
          });
          meuGiroPerfLog_('registrar-atividade', 'comparacoes_duplicidade',
            meuGiroPerfNow_() - duracaoComparacoesMs, {
              total_comparacoes: totalComparacoes
            });
          meuGiroPerfLog_('registrar-atividade', 'busca_duplicidade_total', perfBuscaDuplicidadeInicio, {
            quantidade_linhas_verificadas: quantidadeLinhasVerificadas,
            total_comparacoes: totalComparacoes,
            duplicidade_encontrada: duplicidadeEncontrada
          });
          return {
            ok:false,
            code:'DUPLICIDADE',
            msg:'Já existe atividade com mesmo ID, data e KM informado.'
          };
        }
      }
    }
    meuGiroPerfLog_('registrar-atividade', 'normalizacao_dados_duplicidade',
      meuGiroPerfNow_() - duracaoNormalizacaoMs, {
        campos_preparados: ['data_atividade', 'km', 'id_dgmb', 'activity_id'],
        campos_comparados: ['id_dgmb', 'data_atividade', 'km'],
        data_atividade: dataAtividade,
        km: km,
        id_dgmb: idDgmb,
        activity_id: activityId
      });
    meuGiroPerfLog_('registrar-atividade', 'loop_busca_duplicidade', perfLoopDuplicidadeInicio, {
      quantidade_linhas_verificadas: quantidadeLinhasVerificadas
    });
    meuGiroPerfLog_('registrar-atividade', 'comparacoes_duplicidade',
      meuGiroPerfNow_() - duracaoComparacoesMs, {
        total_comparacoes: totalComparacoes
      });
    meuGiroPerfLog_('registrar-atividade', 'busca_duplicidade_total', perfBuscaDuplicidadeInicio, {
      quantidade_linhas_verificadas: quantidadeLinhasVerificadas,
      total_comparacoes: totalComparacoes,
      duplicidade_encontrada: duplicidadeEncontrada
    });

    var rowLength = Math.max(sheet.getLastColumn(), maiorIndiceRegistroKm_(cols) + 1);
    var row = preencherLinhaRegistroKmBruto_(rowLength, cols, {
      timestamp: new Date(),
      id_dgmb: idDgmb,
      data_atividade: dataAtividade,
      km: km,
      origem_registro: 'MANUAL',
      observacao: 'Lançamento manual Meu Giro',
      status_validacao: 'PENDENTE',
      activity_id: activityId
    });

    var linhaInserida = sheet.getLastRow() + 1;
    perfEtapaInicio = meuGiroPerfNow_();
    sheet.getRange(linhaInserida, 1, 1, rowLength).setValues([row]);
    dados.push(row.slice());
    var opcoesRegistroKm = criarOpcoesRegistroKmReaproveitado_(idDgmb, dados, cols);
    meuGiroPerfLog_('registrar-atividade', 'escrita_atividade_REGISTRO_KM', perfEtapaInicio, {
      quantidade_linhas_escritas: 1
    });

    try {
      perfEtapaInicio = meuGiroPerfNow_();
      atualizarDistanciaRealizada_(idDgmb, opcoesRegistroKm);
      meuGiroPerfLog_('registrar-atividade', 'atualizarDistanciaRealizada_', perfEtapaInicio);
      perfEtapaInicio = meuGiroPerfNow_();
      atualizarMeuGiroResumo_(idDgmb, opcoesRegistroKm);
      meuGiroPerfLog_('registrar-atividade', 'atualizarMeuGiroResumo_', perfEtapaInicio);
    } catch (syncErr) {
      sheet.deleteRow(linhaInserida);
      throw syncErr;
    }

    return {
      ok:true,
      activity_id: activityId,
      registros_criados: 1,
      msg:'Atividade registrada com sucesso.'
    };

  } catch(err) {
    Logger.log('registrarAtividade erro: ' + (err && err.stack ? err.stack : err));

    return {
      ok:false,
      code:'REGISTRAR_ATIVIDADE_EXCEPTION',
      msg:'Erro interno ao registrar atividade na aba REGISTRO_KM.'
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
    meuGiroPerfLog_('registrar-atividade', 'registrarAtividade_total', perfTotalInicio);
    MEU_GIRO_PERF_OPERACAO_ATUAL_ = perfOperacaoAnterior;
  }
}

function gerarActivityId_() {
  return Utilities.getUuid();
}

function ensureRegistroKmActivityIdColumn_(sheet, dados, cols) {
  if (cols && cols.idxActivityId > -1) {
    return cols;
  }

  var headerLength = (dados && dados[0] && dados[0].length) ? dados[0].length : 0;
  var newIndex = headerLength;

  sheet.getRange(1, newIndex + 1).setValue('activity_id');

  cols.idxActivityId = newIndex;
  return cols;
}

function maiorIndiceRegistroKm_(cols) {
  var maior = -1;
  Object.keys(cols || {}).forEach(function(chave) {
    if (typeof cols[chave] === 'number' && cols[chave] > maior) {
      maior = cols[chave];
    }
  });
  return maior;
}

function preencherLinhaRegistroKmBruto_(rowLength, cols, registro) {
  var row = [];
  for (var i = 0; i < rowLength; i++) row[i] = '';

  // REGISTRO_KM representa a pedalada real, não o vínculo com inscrição/desafio.
  // A linha começa vazia para preservar, sem preencher, eventuais colunas legadas.
  row[cols.idxTimestamp] = registro.timestamp;
  row[cols.idxId] = registro.id_dgmb;
  row[cols.idxData] = registro.data_atividade;
  row[cols.idxKm] = registro.km;
  if (cols.idxOrigemRegistro > -1) row[cols.idxOrigemRegistro] = registro.origem_registro || '';
  if (cols.idxObservacao > -1) row[cols.idxObservacao] = registro.observacao || '';
  if (cols.idxStatusValidacao > -1) row[cols.idxStatusValidacao] = registro.status_validacao || '';
  row[cols.idxActivityId] = registro.activity_id;

  return row;
}

function criarOpcoesRegistroKmReaproveitado_(idDgmb, valores, layoutRegistroKm) {
  return {
    idDgmb: String(idDgmb || '').trim(),
    registrosKmValores: valores,
    registrosKmObjetos: converterValoresRegistroKmEmObjetos_(valores),
    layoutRegistroKm: layoutRegistroKm
  };
}

function atualizarDistanciaRealizada_(idDgmb, opcoes){
  var perfTotalInicio = meuGiroPerfNow_();
  var perfEtapaInicio = meuGiroPerfNow_();
  var registros = obterRegistrosKmObjetosReaproveitados_(idDgmb, opcoes);
  meuGiroPerfLog_('atualizar-distancia-realizada', registros.reaproveitados ? 'leitura_REGISTRO_KM_reaproveitada' : 'leitura_REGISTRO_KM', perfEtapaInicio, {
    quantidade_linhas_registro_km: registros.valores.length
  });
  var registrosValores = registros.valores;
  var total = 0;
  var activityIdsSomados = {};

  registrosValores.forEach(function(r){
    if(String(r.ID_DGMB).trim() !== String(idDgmb).trim()) return;

    var activityId = obterActivityIdRegistroKm_(r);

    if (activityId) {
      if (activityIdsSomados[activityId]) return;
      activityIdsSomados[activityId] = true;
    }

    total += Number(r.KM || 0);
  });

  perfEtapaInicio = meuGiroPerfNow_();
  var cacheDesafios = obterDgmbDesafiosCacheExecucao_('atualizarDistanciaRealizada_');
  var abaDesafio = cacheDesafios.aba;
  var sheet = cacheDesafios.sheet;
  var dados = cacheDesafios.values;
  meuGiroPerfLog_('atualizar-distancia-realizada', 'leitura_dgmbDesafios_cache', perfEtapaInicio, {
    quantidade_linhas_dgmbDesafios: dados && dados.length ? dados.length - 1 : 0,
    usou_cache_dgmbDesafios: cacheDesafios.usouCache
  });

  var inscricao = obterDadosInscricaoUsuario_(idDgmb, {
    abaDesafio: abaDesafio,
    values: dados,
    cache: cacheDesafios
  });
  if (!inscricao || !inscricao.aba_desafio) return;
  if (!dados || dados.length < 2) return;

  var map = buildHeaderMap_(dados[0]);
  var idxId = getRequiredColumnIndex_(map, ['id_dgmb'], abaDesafio);
  var idxRealizado = getRequiredColumnIndex_(map, ['distancia_realizada', 'distancia realizada'], abaDesafio);

  for(var i=1;i<dados.length;i++){
    if(String(dados[i][idxId]).trim() === String(idDgmb).trim()){
      perfEtapaInicio = meuGiroPerfNow_();
      sheet.getRange(i + 1, idxRealizado + 1).setValue(total);
      dados[i][idxRealizado] = total;
      meuGiroPerfLog_('atualizar-distancia-realizada', 'escrita_dgmbDesafios_distancia_realizada', perfEtapaInicio, {
        linha_atualizada: i + 1,
        cache_memoria_atualizado: true
      });
      break;
    }
  }
  meuGiroPerfLog_('atualizar-distancia-realizada', 'atualizarDistanciaRealizada_total', perfTotalInicio);
}

function editarAtividade(payload) {
  var perfTotalInicio = meuGiroPerfNow_();
  var perfOperacaoAnterior = MEU_GIRO_PERF_OPERACAO_ATUAL_;
  MEU_GIRO_PERF_OPERACAO_ATUAL_ = 'editarAtividade';
  var lock = LockService.getScriptLock();
  try {
    var perfEtapaInicio = meuGiroPerfNow_();
    lock.waitLock(30000);
    meuGiroPerfLog_('editar-atividade', 'LockService', perfEtapaInicio);
    payload = payload || {};

    var idDgmb = String(payload.id_dgmb || '').trim();
    var activityId = String(payload.activity_id || '').trim();
    var chaveEdicao = String(payload.chave_edicao || '').trim();
    var novaDataAtividade = normalizarDataISO_(payload.data_atividade);
    var novoKm = parseKmInputSeguro_(payload.km);

    if (!idDgmb) {
      return { ok: false, code: 'ID_OBRIGATORIO', msg: 'ID do atleta é obrigatório.' };
    }
    if (!activityId && !chaveEdicao) {
      return { ok: false, code: 'IDENTIFICADOR_ATIVIDADE_OBRIGATORIO', msg: 'activity_id ou chave_edicao é obrigatório para edição.' };
    }
    if (!novaDataAtividade) {
      return { ok: false, code: 'DATA_OBRIGATORIA', msg: 'Informe o dia da atividade.' };
    }
    if (!novoKm || novoKm <= 0) {
      return { ok: false, code: 'KM_INVALIDO', msg: 'Informe um valor de KM maior que zero.' };
    }

    perfEtapaInicio = meuGiroPerfNow_();
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.REGISTRO_KM);
    var dados = sheet.getDataRange().getValues();
    meuGiroPerfLog_('editar-atividade', 'leitura_REGISTRO_KM', perfEtapaInicio, {
      quantidade_linhas_registro_km: dados && dados.length ? dados.length - 1 : 0
    });
    var cols = getRegistroKmColumnIndexes_(dados);
    perfEtapaInicio = meuGiroPerfNow_();
    var linhasEncontradas = localizarLinhasAtividade_(dados, cols, idDgmb, activityId, chaveEdicao);
    meuGiroPerfLog_('editar-atividade', 'localizar_linhas_atividade', perfEtapaInicio, {
      quantidade_linhas_verificadas: Math.max(dados.length - 1, 0),
      quantidade_linhas_encontradas: linhasEncontradas.length
    });

    if (!linhasEncontradas.length) {
      return { ok: false, code: 'ATIVIDADE_NAO_ENCONTRADA', msg: 'Atividade não encontrada para edição com a chave e ID informados.' };
    }

    var linhasDoLancamento = {};
    linhasEncontradas.forEach(function(linha) { linhasDoLancamento[linha] = true; });

    var perfBuscaDuplicidadeInicio = meuGiroPerfNow_();
    var perfNormalizacaoDuplicidadeInicio = meuGiroPerfNow_();
    var quantidadeLinhasVerificadas = 0;
    var totalComparacoes = 0;
    var duracaoComparacoesMs = 0;
    var duracaoNormalizacaoMs = 0;
    var duplicidadeEncontrada = false;
    duracaoNormalizacaoMs += meuGiroPerfNow_() - perfNormalizacaoDuplicidadeInicio;

    var perfLoopDuplicidadeInicio = meuGiroPerfNow_();
    for (var j = 1; j < dados.length; j++) {
      quantidadeLinhasVerificadas++;
      perfNormalizacaoDuplicidadeInicio = meuGiroPerfNow_();
      var checkId = String(dados[j][cols.idxId] || '').trim();
      var checkData = normalizarDataEdicao_(dados[j][cols.idxData]);
      var checkKm = normalizarKmEdicao_(dados[j][cols.idxKm]);
      var linhaAtual = j + 1;
      duracaoNormalizacaoMs += meuGiroPerfNow_() - perfNormalizacaoDuplicidadeInicio;

      var perfComparacaoInicio = meuGiroPerfNow_();
      totalComparacoes++;
      var foraDoLancamento = !linhasDoLancamento[linhaAtual];
      var mesmoId = false;
      var mesmaData = false;
      var mesmoKm = false;
      if (foraDoLancamento) {
        totalComparacoes++;
        mesmoId = checkId === idDgmb;
        if (mesmoId) {
          totalComparacoes++;
          mesmaData = checkData === novaDataAtividade;
          if (mesmaData) {
            totalComparacoes++;
            mesmoKm = kmsIguaisEdicao_(checkKm, novoKm);
          }
        }
      }
      duracaoComparacoesMs += meuGiroPerfNow_() - perfComparacaoInicio;

      if (foraDoLancamento && mesmoId && mesmaData && mesmoKm) {
        duplicidadeEncontrada = true;
        meuGiroPerfLog_('editar-atividade', 'duplicidade_encontrada', perfBuscaDuplicidadeInicio, {
          linha_encontrada: linhaAtual
        });
        meuGiroPerfLog_('editar-atividade', 'normalizacao_dados_duplicidade',
          meuGiroPerfNow_() - duracaoNormalizacaoMs, {
            campos_preparados: ['data_atividade', 'km', 'id_dgmb', 'activity_id', 'chave_edicao'],
            campos_comparados: ['linha_do_lancamento', 'id_dgmb', 'data_atividade', 'km'],
            data_atividade: novaDataAtividade,
            km: novoKm,
            id_dgmb: idDgmb,
            activity_id: activityId,
            chave_edicao: chaveEdicao
          });
        meuGiroPerfLog_('editar-atividade', 'loop_busca_duplicidade', perfLoopDuplicidadeInicio, {
          quantidade_linhas_verificadas: quantidadeLinhasVerificadas
        });
        meuGiroPerfLog_('editar-atividade', 'comparacoes_duplicidade',
          meuGiroPerfNow_() - duracaoComparacoesMs, {
            total_comparacoes: totalComparacoes
          });
        meuGiroPerfLog_('editar-atividade', 'busca_duplicidade_total', perfBuscaDuplicidadeInicio, {
          quantidade_linhas_verificadas: quantidadeLinhasVerificadas,
          total_comparacoes: totalComparacoes,
          duplicidade_encontrada: duplicidadeEncontrada
        });
        return { ok: false, code: 'DUPLICIDADE_EDICAO', msg: 'Já existe uma atividade com esta mesma data e KM.' };
      }
    }
    meuGiroPerfLog_('editar-atividade', 'normalizacao_dados_duplicidade',
      meuGiroPerfNow_() - duracaoNormalizacaoMs, {
        campos_preparados: ['data_atividade', 'km', 'id_dgmb', 'activity_id', 'chave_edicao'],
        campos_comparados: ['linha_do_lancamento', 'id_dgmb', 'data_atividade', 'km'],
        data_atividade: novaDataAtividade,
        km: novoKm,
        id_dgmb: idDgmb,
        activity_id: activityId,
        chave_edicao: chaveEdicao
      });
    meuGiroPerfLog_('editar-atividade', 'loop_busca_duplicidade', perfLoopDuplicidadeInicio, {
      quantidade_linhas_verificadas: quantidadeLinhasVerificadas
    });
    meuGiroPerfLog_('editar-atividade', 'comparacoes_duplicidade',
      meuGiroPerfNow_() - duracaoComparacoesMs, {
        total_comparacoes: totalComparacoes
      });
    meuGiroPerfLog_('editar-atividade', 'busca_duplicidade_total', perfBuscaDuplicidadeInicio, {
      quantidade_linhas_verificadas: quantidadeLinhasVerificadas,
      quantidade_linhas_do_lancamento: linhasEncontradas.length,
      total_comparacoes: totalComparacoes,
      duplicidade_encontrada: duplicidadeEncontrada
    });

    var valoresOriginais = linhasEncontradas.map(function(linha) {
      return [dados[linha - 1][cols.idxData], dados[linha - 1][cols.idxKm]];
    });

    perfEtapaInicio = meuGiroPerfNow_();
    linhasEncontradas.forEach(function(linha) {
      sheet.getRange(linha, cols.idxData + 1).setValue(novaDataAtividade);
      sheet.getRange(linha, cols.idxKm + 1).setValue(novoKm);
      dados[linha - 1][cols.idxData] = novaDataAtividade;
      dados[linha - 1][cols.idxKm] = novoKm;
    });
    var opcoesRegistroKm = criarOpcoesRegistroKmReaproveitado_(idDgmb, dados, cols);
    meuGiroPerfLog_('editar-atividade', 'edicao_atividade_REGISTRO_KM', perfEtapaInicio, {
      quantidade_linhas_editadas: linhasEncontradas.length,
      quantidade_escritas_celula: linhasEncontradas.length * 2
    });

    try {
      perfEtapaInicio = meuGiroPerfNow_();
      atualizarDistanciaRealizada_(idDgmb, opcoesRegistroKm);
      meuGiroPerfLog_('editar-atividade', 'atualizarDistanciaRealizada_', perfEtapaInicio);
      perfEtapaInicio = meuGiroPerfNow_();
      atualizarMeuGiroResumo_(idDgmb, opcoesRegistroKm);
      meuGiroPerfLog_('editar-atividade', 'atualizarMeuGiroResumo_', perfEtapaInicio);
    } catch (syncErr) {
      linhasEncontradas.forEach(function(linha, index) {
        sheet.getRange(linha, cols.idxData + 1).setValue(valoresOriginais[index][0]);
        sheet.getRange(linha, cols.idxKm + 1).setValue(valoresOriginais[index][1]);
      });
      throw syncErr;
    }

    return { ok: true, registros_atualizados: linhasEncontradas.length, msg: 'Atividade atualizada com sucesso.' };
  } catch (err) {
    Logger.log('editarAtividade erro: ' + (err && err.stack ? err.stack : err));
    return { ok: false, code: 'EDITAR_ATIVIDADE_EXCEPTION', msg: 'Erro interno ao editar atividade na aba REGISTRO_KM.' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
    meuGiroPerfLog_('editar-atividade', 'editarAtividade_total', perfTotalInicio);
    MEU_GIRO_PERF_OPERACAO_ATUAL_ = perfOperacaoAnterior;
  }
}

function parseKmInputSeguro_(value) {
  var text = String(value === null || value === undefined ? '' : value).trim();
  if (!text) return NaN;

  text = text.replace(/\s+/g, '');
  if (!/^\d+(?:[.,]\d+)?$/.test(text)) return NaN;

  var parsed = parseLocalizedNumber_(text);
  return isFinite(parsed) ? parsed : NaN;
}

function normalizarDataEdicao_(valor) {
  if (valor === null || valor === undefined || valor === '') return '';

  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  var s = String(valor).trim();
  if (!s) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    return s;
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    return s.slice(6, 10) + '-' + s.slice(3, 5) + '-' + s.slice(0, 2);
  }

  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return s;
}

function normalizarKmEdicao_(valor) {
  if (valor === null || valor === undefined || valor === '') return 0;

  var s = String(valor).trim().replace(/\s/g, '').replace(',', '.');
  var n = Number(s);

  if (isNaN(n)) return 0;

  return Math.round(n * 1000) / 1000;
}

function kmsIguaisEdicao_(a, b) {
  return Math.abs(Number(a || 0) - Number(b || 0)) < 0.0001;
}

function excluirAtividade(payload) {
  var perfTotalInicio = meuGiroPerfNow_();
  var perfOperacaoAnterior = MEU_GIRO_PERF_OPERACAO_ATUAL_;
  MEU_GIRO_PERF_OPERACAO_ATUAL_ = 'excluirAtividade';
  var lock = LockService.getScriptLock();
  try {
    var perfEtapaInicio = meuGiroPerfNow_();
    lock.waitLock(30000);
    meuGiroPerfLog_('excluir-atividade', 'LockService', perfEtapaInicio);
    payload = payload || {};

    var idDgmb = String(payload.id_dgmb || '').trim();
    var activityId = String(payload.activity_id || '').trim();
    var chaveEdicao = String(payload.chave_edicao || '').trim();

    if (!idDgmb) {
      return { ok: false, code: 'ID_OBRIGATORIO', msg: 'ID do atleta é obrigatório.' };
    }
    if (!activityId && !chaveEdicao) {
      return { ok: false, code: 'IDENTIFICADOR_ATIVIDADE_OBRIGATORIO', msg: 'activity_id ou chave_edicao é obrigatório para exclusão.' };
    }

    perfEtapaInicio = meuGiroPerfNow_();
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.REGISTRO_KM);
    var dados = sheet.getDataRange().getValues();
    meuGiroPerfLog_('excluir-atividade', 'leitura_REGISTRO_KM', perfEtapaInicio, {
      quantidade_linhas_registro_km: dados && dados.length ? dados.length - 1 : 0
    });
    var cols = getRegistroKmColumnIndexes_(dados);
    perfEtapaInicio = meuGiroPerfNow_();
    var linhasEncontradas = localizarLinhasAtividade_(dados, cols, idDgmb, activityId, chaveEdicao);
    meuGiroPerfLog_('excluir-atividade', 'busca_atividade_para_exclusao', perfEtapaInicio, {
      quantidade_linhas_verificadas: Math.max(dados.length - 1, 0),
      quantidade_linhas_encontradas: linhasEncontradas.length
    });

    if (!linhasEncontradas.length) {
      return { ok: false, code: 'ATIVIDADE_NAO_ENCONTRADA', msg: 'Atividade não encontrada para exclusão com a chave e ID informados.' };
    }

    var linhasOriginais = linhasEncontradas.map(function(linha) {
      return { numero: linha, valores: dados[linha - 1].slice() };
    });

    perfEtapaInicio = meuGiroPerfNow_();
    linhasEncontradas.slice().sort(function(a, b) { return b - a; }).forEach(function(linha) {
      sheet.deleteRow(linha);
      dados.splice(linha - 1, 1);
    });
    var opcoesRegistroKm = criarOpcoesRegistroKmReaproveitado_(idDgmb, dados, cols);
    meuGiroPerfLog_('excluir-atividade', 'exclusao_atividade_REGISTRO_KM', perfEtapaInicio, {
      quantidade_linhas_excluidas: linhasEncontradas.length
    });

    try {
      perfEtapaInicio = meuGiroPerfNow_();
      atualizarDistanciaRealizada_(idDgmb, opcoesRegistroKm);
      meuGiroPerfLog_('excluir-atividade', 'atualizarDistanciaRealizada_', perfEtapaInicio);
      perfEtapaInicio = meuGiroPerfNow_();
      atualizarMeuGiroResumo_(idDgmb, opcoesRegistroKm);
      meuGiroPerfLog_('excluir-atividade', 'atualizarMeuGiroResumo_', perfEtapaInicio);
    } catch (syncErr) {
      linhasOriginais.forEach(function(item) {
        sheet.insertRowBefore(item.numero);
        sheet.getRange(item.numero, 1, 1, item.valores.length).setValues([item.valores]);
      });
      throw syncErr;
    }

    return { ok: true, registros_excluidos: linhasEncontradas.length, msg: 'Atividade excluída com sucesso.' };
  } catch (err) {
    Logger.log('excluirAtividade erro: ' + (err && err.stack ? err.stack : err));
    return { ok: false, code: 'EXCLUSAO_ATIVIDADE_ERROR', msg: 'Erro interno ao excluir atividade na aba REGISTRO_KM.' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
    meuGiroPerfLog_('excluir-atividade', 'excluirAtividade_total', perfTotalInicio);
    MEU_GIRO_PERF_OPERACAO_ATUAL_ = perfOperacaoAnterior;
  }
}

function getRegistroKmColumnIndexes_(dados) {
  var fallback = {
    idxTimestamp: 0,
    idxId: 1,
    idxData: 2,
    idxKm: 3,
    idxInscricao: -1,
    idxDesafio: -1,
    idxItemEstoque: -1,
    idxPeriodoDesafio: -1,
    idxOrigemRegistro: -1,
    idxObservacao: -1,
    idxStatusValidacao: -1,
    idxActivityId: -1
  };

  if (!dados || !dados.length || !dados[0] || !dados[0].length) return fallback;

  var map = buildHeaderMap_(dados[0]);
  function indice(aliases, fallbackIndex) {
    var idx = getOptionalColumnIndex_(map, aliases);
    return idx > -1 ? idx : fallbackIndex;
  }

  return {
    idxTimestamp: indice(['timestamp', 'data_hora', 'data hora', 'criado_em', 'criado em'], fallback.idxTimestamp),
    idxId: indice(['id_dgmb'], fallback.idxId),
    idxInscricao: indice(['id_inscricao', 'id inscrição', 'id inscricao'], -1),
    idxDesafio: indice(['id_desafio', 'id desafio'], -1),
    idxItemEstoque: indice(['id_item_estoque', 'id item estoque'], -1),
    idxPeriodoDesafio: indice(['periodo_desafio', 'periodo desafio', 'período_desafio', 'período desafio'], -1),
    idxData: indice(['data_atividade', 'data atividade', 'data'], fallback.idxData),
    idxKm: indice(['km', 'distancia_km', 'distancia km'], fallback.idxKm),
    idxOrigemRegistro: indice(['origem_registro', 'origem registro'], -1),
    idxObservacao: indice(['observacao', 'observação'], -1),
    idxStatusValidacao: indice(['status_validacao', 'status validação', 'status validacao'], -1),
    idxActivityId: indice(['activity_id', 'activity id', 'id_atividade', 'id atividade'], -1)
  };
}

function localizarLinhasAtividade_(dados, cols, idDgmb, activityId, chaveEdicao) {
  var idNormalizado = String(idDgmb || '').trim();
  var activityIdNormalizado = String(activityId || '').trim();
  var chaveNormalizada = String(chaveEdicao || '').trim();
  var linhas = [];

  if (!dados || dados.length < 2) return linhas;

  if (activityIdNormalizado && cols.idxActivityId > -1) {
    for (var i = 1; i < dados.length; i++) {
      var rowIdByActivity = String(dados[i][cols.idxId] || '').trim();
      var rowActivityId = String(dados[i][cols.idxActivityId] || '').trim();
      if (rowIdByActivity === idNormalizado && rowActivityId === activityIdNormalizado) {
        linhas.push(i + 1);
      }
    }
    if (linhas.length) return linhas;
  }

  if (!chaveNormalizada) return linhas;

  for (var j = 1; j < dados.length; j++) {
    var rowTimestamp = normalizarTimestampEdicao_(dados[j][cols.idxTimestamp]);
    var rowId = String(dados[j][cols.idxId] || '').trim();
    if (rowTimestamp === chaveNormalizada && rowId === idNormalizado) {
      linhas.push(j + 1);
    }
  }

  return linhas;
}

function localizarLinhaAtividade_(dados, cols, idDgmb, activityId, chaveEdicao) {
  var linhas = localizarLinhasAtividade_(dados, cols, idDgmb, activityId, chaveEdicao);
  return linhas.length ? linhas[0] : -1;
}
