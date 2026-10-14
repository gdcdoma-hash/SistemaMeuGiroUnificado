'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const source = fs.readFileSync(path.join(__dirname, '..', 'Meu Giro', 'ModalidadeEsportivaService.gs'), 'utf8');

// Usa a função real do app, sem planilha nem gravação. O relógio é fixado
// apenas nesta simulação para que registros futuros não adiantem o progresso.
function motor(hoje) {
  const periodo = (data, inicio, fim) => data >= inicio && data <= fim;
  const isData = value => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
    const d = new Date(value + 'T12:00:00Z');
    return !Number.isNaN(d.getTime()) && d.toISOString().slice(0, 10) === value;
  };
  return new Function(
    'atividadeDentroPeriodoOficial_', 'normalizarDataISO_',
    'isDataIsoValida_', 'Utilities',
    source + '\nreturn {somar:meuGiroSomarKmPorVinculo_, validar:meuGiroValidarDataModalidadeInscricao_};'
  )(periodo, value => String(value || ''), isData, {formatDate: () => hoje});
}

const vinculo = (inicio, fim, modalidade) => ({
  periodo_inicio: inicio, periodo_fim: fim, modalidade_esportiva: modalidade
});

test('1603 simulado: 10 km de setembro entram no 45 dias, não no Outubro Rosa nem na Corrida', () => {
  const api = motor('2026-10-14');
  const registros = [
    {data_atividade:'2026-08-31', km:3, modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-09-30', km:10, modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-10-12', km:6, modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-10-12', km:5.5, modalidade_esportiva:'CAMINHADA_CORRIDA'},
    {data_atividade:'2026-10-16', km:7, modalidade_esportiva:'CICLISMO'}
  ];

  assert.equal(api.somar(registros, vinculo('2026-09-01', '2026-10-15', 'CICLISMO')), 16);
  assert.equal(api.somar(registros, vinculo('2026-10-01', '2026-10-31', 'CICLISMO')), 6);
  assert.equal(api.somar(registros, vinculo('2026-10-01', '2026-10-31', 'CAMINHADA_CORRIDA')), 5.5);
  // A mesma atividade válida (6 km) soma em outro desafio de Ciclismo com
  // janela sobreposta, mas não gera outra linha no histórico.
  assert.equal(registros.filter(r => r.data_atividade === '2026-10-12' &&
    r.modalidade_esportiva === 'CICLISMO').length, 1);
});

test('registro retroativo é permitido no histórico; futuro é bloqueado', () => {
  const api = motor('2026-10-14');
  const retroativo = api.validar('2026-09-30', 'CICLISMO', []);
  assert.equal(retroativo.ok, true);
  const futuro = api.validar('2026-10-16', 'CICLISMO', []);
  assert.equal(futuro.ok, false);
  assert.equal(futuro.code, 'DATA_FUTURA');
});
