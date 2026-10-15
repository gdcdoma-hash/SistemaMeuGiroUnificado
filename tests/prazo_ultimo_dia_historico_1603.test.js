'use strict';
const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname,'..','Meu Giro');
const modalSource = fs.readFileSync(path.join(root,'ModalidadeEsportivaService.gs'),'utf8');
const utilsSource = fs.readFileSync(path.join(root,'Utils.gs'),'utf8');
const uiSource = fs.readFileSync(path.join(root,'Script.html'),'utf8');

function getFunction(source,name) {
  const begin=source.indexOf('function '+name+'(');
  assert.ok(begin>=0,'missing function '+name);
  const end=source.indexOf('\nfunction ',begin+9);
  assert.ok(end>begin,'missing next function after '+name);
  return source.slice(begin,end);
}

const inRange=getFunction(utilsSource,'atividadeDentroPeriodoOficial_');
const uiDate=getFunction(uiSource,'parseIsoDateLocal');
const uiRemaining=getFunction(uiSource,'obterTempoRestanteDesafioPainel_');

function isISO(date) {
  if(!/^\d{4}-\d{2}-\d{2}$/.test(String(date||'')))return false;
  const d=new Date(date+'T12:00:00Z');
  return !Number.isNaN(d.getTime()) && d.toISOString().slice(0,10)===date;
}

function server(today) {
  return new Function('isDataIsoValida_','normalizarDataISO_','Utilities',
    inRange+'\n'+modalSource+
    '\nreturn {sum:meuGiroSomarKmPorVinculo_,validate:meuGiroValidarDataModalidadeInscricao_};'
  )(isISO, v=>String(v||''), {formatDate:()=>today});
}

function remaining(today, end) {
  const date = today+'T12:00:00-03:00';
  const Clock=class extends Date {
    constructor(...args){super(...(args.length?args:[date]));}
  };
  const get = new Function('Date',uiDate+'\n'+uiRemaining+
    '\nreturn obterTempoRestanteDesafioPainel_;')(Clock);
  return get(end);
}

const activity=(date,km,modality='CICLISMO')=>({
  data_atividade:date,km,modalidade_esportiva:modality
});
const historical=[
  activity('2026-08-31',1),                 // too early: exclude from 45 days
  activity('2026-09-30',10),activity('2026-09-30',2),
  activity('2026-10-01',3),activity('2026-10-01',3),
  activity('2026-09-29',44,'CAMINHADA_CORRIDA'),
  activity('2026-10-12',5.5,'CAMINHADA_CORRIDA')
];
const deadlines=[
  {periodo_inicio:'2026-09-01',periodo_fim:'2026-10-15',modalidade_esportiva:'CICLISMO'},
  {periodo_inicio:'2026-10-01',periodo_fim:'2026-10-31',modalidade_esportiva:'CICLISMO'},
  {periodo_inicio:'2026-10-01',periodo_fim:'2026-10-31',modalidade_esportiva:'CAMINHADA_CORRIDA'}
];
const totals=(api,activities)=>deadlines.map(d=>api.sum(activities,d));

test('15/10 é inclusivo nos dois desafios de ciclismo sem alterar corrida',()=>{
  const api=server('2026-10-15');
  assert.deepEqual(totals(api,historical),[18,6,5.5]);
  assert.deepEqual(totals(api,[...historical,activity('2026-10-15',1)]),[19,7,5.5]);
  assert.equal(api.validate('2026-10-15','CICLISMO').ok,true);
  assert.equal(remaining('2026-10-15','2026-10-15'),'Último dia');
});

test('14/10 bloqueia 15/10 e 16/10 e não antecipa atividade futura legada',()=>{
  const api=server('2026-10-14');
  assert.equal(api.validate('2026-10-15','CICLISMO').code,'DATA_FUTURA');
  assert.equal(api.validate('2026-10-16','CICLISMO').code,'DATA_FUTURA');
  assert.deepEqual(totals(api,[...historical,activity('2026-10-15',1),
    activity('2026-10-16',1)]),[18,6,5.5]);
  assert.equal(remaining('2026-10-14','2026-10-15'),'1 dia');
});

test('16/10 entra no histórico/outubro, mas não entra na meta de 45 dias',()=>{
  const api=server('2026-10-16');
  assert.equal(api.validate('2026-10-16','CICLISMO').ok,true);
  const withEnd=[...historical,activity('2026-10-15',1)];
  assert.deepEqual(totals(api,withEnd),[19,7,5.5]);
  assert.deepEqual(totals(api,[...withEnd,activity('2026-10-16',1)]),[19,8,5.5]);
  assert.equal(remaining('2026-10-16','2026-10-15'),'Encerrado');
});

test('encerramento preserva atividade e total final mesmo em novembro',()=>{
  const api=server('2026-11-01');
  assert.deepEqual(totals(api,[...historical,activity('2026-10-15',1),
    activity('2026-10-16',1),activity('2026-11-01',20)]),[19,8,5.5]);
});

test('data inválida continua rejeitada, sem corrigir silenciosamente',()=>{
  const api=server('2026-10-16');
  assert.equal(api.validate('2026-02-30','CICLISMO').code,'DATA_INVALIDA');
  assert.equal(api.validate('2026-13-16','CICLISMO').code,'DATA_INVALIDA');
});

test('somente a modalidade correspondente recebe os km do último dia',()=>{
  const api=server('2026-10-15');
  assert.deepEqual(totals(api,[...historical,activity('2026-10-15',1,'CAMINHADA_CORRIDA')]),
    [18,6,6.5]);
});

test('remover o registro temporário devolve os quatro resumos anteriores',()=>{
  const api=server('2026-10-15');
  const withTemp=[...historical,activity('2026-10-15',1)];
  assert.deepEqual(totals(api,withTemp),[19,7,5.5]);
  const withoutTemp=withTemp.filter((_,i)=>i!==withTemp.length-1);
  assert.deepEqual(totals(api,withoutTemp),[18,6,5.5]);
  assert.equal(withoutTemp.length,historical.length);
});
