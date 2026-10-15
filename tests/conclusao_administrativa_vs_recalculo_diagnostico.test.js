'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root=path.resolve(__dirname,'..','Meu Giro');
const modal=fs.readFileSync(path.join(root,'ModalidadeEsportivaService.gs'),'utf8');
const utils=fs.readFileSync(path.join(root,'Utils.gs'),'utf8');
const frontend=fs.readFileSync(path.join(root,'Script.html'),'utf8');

function sourceFunction(source,name) {
  const start=source.indexOf('function '+name+'(');
  assert.ok(start>=0,'function missing: '+name);
  const end=source.indexOf('\nfunction ',start+9);
  assert.ok(end>start,'next function missing: '+name);
  return source.slice(start,end);
}
const code=sourceFunction(utils,'atividadeDentroPeriodoOficial_')+
  '\n'+sourceFunction(utils,'calcularStatusMeuGiroPorPercentual_')+'\n'+modal+'\n'+
  ['isStatusFinalFoco_','isDesafioAtivo_','classificarDesafioV2_',
   'isDesafioElegivelCertificado_'].map(name=>sourceFunction(frontend,name)).join('\n')+
  '\nreturn {sum:meuGiroSomarKmPorVinculo_,status:calcularStatusMeuGiroPorPercentual_,'+
  'isActive:isDesafioAtivo_,classify:classificarDesafioV2_,certificate:isDesafioElegivelCertificado_};';

const api=new Function('normalizarDataISO_','isDataIsoValida_','Utilities','parseLocalizedNumber_',
  'toFiniteNumber', code)(v=>String(v||''),v=>/^\d{4}-\d{2}-\d{2}$/.test(v),
  {formatDate:()=> '2026-10-15'},v=>Number(v),v=>Number(v)||0);
const enrollment={periodo_inicio:'2026-09-02',periodo_fim:'2026-11-30',
  modalidade_esportiva:'CICLISMO'};
const completed={...enrollment,status_usuario_desafio:'CONCLUIDO',
  status_apuracao:'CONCLUIDO',status_desafio:'Ativo',meta_km:1000,distancia_realizada:1000};

test('atividade anterior ao início nunca contribui para PRAZO_DIAS',()=>{
  assert.equal(api.sum([{data_atividade:'2026-09-01',km:1000,
    modalidade_esportiva:'CICLISMO'}],enrollment),0);
});

test('atividade do primeiro dia conta e pode completar a meta',()=>{
  const km=api.sum([{data_atividade:'2026-09-02',km:1000,
    modalidade_esportiva:'CICLISMO'}],enrollment);
  assert.equal(km,1000);
  assert.equal(api.status(km/1000*100),'CONCLUIDO');
});

test('a classificação atual respeita conclusão oficial antes do prazo',()=>{
  assert.deepEqual(api.classify(completed),
    {grupo:'concluidos',statusHumano:'Concluído'});
  assert.equal(api.isActive(completed),false);
  assert.equal(api.certificate(completed),true);
});

test('DIAGNÓSTICO: recomputar resumo incompatível pode desabilitar certificado sem tirar de Concluídos',()=>{
  const km=api.sum([{data_atividade:'2026-09-01',km:1000,
    modalidade_esportiva:'CICLISMO'}],enrollment);
  const after={...completed,distancia_realizada:km,status_apuracao:api.status(km/1000*100)};
  assert.equal(after.status_apuracao,'STATUS_EM_ANALISE');
  assert.equal(api.classify(after).grupo,'concluidos');
  assert.equal(api.isActive(after),false);
  assert.equal(api.certificate(after),false); // Falha de consistência documentada; NÃO é uma correção
});

test('não misturar encerramento do prazo e conclusão já validada',()=>{
  assert.equal(completed.periodo_fim,'2026-11-30');
  assert.equal(api.classify(completed).grupo,'concluidos');
});
