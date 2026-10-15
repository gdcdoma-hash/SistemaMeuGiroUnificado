'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const source=fs.readFileSync(path.join(__dirname,'..','Meu Giro','CertificadoService.gs'),'utf8');
const start=source.indexOf('function certificadoBuscarStatusApuracaoResumo_(');
assert.ok(start>=0,'lookup function missing');
const lookup=source.slice(start); // última função do arquivo

const header=['ID_INSCRICAO','ID_DGMB','ID_DESAFIO','id_item_estoque',
  'Meta_KM','Status_Apuracao'];
const row=(ins,athlete,challenge,item,status)=>[ins,athlete,challenge,item,1000,status];

function build(rows=[]) {
  let reads=0, writes=0, recalcs=0;
  const table=[header,...rows];
  const sheet={
    getDataRange:()=>({getValues:()=>{reads++;return table.map(r=>r.slice());}}),
    getRange:()=>{writes++;throw Error('write attempted');},
    appendRow:()=>{writes++;throw Error('write attempted');}
  };
  const deps={
    normalizeText_:v=>String(v==null?'':v).trim(),
    getSpreadsheet_:()=>({getSheetByName:name=>name==='MEU_GIRO_RESUMO'?sheet:null}),
    SHEETS:{MEU_GIRO_RESUMO:'MEU_GIRO_RESUMO'},
    buildHeaderMap_:headers=>Object.fromEntries(headers.map((h,i)=>[h.toLowerCase(),i])),
    getOptionalColumnIndex_:(map,aliases)=>{
      for(const alias of aliases){const n=map[alias.toLowerCase()];if(n!==undefined)return n;}
      return -1;
    },
    atualizarMeuGiroResumo_:()=>{recalcs++;throw Error('recalc is forbidden');}
  };
  const fn=new Function(...Object.keys(deps),lookup+
    '\nreturn certificadoBuscarStatusApuracaoResumo_;')(...Object.values(deps));
  return {query:(ins,athlete='fake-user',challenge='X',item='')=>
    fn(ins,athlete,challenge,item),stats:()=>({reads,writes,recalcs})};
}

test('consulta em inscrição única concluída não recalcula nem escreve',()=>{
  const s=build([row('enrollment-A','fake-user','X','','CONCLUIDO')]);
  assert.equal(s.query('enrollment-A'),'CONCLUIDO');
  assert.deepEqual(s.stats(),{reads:1,writes:0,recalcs:0});
});

test('resumo pendente continua pendente, não gera conclusão implícita',()=>{
  const s=build([row('enrollment-A','fake-user','X','','STATUS_EM_ANALISE')]);
  assert.equal(s.query('enrollment-A'),'STATUS_EM_ANALISE');
  assert.equal(s.stats().recalcs,0);
});

test('inscrições distintas do mesmo atleta/desafio são identificadas por inscrição',()=>{
  const s=build([row('enrollment-A','fake-user','X','','STATUS_EM_ANALISE'),
    row('enrollment-B','fake-user','X','','CONCLUIDO')]);
  assert.equal(s.query('enrollment-B'),'CONCLUIDO');
  assert.equal(s.query('enrollment-A'),'STATUS_EM_ANALISE');
});

test('id de participante incorreto ou inscrição inexistente não obtém certificado',()=>{
  const s=build([row('enrollment-A','fake-user','X','','CONCLUIDO')]);
  assert.equal(s.query('enrollment-A','other-athlete'),'');
  assert.equal(s.query('enrollment-X'),'');
});

test('mesmo ID_INSCRICAO repetido falha fechado',()=>{
  const s=build([row('enrollment-A','fake-user','X','','CONCLUIDO'),
    row('enrollment-A','fake-user','X','','CONCLUIDO')]);
  assert.equal(s.query('enrollment-A'),'');
});

test('busca legada sem inscrição só retorna linha sem ID_INSCRICAO',()=>{
  const s=build([row('','fake-user','X','ITEM-1','CONCLUIDO'),
    row('enrollment-A','fake-user','X','ITEM-1','STATUS_EM_ANALISE')]);
  assert.equal(s.query('','fake-user','X','ITEM-1'),'CONCLUIDO');
  assert.equal(s.query('','fake-user','X','ITEM-2'),'');
});

test('busca legada ambígua é recusada',()=>{
  const s=build([row('','fake-user','X','ITEM-1','CONCLUIDO'),
    row('','fake-user','X','ITEM-1','CONCLUIDO')]);
  assert.equal(s.query('','fake-user','X','ITEM-1'),'');
});

test('sem resumo persistido não deve criar ou atualizar resumo',()=>{
  const s=build();
  assert.equal(s.query('enrollment-A'),'');
  assert.deepEqual(s.stats(),{reads:1,writes:0,recalcs:0});
});

test('falha de leitura é tratada sem conceder certificado',()=>{
  const dependencies={
    normalizeText_:v=>String(v||'').trim(),
    getSpreadsheet_:()=>{throw Error('simulated outage')},
    SHEETS:{MEU_GIRO_RESUMO:'MEU_GIRO_RESUMO'},
    buildHeaderMap_:()=>({}),getOptionalColumnIndex_:()=>-1
  };
  const fn=new Function(...Object.keys(dependencies),lookup+
    '\nreturn certificadoBuscarStatusApuracaoResumo_;')(...Object.values(dependencies));
  assert.equal(fn('enrollment-A','fake-user','X',''),'');
});
