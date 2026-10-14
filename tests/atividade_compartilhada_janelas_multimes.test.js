'use strict';
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const test=require('node:test');
const root=path.resolve(__dirname,'..','Meu Giro');
const utils=fs.readFileSync(path.join(root,'Utils.gs'),'utf8');
const modal=fs.readFileSync(path.join(root,'ModalidadeEsportivaService.gs'),'utf8');

const periodo=(data,inicio,fim)=>data>=inicio && data<=fim;
function dataValida(valor) {
  if(!/^\d{4}-\d{2}-\d{2}$/.test(valor))return false;
  const d=new Date(valor+'T12:00:00Z');
  return !Number.isNaN(d.getTime()) && d.toISOString().slice(0,10)===valor;
}
function calculador(hoje) {
  const iso=valor=>String(valor||'');
  return new Function(
    'atividadeDentroPeriodoOficial_','normalizarDataISO_','isDataIsoValida_','Utilities',
    modal+'\nreturn {sum:meuGiroSomarKmPorVinculo_, contains:meuGiroAtividadeNoVinculo_};'
  )(periodo,iso,dataValida,{formatDate:()=>hoje});
}

test('90 dias iniciados em 01/09 contam setembro, outubro e novembro até 29/11',()=>{
  // 01/09 é o primeiro dia; o nonagésimo é 29/11, não 30/11.
  const inicio=new Date('2026-09-01T12:00:00Z');
  const nonagesimo=new Date(inicio.getTime()+89*86400000).toISOString().slice(0,10);
  assert.equal(nonagesimo,'2026-11-29');
  const c=calculador('2026-12-10');
  const registros=[
    {data_atividade:'2026-09-01',km:5,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-09-30',km:10,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-10-01',km:15,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-10-15',km:20,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-10-31',km:25,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-11-01',km:30,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-11-29',km:35,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-11-30',km:40,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-10-10',km:7,modalidade_esportiva:'CAMINHADA_CORRIDA'},
    {data_atividade:'2026-12-01',km:100,modalidade_esportiva:'CICLISMO'}
  ];
  const desafio=(periodo_inicio,periodo_fim,modalidade_esportiva='CICLISMO')=>({
    periodo_inicio,periodo_fim,modalidade_esportiva
  });
  assert.equal(c.sum(registros,desafio('2026-09-01','2026-11-29')),140);
  assert.equal(c.sum(registros,desafio('2026-09-01','2026-09-30')),15);
  assert.equal(c.sum(registros,desafio('2026-10-01','2026-10-31')),60);
  assert.equal(c.sum(registros,desafio('2026-11-01','2026-11-30')),105);
  // A mesma atividade de outubro pode contar para o mensal e o de 90 dias.
  assert.equal(c.contains(registros[2],desafio('2026-09-01','2026-11-29')),true);
  assert.equal(c.contains(registros[2],desafio('2026-10-01','2026-10-31')),true);
  assert.equal(c.contains(registros[7],desafio('2026-09-01','2026-11-29')),false);
  assert.equal(c.sum(registros,desafio('2026-10-01','2026-10-31','CAMINHADA_CORRIDA')),7);
});

test('fixture LAB de 45 dias passa de setembro a outubro sem herdar período mensal',()=>{
  const c=calculador('2026-10-16');
  const registros=[
    {data_atividade:'2026-08-31',km:3,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-09-01',km:4,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-09-30',km:5,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-10-01',km:6,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-10-15',km:7,modalidade_esportiva:'CICLISMO'},
    {data_atividade:'2026-10-16',km:8,modalidade_esportiva:'CICLISMO'}
  ];
  const vinculo={periodo_inicio:'2026-09-01',periodo_fim:'2026-10-15',modalidade_esportiva:'CICLISMO'};
  assert.equal(c.sum(registros,vinculo),22);
  assert.equal(c.contains(registros[0],vinculo),false);
  assert.equal(c.contains(registros[5],vinculo),false);
});

test('leitura leve usa contexto local do catálogo para modalidades de inscrições antigas',()=>{
  const a=utils.indexOf('function buildPeriodosDgmbDesafiosPorChave_(');
  const b=utils.indexOf('\nfunction meuGiroResumoPossuiInscricaoAusente_',a);
  assert.ok(a>=0&&b>a);
  const trecho=utils.slice(a,b);
  const match=trecho.match(/var statusDgmb = \{[\s\S]*?\n    \};/);
  assert.ok(match,'bloco de modalidade do vínculo deve existir');
  const status= new Function(
    'row','idxModalidadeInscricao','idDesafio','contextoLista','meuGiroModalidadeCanonica_','normalizeText_',
    'var statusUsuarioDesafio="",idxStatusValidacaoCertificado=-1,idxStatusDesafio=-1,'+
    'idxStatusPag=-1,tipoMeta="KM",prazoDias=0;\n'+match[0]+'\nreturn statusDgmb.modalidade_esportiva;'
  );
  const canon=v=>v||'CICLISMO';
  const contextoLista={modalidade:{byId:{'901601':'CAMINHADA_CORRIDA','153':'CICLISMO'}}};
  assert.equal(status([''],0,'901601',contextoLista,canon,v=>String(v||'')),'CAMINHADA_CORRIDA');
  assert.equal(status([''],0,'153',contextoLista,canon,v=>String(v||'')),'CICLISMO');
  assert.equal(status(['CICLISMO'],0,'901601',contextoLista,canon,v=>String(v||'')),'CICLISMO');
  assert.equal(status([''], -1,'DESCONHECIDO',contextoLista,canon,v=>String(v||'')),'CICLISMO');
  assert.doesNotMatch(trecho,/contextoListaDesafios\.modalidade/);
});

test('cálculo de resumo continua independente por vínculo e não soma atividades fora da modalidade',()=>{
  const a=utils.indexOf('function atualizarMeuGiroResumoComLockAdquirido_(');
  const b=utils.indexOf('\nfunction atualizarMeuGiroResumoEmLote_',a);
  assert.ok(a>=0&&b>a);
  const fonte=utils.slice(a,b);
  assert.match(fonte,/for \(var v = 0; v < vinculos\.length; v\+\+\)/);
  assert.match(fonte,/for \(var r = 0; r < registros\.length; r\+\+\)/);
  assert.match(fonte,/meuGiroAtividadeNoVinculo_\(reg, vinculo\)/);
});
