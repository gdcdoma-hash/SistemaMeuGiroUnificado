'use strict';
const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');

const source=fs.readFileSync(path.join(__dirname,'..','Meu Giro','CertificadoService.gs'),'utf8');
const index=source.indexOf('function certificadoBuscarStatusApuracaoResumo_(');
assert.ok(index>=0);
const fSource=source.slice(index);
const header=['ID_DGMB','ID_INSCRICAO','ID_DESAFIO','id_item_estoque',
  'Meta_KM','Distancia_Realizada','Status_Apuracao'];
const row=(id,ins,km,meta=100,status='CONCLUIDO')=>[id,ins,'X','ITEM',meta,km,status];
const normalize=v=>String(v==null?'':v).trim();
const normKey=v=>normalize(v).toLowerCase().normalize('NFD')
  .replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'');

function fixture(rows) {
  let writes=0,reads=0;
  const sheet={
    getDataRange:()=>({getValues:()=>{reads++;return [header,...rows].map(r=>r.slice())}}),
    getRange:()=>{writes++;throw Error('write')}
  };
  const dep={
    normalizeText_:normalize,
    getSpreadsheet_:()=>({getSheetByName:()=>sheet}),
    SHEETS:{MEU_GIRO_RESUMO:'MEU_GIRO_RESUMO'},
    buildHeaderMap_:headers=>Object.fromEntries(headers.map((x,i)=>[normKey(x),i])),
    getOptionalColumnIndex_:(map,aliases)=>{
      for(const item of aliases){const i=map[normKey(item)];if(i!==undefined)return i;}
      return -1;
    },
    parseLocalizedNumber_:v=>Number(String(v==null?'0':v).replace(',','.')),
    atualizarMeuGiroResumo_:()=>{writes++;throw Error('recalc')}
  };
  const method=new Function(...Object.keys(dep),fSource+
    '\nreturn certificadoBuscarStatusApuracaoResumo_;')(...Object.values(dep));
  return {read:(ins='i1',athlete='fake')=>method(ins,athlete,'X','ITEM',{detalharMeta:true}),
    legacy:(ins='i1',athlete='fake')=>method(ins,athlete,'X','ITEM'),
    stats:()=>({writes,reads})};
}
test('100% exato gera meta_atingida true com consulta somente leitura',()=>{
  const a=fixture([row('fake','i1',100)]);
  assert.deepEqual(a.read(),{status_apuracao:'CONCLUIDO',meta_atingida:true,meta_km:100,distancia_realizada:100});
  assert.deepEqual(a.stats(),{writes:0,reads:1});
});

test('99,9% não é 100%',()=>{
  assert.equal(fixture([row('fake','i1',99.9)]).read().meta_atingida,false);
});

test('100,5% é válido e não requer exatamente 100',()=>{
  assert.equal(fixture([row('fake','i1',100.5)]).read().meta_atingida,true);
});

test('status CONCLUIDO não pode converter 0 km em 100%',()=>{
  const v=fixture([row('fake','i1',0)]).read();
  assert.equal(v.meta_atingida,false);
  assert.equal(v.status_apuracao,'CONCLUIDO');
});

test('meta zero não pode autorizar certificado',()=>{
  assert.equal(fixture([row('fake','i1',0,0)]).read().meta_atingida,false);
});

test('ID inscricao incorreto ou participante incorreto não recupera KM',()=>{
  const a=fixture([row('fake','i1',100)]);
  assert.equal(a.read('i2'),'');
  assert.equal(a.read('i1','outro'),'');
});

test('duplicidade de ID_INSCRICAO falha sem tentar usar outra linha',()=>{
  assert.equal(fixture([row('fake','i1',100),row('fake','i1',100)]).read(),'');
});

test('chamada legada sem detalhes mantém retorno de texto',()=>{
  assert.equal(fixture([row('fake','i1',100)]).legacy(),'CONCLUIDO');
});

test('não tenta gravar/reapurar dados em nenhuma leitura',()=>{
  const a=fixture([row('fake','i1',99.9)]);
  for(let i=0;i<3;i++)a.read();
  assert.deepEqual(a.stats(),{writes:0,reads:3});
});
