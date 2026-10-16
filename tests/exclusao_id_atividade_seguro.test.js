'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const source=fs.readFileSync(path.join(__dirname,'..','Meu Giro','RegistroService.gs'),'utf8');
const init=source.indexOf('function excluirAtividade(');
const stop=source.indexOf('\nfunction getRegistroKmColumnIndexes_(',init);
const loc=source.indexOf('function localizarLinhasAtividade_(');
assert.ok(init>=0 && stop>init && loc>stop);
const excluded=source.slice(init,stop);
const locator=source.slice(loc);
const headers=['Timestamp','ID_DGMB','Data_Atividade','KM','activity_id','Modalidade_Esportiva'];
const row=(t,d,km,aid,modo='CICLISMO')=>[t,'1603',d,km,aid,modo];
const dataset=()=>[
  headers.slice(),
  row('s10','2026-09-30',10,'sep-10'),
  row('s2','2026-09-30',2,'sep-2'),
  row('a','2026-10-01',3,'a-id'),
  row('b','2026-10-01',3,'b-id'),
  row('walk','2026-10-12',5.5,'walk-id','CAMINHADA_CORRIDA'),
  row('temp','2026-10-02',2,'temp-id')
];

function execute(payload, initial=dataset(), throwOnRecalc=false) {
  const table=initial.map(x=>x.slice());
  const original=JSON.stringify(table);
  const idx={idxTimestamp:0,idxId:1,idxData:2,idxKm:3,idxActivityId:4,idxModalidade:5};
  let del=0, recalc=0, rollback=0;
  const sheet={
    getDataRange:()=>({getValues:()=>table.map(x=>x.slice())}),
    deleteRow:(n)=>{del++;table.splice(n-1,1)},
    insertRowBefore:(n)=>{rollback++;table.splice(n-1,0,[])},
    getRange:(rowNum,col,numRows,numCols)=>({
      setValues:(matrix)=>{table[rowNum-1]=matrix[0].slice();}
    })
  };
  const deps={
    MEU_GIRO_PERF_OPERACAO_ATUAL_:'',
    meuGiroPerfNow_:()=>0,meuGiroPerfLog_:()=>{},
    // A aplicação PROD consulta a trava de manutenção antes de editar os dados.
    // Aqui simulamos explicitamente o estado LIBERADO; nenhum GAS real é chamado.
    dgmbCorteGravacaoPausada_: () => false,
    LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
    Logger:{log(){}},
    getSpreadsheet_:()=>({getSheetByName:()=>sheet}),
    SHEETS:{REGISTRO_KM:'REGISTRO_KM'},
    getRegistroKmColumnIndexes_:()=>idx,
    meuGiroModalidadeCanonica_:v=>String(v||'CICLISMO'),
    normalizarTimestampEdicao_:v=>String(v||''),
    criarOpcoesRegistroKmReaproveitado_:()=>({}),
    atualizarDistanciaRealizada_:()=>{},
    atualizarMeuGiroResumoComLockAdquirido_:()=>{recalc++;if(throwOnRecalc)throw Error('simulated summary failure');}
  };
  const fn=new Function(...Object.keys(deps),locator+'\n'+excluded+'\nreturn excluirAtividade;')(...Object.values(deps));
  const result=fn({id_dgmb:'1603',activity_id:'temp-id',chave_edicao:'temp',...payload});
  const sum=(modo,begin,end)=>table.slice(1).filter(r=>r[5]===modo&&r[2]>=begin&&r[2]<=end).reduce((n,r)=>n+Number(r[3]),0);
  return {result,table,original,del,recalc,rollback,sum,
    sameAsOriginal:JSON.stringify(table)===original};
}

test('apaga apenas o registro temporário e recalcula ambos os desafios ciclismo',()=>{
  const s=execute({});
  assert.equal(s.result.ok,true);
  assert.equal(s.result.registros_excluidos,1);
  assert.equal(s.del,1);
  assert.equal(s.recalc,1);
  assert.equal(s.sum('CICLISMO','2026-09-01','2026-10-15'),18);
  assert.equal(s.sum('CICLISMO','2026-10-01','2026-10-31'),6);
  assert.equal(s.sum('CAMINHADA_CORRIDA','2026-10-01','2026-10-31'),5.5);
  assert.equal(s.table.filter(r=>r[4]==='a-id'||r[4]==='b-id').length,2);
});

test('duas atividades com mesma data e KM mantêm IDs diferentes: excluir uma não remove a outra',()=>{
  const s=execute({activity_id:'a-id',chave_edicao:'a'},dataset().slice(0,-1));
  assert.equal(s.result.ok,true);
  assert.equal(s.del,1);
  assert.equal(s.table.some(r=>r[4]==='a-id'),false);
  assert.equal(s.table.some(r=>r[4]==='b-id'),true);
  assert.equal(s.sum('CICLISMO','2026-10-01','2026-10-31'),3);
});

test('activity_id incorreto não recorre ao timestamp de outro registro',()=>{
  const s=execute({activity_id:'id-nao-existe',chave_edicao:'a'});
  assert.equal(s.result.ok,false);
  assert.equal(s.result.code,'IDENTIFICADOR_ATIVIDADE_INCONSISTENTE');
  assert.equal(s.del,0);
  assert.equal(s.recalc,0);
  assert.equal(s.sameAsOriginal,true);
});

test('registro legado sem activity_id e timestamp único pode ser excluído',()=>{
  const rows=dataset();
  rows.push(row('old','2026-08-15',1,'','CICLISMO'));
  const s=execute({activity_id:'',chave_edicao:'old'},rows);
  assert.equal(s.result.ok,true);
  assert.equal(s.del,1);
  assert.equal(s.table.some(r=>r[0]==='old'),false);
});

test('registro legado ambíguo, mesmo timestamp, fica bloqueado',()=>{
  const rows=dataset();
  rows.push(row('old','2026-08-15',1,'','CICLISMO'));
  rows.push(row('old','2026-08-15',1,'','CICLISMO'));
  const s=execute({activity_id:'',chave_edicao:'old'},rows);
  assert.equal(s.result.code,'ATIVIDADE_AMBIGUA');
  assert.equal(s.del,0);
  assert.equal(s.sameAsOriginal,true);
});

test('registros legados de modalidades diferentes com o mesmo timestamp não são excluídos juntos',()=>{
  const rows=dataset();
  rows.push(row('old','2026-08-15',1,'','CICLISMO'));
  rows.push(row('old','2026-08-15',1,'','CAMINHADA_CORRIDA'));
  const s=execute({activity_id:'',chave_edicao:'old'},rows);
  assert.equal(s.result.ok,false);
  assert.equal(s.del,0);
  assert.equal(s.sameAsOriginal,true);
});

test('se o resumo falha, restaura a linha excluída na planilha simulada',()=>{
  const s=execute({},dataset(),true);
  assert.equal(s.result.ok,false);
  assert.equal(s.result.code,'EXCLUSAO_ATIVIDADE_ERROR');
  assert.equal(s.del,1);
  assert.equal(s.rollback,1);
  assert.equal(s.sameAsOriginal,true);
});
