'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..', 'Meu Giro');
const source = fs.readFileSync(path.join(root, 'RegistroService.gs'), 'utf8');
const modalities = fs.readFileSync(path.join(root, 'ModalidadeEsportivaService.gs'), 'utf8');
const utils = fs.readFileSync(path.join(root, 'Utils.gs'), 'utf8');

const sourceEnd=source.indexOf('function meuGiroPodeConfirmarDuplicidadeEdicao_(');
assert.ok(sourceEnd>0);
const registeredSource=source.slice(0,sourceEnd);
const boundStart=utils.indexOf('function atividadeDentroPeriodoOficial_(');
const boundEnd=utils.indexOf('\nfunction normalizarPeriodoMensal_(',boundStart);
assert.ok(boundStart>=0 && boundEnd>boundStart);
const periodBound=utils.slice(boundStart,boundEnd);

const headers=['Timestamp','ID_DGMB','ID_INSCRICAO','ID_DESAFIO','id_item_estoque',
  'periodo_desafio','Data_Atividade','KM','Origem_Registro','Observacao',
  'Status_Validacao','activity_id','Modalidade_Esportiva'];
const idx={idxTimestamp:0,idxId:1,idxData:6,idxKm:7,idxOrigemRegistro:8,
  idxObservacao:9,idxStatusValidacao:10,idxActivityId:11,idxModalidade:12};
const binds=[
  {id_desafio:'153',meta_km:1000,modalidade_esportiva:'CICLISMO',
    periodo_inicio:'2026-09-01',periodo_fim:'2026-10-15'},
  {id_desafio:'155',meta_km:50,modalidade_esportiva:'CICLISMO',
    periodo_inicio:'2026-10-01',periodo_fim:'2026-10-31'},
  {id_desafio:'901601',meta_km:50,modalidade_esportiva:'CAMINHADA_CORRIDA',
    periodo_inicio:'2026-10-01',periodo_fim:'2026-10-31'},
  {id_desafio:'901602',meta_km:75,modalidade_esportiva:'CAMINHADA_CORRIDA',
    periodo_inicio:'2026-10-01',periodo_fim:'2026-10-31'}
].map((v,i)=>({...v,id_dgmb:'1603',linha_inscricao:i+2,apto:true}));
function fixtureRow(date, km, modality, activityId) {
  const r=Array(headers.length).fill('');
  r[0]='fixture';r[1]='1603';r[6]=date;r[7]=km;
  r[8]='MANUAL';r[11]=activityId;r[12]=modality;
  return r;
}
function setup(asOf) {
  const rows=[headers.slice(),
    fixtureRow('2026-10-12',5.5,'CAMINHADA_CORRIDA','walk5'),
    fixtureRow('2026-09-29',44,'CAMINHADA_CORRIDA','walk44'),
    fixtureRow('2026-09-30',10,'CICLISMO','s10'),
    fixtureRow('2026-09-30',2,'CICLISMO','s2'),
    fixtureRow('2026-10-01',3,'CICLISMO','o3a'),
    fixtureRow('2026-10-01',3,'CICLISMO','o3b'),
    fixtureRow('2026-08-31',1,'CICLISMO','a1')
  ];
  const enrolled=[['ID_DGMB','Distancia_Realizada'],...binds.map(()=>['1603',0])];
  let next=0;
  let summaryWrites=0;
  let rowWrites=0;
  const sheet={
    getDataRange:()=>({getValues:()=>rows.map(r=>r.slice())}),
    getLastColumn:()=>headers.length,
    getLastRow:()=>rows.length,
    getRange:(r,c,quantity,columns)=>({
      setValue:v=>{rows[r-1][c-1]=v;rowWrites++;},
      setValues:values=>{rows[r-1]=values[0].slice();rowWrites++;}
    }),
    deleteRow:(r)=>{rows.splice(r-1,1);}
  };
  const memberSheet={
    getRange:(r,c)=>({setValue:v=>{enrolled[r-1][c-1]=v;}})
  };
  const cache={values:enrolled,sheet:memberSheet,aba:'dgmbDesafios'};
  const summary={};
  const records=()=>rows.slice(1).filter(r=>r[1]==='1603').map(r=>({
    data_atividade:r[6],km:Number(r[7]),modalidade_esportiva:r[12],activity_id:r[11]
  }));
  const deps={
    MEU_GIRO_PERF_OPERACAO_ATUAL_:'',meuGiroPerfNow_:()=>0,
    meuGiroPerfLog_:()=>{},Logger:{log:()=>{}},
    LockService:{getScriptLock:()=>({waitLock:()=>{},releaseLock:()=>{}})},
    normalizarDataISO_:v=>String(v||''),
    isDataIsoValida_:v=>/^\d{4}-\d{2}-\d{2}$/.test(v) &&
      new Date(v+'T12:00:00Z').toISOString().slice(0,10)===v,
    parseKmInputSeguro_:v=>Number(String(v).replace(',','.')),
    normalizarKmEdicao_:v=>Number(v),
    kmsIguaisEdicao_:(a,b)=>Number(a)===Number(b),
    Utilities:{formatDate:()=>asOf,getUuid:()=>('temp-'+(++next))},
    getSpreadsheet_:()=>({getSheetByName:()=>sheet}),
    SHEETS:{REGISTRO_KM:'REGISTRO_KM'},
    getRegistroKmColumnIndexes_:()=>({...idx}),
    obterVinculosDesafioUsuario_:()=>binds,
    converterValoresRegistroKmEmObjetos_:values=>values.slice(1).map(r=>({
      ID_DGMB:r[1],Data_Atividade:r[6],KM:r[7],Modalidade_Esportiva:r[12],activity_id:r[11]
    })),
    obterRegistrosKmUsuario_:(id,ops)=>records(),
    obterDgmbDesafiosCacheExecucao_:()=>cache,
    buildHeaderMap_:()=>({id_dgmb:0,distancia_realizada:1}),
    getRequiredColumnIndex_:(map,aliases)=>aliases.some(x=>x.includes('distancia'))?1:0,
    atualizarMeuGiroResumoComLockAdquirido_:()=>{
      summaryWrites++;
      binds.forEach(v=>{summary[v.id_desafio]=Number(
        records().filter(r=>r.data_atividade<=asOf&&r.modalidade_esportiva===v.modalidade_esportiva&&
          r.data_atividade>=v.periodo_inicio&&r.data_atividade<=v.periodo_fim)
        .reduce((t,r)=>t+r.km,0).toFixed(2)
      )});
    }
  };
  const register=new Function(...Object.keys(deps),periodBound+'\n'+modalities+'\n'+registeredSource+
    '\nreturn registrarAtividade;')(...Object.values(deps));
  const totals=()=>binds.map(v=>enrolled[v.linha_inscricao-1][1]);
  const summaryTotals=()=>binds.map(v=>summary[v.id_desafio]);
  return {rows,register,summary,totals,summaryTotals,records,
    getWrites:()=>({rowWrites,summaryWrites}),asOf};
}

test('15/10: registro autêntico do serviço aumenta os dois desafios Ciclismo',()=>{
  const a=setup('2026-10-15');
  const res=a.register('1603','2026-10-15','1',false,'CICLISMO');
  assert.equal(res.ok,true);
  assert.equal(res.registros_criados,1);
  assert.equal(a.rows.length,9);
  assert.deepEqual(a.totals(),[19,7,5.5,5.5]);
  assert.deepEqual(a.summaryTotals(),[19,7,5.5,5.5]);
  assert.equal(a.records().filter(r=>r.data_atividade==='2026-10-15').length,1);
});

test('15/10: data 16/10 não pode gravar nem recalcular antes de virar o dia',()=>{
  const a=setup('2026-10-15');
  const res=a.register('1603','2026-10-16',1,false,'CICLISMO');
  assert.equal(res.ok,false);
  assert.equal(res.code,'DATA_FUTURA');
  assert.equal(a.rows.length,8);
  assert.deepEqual(a.getWrites(),{rowWrites:0,summaryWrites:0});
});

test('16/10: registro histórico do dia seguinte preserva prazo, avança Outubro Rosa',()=>{
  const a=setup('2026-10-16');
  const res=a.register('1603','2026-10-16',1,false,'CICLISMO');
  assert.equal(res.ok,true);
  assert.deepEqual(a.totals(),[18,7,5.5,5.5]);
  assert.deepEqual(a.summaryTotals(),[18,7,5.5,5.5]);
  assert.equal(a.rows.length,9);
});

test('16/10: lançamento retroativo de 15/10 ainda pertence aos 45 dias',()=>{
  const a=setup('2026-10-16');
  assert.equal(a.register('1603','2026-10-15',1,false,'CICLISMO').ok,true);
  assert.deepEqual(a.totals(),[19,7,5.5,5.5]);
});

test('duas modalidades: sem escolha, backend rejeita registro antes de gravar',()=>{
  const a=setup('2026-10-15');
  const res=a.register('1603','2026-10-15',1,false,'');
  assert.equal(res.code,'MODALIDADE_OBRIGATORIA');
  assert.equal(a.rows.length,8);
});

test('mesma data e km já registrados são protegidos contra duplicação acidental',()=>{
  const a=setup('2026-10-15');
  const res=a.register('1603','2026-10-01',3,false,'CICLISMO');
  assert.equal(res.code,'DUPLICIDADE');
  assert.equal(a.rows.length,8);
  assert.deepEqual(a.getWrites(),{rowWrites:0,summaryWrites:0});
});

test('atividade caminhada no último dia atualiza só as duas metas da modalidade',()=>{
  const a=setup('2026-10-15');
  const res=a.register('1603','2026-10-15',1,false,'CAMINHADA_CORRIDA');
  assert.equal(res.ok,true);
  assert.deepEqual(a.totals(),[18,6,6.5,6.5]);
  assert.deepEqual(a.summaryTotals(),[18,6,6.5,6.5]);
});

test('atleta preserva histórico quando o prazo acabou, sem somar 31/08 ao 1000 KM',()=>{
  const a=setup('2026-10-16');
  const res=a.register('1603','2026-08-30',2,false,'CICLISMO');
  assert.equal(res.ok,true);
  assert.deepEqual(a.totals(),[18,6,5.5,5.5]);
  assert.equal(a.rows.length,9);
});
