'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const source = fs.readFileSync(path.resolve(__dirname, '../Meu Giro/RankingService.gs'), 'utf8');
function extrair(nome) {
  const inicio = source.indexOf('function ' + nome + '(');
  const fim = source.indexOf('\nfunction ', inicio + 5);
  assert.ok(inicio >= 0, 'função ausente: ' + nome);
  return source.slice(inicio, fim < 0 ? undefined : fim);
}
const primeiro = (row, aliases) => aliases.map(x => row[x]).find(v => v !== undefined && v !== null && String(v) !== '') || '';
const catalogo = {
  periodos: {byId:{
    '155': {inicio:'2026-10-01',fim:'2026-10-31'},
    '901601': {inicio:'2026-10-01',fim:'2026-10-31'},
    '1000': {inicio:'2026-09-01',fim:'2026-09-30'}
  },byAba:{}},
  tipoMeta:{byId:{'155':'KM','901601':'KM','1000':'PRAZO_DIAS'}},
  modalidade:{byId:{'155':'CICLISMO','901601':'CAMINHADA_CORRIDA','1000':'CICLISMO'}}
};
const rows=[
  {ID_DGMB:'1603',ID_INSCRICAO:'bike',Observacao:'[ID_DESAFIO:155]',Distancia_KM:50,
    tipo_do_desafio:'Normal',periodo_desafio:'outubro/2026',
    data_inicio_desafio:'2026-10-01',data_fim_desafio:'2026-10-31'},
  {ID_DGMB:'1603',ID_INSCRICAO:'walk',Observacao:'[ID_DESAFIO:901601]',Distancia_KM:50,
    tipo_do_desafio:'Normal',periodo_desafio:'outubro/2026',
    data_inicio_desafio:'2026-10-01',data_fim_desafio:'2026-10-31'},
  {ID_DGMB:'1900',ID_INSCRICAO:'days',Observacao:'[ID_DESAFIO:1000]',Distancia_KM:1000,
    tipo_do_desafio:'Normal',periodo_desafio:'setembro/2026',
    data_inicio_desafio:'2026-09-01',data_fim_desafio:'2026-10-15'}
];
const deps={
  painelMG_incrementarAuditoriaCarregamentoInicial_:()=>{},
  getAllObjects_:()=>rows,SHEETS:{DESAFIO:'dgmbDesafios'},
  buildListaDesafiosContexto_:()=>catalogo,getSpreadsheet_:()=>({}),
  rankingMG_obterIdInscricao_:r=>r.ID_INSCRICAO,rankingMG_obterIdDgmb_:r=>r.ID_DGMB,
  rankingMG_norm_:v=>String(v||''),rankingMG_firstFilled_:primeiro,
  extrairIdDesafioObservacao_:v=>String(v||'').match(/ID_DESAFIO:(\d+)/)?.[1]||'',
  rankingMG_normalizarTipo_:v=>String(v||'').toUpperCase(),
  resolverTipoMetaListaDesafio_:(map,id)=>map.byId[id]||'',
  rankingMG_toNumber_:Number,meuGiroModalidadeCanonica_:v=>String(v||'CICLISMO').toUpperCase(),
  normalizarDataISO_:v=>String(v||''),
  extrairPeriodoDesafioTexto_:v=>v==='setembro/2026'?
    {inicio:'2026-09-01',fim:'2026-09-30'}:{inicio:'2026-10-01',fim:'2026-10-31'},
  periodoCompletoValido_:v=>!!(v&&v.inicio&&v.fim&&v.inicio<=v.fim),
  ehTipoMetaPrazoDias_:v=>v==='PRAZO_DIAS',
  rankingMG_chaveInscricao_:(id,ins)=>id+'|'+ins,
  rankingMG_validarAtributosCompetitivos_:v=>({valido:!!v.id_desafio}),
  rankingMG_normalizarDataCompetitiva_:v=>String(v||''),
  rankingMG_formatarMetaCompetitiva_:v=>Number(v).toFixed(1),
  rankingMG_obterItemCompetitivo_:v=>v.id_item_estoque||''
};
const helpers=new Function(...Object.keys(deps),
  ['rankingMG_resolverPeriodoCompetitivo_','rankingMG_criarIndiceCompetitivo_',
   'rankingMG_montarChaveCompetitiva_'].map(extrair).join('\n')+
  '\nreturn {periodo:rankingMG_resolverPeriodoCompetitivo_,indice:rankingMG_criarIndiceCompetitivo_,chave:rankingMG_montarChaveCompetitiva_};'
)(...Object.values(deps));

test('prazo individual passa pela virada setembro-outubro no ranking',()=>{
  const periodo=helpers.periodo(rows[2],catalogo.periodos.byId['1000'],'PRAZO_DIAS');
  assert.deepEqual(periodo,{inicio:'2026-09-01',fim:'2026-10-15'});
  assert.deepEqual(helpers.periodo({...rows[2],data_fim_desafio:''},
    catalogo.periodos.byId['1000'],'PRAZO_DIAS'),{inicio:'',fim:''});
  assert.deepEqual(helpers.periodo(rows[2],catalogo.periodos.byId['1000'],'KM'),
    {inicio:'2026-09-01',fim:'2026-09-30'});
});
test('ranking preserva modalidade mesmo em inscrição legada sem coluna',()=>{
  const ix=helpers.indice().por_inscricao;
  assert.equal(ix['1603|bike'].modalidade_esportiva,'CICLISMO');
  assert.equal(ix['1603|walk'].modalidade_esportiva,'CAMINHADA_CORRIDA');
  assert.equal(ix['1900|days'].periodo_fim,'2026-10-15');
});
test('mesma competição exige esporte e período equivalentes',()=>{
  const base={id_desafio:'1000',tipo_desafio:'Normal',meta_km:1000,
    periodo_inicio:'2026-09-01',periodo_fim:'2026-10-15',modalidade_esportiva:'CICLISMO'};
  const chave=helpers.chave(base);
  assert.equal(chave,helpers.chave({...base}));
  assert.notEqual(chave,helpers.chave({...base,modalidade_esportiva:'CAMINHADA_CORRIDA'}));
  assert.notEqual(chave,helpers.chave({...base,periodo_fim:'2026-10-30'}));
});
