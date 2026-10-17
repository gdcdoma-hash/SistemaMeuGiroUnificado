'use strict';
// Testes de interface sem escrita de dados, sem clasp e sem dependência de planilha.
// A base do patch é exatamente o Meu Giro PROD @10, mas este arquivo jamais é enviado ao GAS.
const assert=require('node:assert/strict');
const fs=require('node:fs');
const cp=require('node:child_process');
const vm=require('node:vm');
const base='3c2bb257ad034bb80174472e723242d642d697ee';
const script=fs.readFileSync('Meu Giro/Script.html','utf8');
const html=fs.readFileSync('Meu Giro/Index.html','utf8');
const css=fs.readFileSync('Meu Giro/Styles.html','utf8');
const backend=fs.readFileSync('Meu Giro/PainelService.gs','utf8');
const backendOriginal=cp.execFileSync('git',['show',base+':Meu Giro/PainelService.gs'],{encoding:'utf8'});
assert.equal(backend,backendOriginal,'O backend e a regra de status NÃO podem ser modificados');
for(const path of ['Meu Giro/RegistroService.gs','Meu Giro/CertificadoService.gs']){
 const now=fs.readFileSync(path,'utf8');
 assert.equal(now,cp.execFileSync('git',['show',base+':'+path],{encoding:'utf8'}),path+' não deve ser modificado');
}
function codeBetween(start,end){
 const a=script.indexOf('function '+start+'('),b=script.indexOf('function '+end+'(',a+1);
 assert.ok(a>=0&&b>a,'Função ausente: '+start);
 return script.slice(a,b);
}
const chunks=[
 codeBetween('dataAtualBrasiliaIso_','buildDesafioCardV2_'),
 codeBetween('parseIsoDateLocal','calcularRitmoDesafio'),
 codeBetween('obterTempoRestanteDesafioPainel_','atualizarResumoPrazoDesafioPainel_'),
 codeBetween('classificarDesafioV2_','renderGrupoDesafiosV2_'),
 codeBetween('isStatusFinalFoco_','isDesafioHistoricoElegivel_'),
 codeBetween('renderDesafiosV2','isConquistaActionLocked_'),
 codeBetween('buildDesafioCardV2_','calcularPercentualDesafio_')
].join('\n');
const baseItem={
 id_inscricao:'FICTICIA_45',id_desafio:'1000',nome_desafio:'1000 KM',
 tipo_meta:'PRAZO_DIAS',prazo_dias:45,
 periodo_inicio:'2026-09-01',periodo_fim:'2026-10-15',
 meta_km:1000,distancia_realizada:18,
 status_usuario_desafio:'EM_ANDAMENTO',status_apuracao:'STATUS_EM_ANALISE',
 status_desafio:'ATIVO',status_lista_desafios:'ATIVO'
};
let total=0;
function ok(name,fn){try{fn();total++;console.log('OK '+name)}catch(e){console.error('FALHOU '+name);throw e;}}
function load(now){
 let fakeNow=new Date(now||'2026-10-16T15:00:00Z').getTime();
 class FakeDate extends Date{
  constructor(...args){super(...(args.length?args:[fakeNow]));}
  static now(){return fakeNow;}
 }
 const captures={groups:{}};
 const ids={};
 const fakeDocument={
  getElementById(name){return ids[name]||null;},
  createElement(){return {innerHTML:'',setAttribute(){},addEventListener(){},className:'',tabIndex:0};}
 };
 const ctx={Date:FakeDate,Intl,console,document:fakeDocument,
  currentPainel:{desafios:[baseItem]}, getDesafioBaseKey_:d=>d.id_inscricao,
  obterTodosDesafiosPainelV2_:p=>p.desafios||[],
  renderGrupoDesafiosV2_:(s,items,group)=>{captures.groups[group]=items.map(x=>x.desafio.id_inscricao)},
  meuGiroIdentidadeModalidade_:()=>({classe:'mg-modalidade-ciclismo',icone:'🚴',nome:'Ciclismo'}),
  debugPeriodoDesafioFrontend_(){},
  getPeriodoDesafioVisual_:()=> '01/09 a 15/10/26',
  toFiniteNumber:(x,f)=>Number(x)||f,
  calcularPercentualDesafio_:(a,b)=>a>0?b/a*100:0,
  getDesafioKey:d=>d.id_inscricao,
  getDesafioPrazoCardInfo_:()=>({titulo:'Prazo',valor:'45 dias',complemento:'01/09 a 15/10/26'}),
  getDesafioTitulo:d=>d.nome_desafio,
  escapeHtml:String,formatNumber:String
 };
 ids['desafios-v2-msg']={innerText:'',className:''};
 ids['desafios-v2-andamento']={};
 ids['desafios-v2-concluidos']={};
 ids['desafios-v2-andamento-title']={textContent:''};
 vm.createContext(ctx);
 vm.runInContext(chunks,ctx);
 return {ctx,ids,captures};
}
const before=load('2026-10-16T02:59:59Z'); // 15/10, 23:59:59 em Brasília
const after=load('2026-10-16T03:00:01Z'); // 16/10, 00:00:01 em Brasília
ok('15/10 inclusivo em America/Sao_Paulo, independentemente do dispositivo',()=>{
 assert.equal(before.ctx.prazoIndividualEncerradoVisual_(baseItem),false);
 assert.equal(before.ctx.obterTempoRestanteDesafioPainel_(baseItem.periodo_fim),'Último dia');
});
ok('16/10 após meia-noite de Brasília indica encerramento',()=>{
 assert.equal(after.ctx.prazoIndividualEncerradoVisual_(baseItem),true);
 assert.equal(after.ctx.obterTempoRestanteDesafioPainel_(baseItem.periodo_fim),'Encerrado');
});
ok('relógio explícito não afeta o desafio, permite teste sem gravar',()=>{
 assert.equal(after.ctx.prazoIndividualEncerradoVisual_(baseItem,'2026-10-15'),false);
 assert.equal(after.ctx.prazoIndividualEncerradoVisual_(baseItem,'2026-10-16'),true);
});
ok('contador visual preserva 1 dia restante antes de 15/10',()=>{
 const diaAnterior=load('2026-10-15T03:00:01Z');
 assert.equal(diaAnterior.ctx.obterTempoRestanteDesafioPainel_(baseItem.periodo_fim),'Último dia');
 const vespera=load('2026-10-14T03:00:01Z');
 assert.equal(vespera.ctx.obterTempoRestanteDesafioPainel_(baseItem.periodo_fim),'1 dia');
});
ok('datas brasileiras DD/MM/AAAA respeitadas',()=>{
 const d={...baseItem,periodo_fim:'15/10/2026'};
 assert.equal(after.ctx.prazoIndividualEncerradoVisual_(d),true);
});
ok('desafio mensal não herda o encerramento visual PRAZO_DIAS',()=>{
 const d={...baseItem,tipo_meta:'KM',prazo_dias:0,periodo_fim:'2026-09-30'};
 assert.equal(after.ctx.prazoIndividualEncerradoVisual_(d),false);
});
ok('conclusão administrativa não é recategorizada',()=>{
 const d={...baseItem,status_usuario_desafio:'CONCLUIDO',status_apuracao:'CONCLUIDO'};
 assert.equal(after.ctx.prazoIndividualEncerradoVisual_(d),false);
 assert.equal(after.ctx.classificarDesafioV2_(d).grupo,'concluidos');
});
ok('status CANCELADO ou EXPIRADO não vira etiqueta de prazo',()=>{
 for(const s of ['CANCELADO','EXPIRADO','DESISTENTE']){
  assert.equal(after.ctx.prazoIndividualEncerradoVisual_({...baseItem,status_usuario_desafio:s}),false);
 }
});
ok('período inválido não produz marcação enganosa',()=>{
 assert.equal(after.ctx.prazoIndividualEncerradoVisual_({...baseItem,periodo_fim:''}),false);
});
ok('desafio encerrado mantém grupo e acesso, sem mudar status em dados',()=>{
 assert.equal(after.ctx.classificarDesafioV2_(baseItem).grupo,'andamento');
 const original=JSON.stringify(baseItem);
 after.ctx.renderDesafiosV2();
 assert.equal(after.ids['desafios-v2-andamento-title'].textContent,'Em acompanhamento');
 assert.equal(Array.from(after.captures.groups.andamento).join(','),'FICTICIA_45');
 assert.equal(JSON.stringify(baseItem),original);
});
ok('grupo volta ao título normal quando só há desafio no prazo',()=>{
 before.ctx.renderDesafiosV2();
 assert.equal(before.ids['desafios-v2-andamento-title'].textContent,'Em andamento');
});
ok('badge aparece no card com prazo encerrado e não no ainda válido',()=>{
 const a=after.ctx.buildDesafioCardV2_(baseItem,'andamento');
 const b=before.ctx.buildDesafioCardV2_(baseItem,'andamento');
 assert.match(a.innerHTML,/desafio-v2-prazo-badge">Prazo encerrado/);
 assert.doesNotMatch(b.innerHTML,/Prazo encerrado/);
});
ok('badge não aparece no card Concluído',()=>{
 const d={...baseItem,status_usuario_desafio:'CONCLUIDO',status_apuracao:'CONCLUIDO'};
 const a=after.ctx.buildDesafioCardV2_(d,'concluidos');
 assert.doesNotMatch(a.innerHTML,/Prazo encerrado/);
});
ok('painel e detalhe têm avisos ocultos por padrão e são controlados por JS',()=>{
 for(const id of ['painel-prazo-status','desafio-detalhe-prazo-status']){
  assert.match(html,new RegExp('id="'+id+'"[^>]*hidden'));
  assert.ok(script.includes("getElementById('"+id+"')"));
 }
 assert.match(script,/Seus desafios em acompanhamento/);
 assert.match(script,/Prazo encerrado' : 'Desafio em andamento'/);
});
ok('etiqueta compacta com ocultação explícita, sem alterar outros módulos',()=>{
 assert.match(css,/\.desafio-v2-prazo-badge\[hidden\]\s*\{\s*display:\s*none\s*!important/);
 assert.match(css,/font-size: 12px/);
 assert.doesNotMatch(script,/setValue\(|SpreadsheetApp|clasp update-deployment/);
});
console.log('UX_PRAZO_ENCERRADO='+total+'/'+total+' PASS; runtime sem deploy, backend idêntico e status preservado');
