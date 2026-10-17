'use strict';
// Diagnóstico SOMENTE LEITURA da versão realmente publicada @10.
// Este teste reproduz a divergência visual sem corrigir status, KM ou dados reais.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const cp = require('node:child_process');
const vm = require('node:vm');
const BASE = '3c2bb257ad034bb80174472e723242d642d697ee';
const frontendPath = 'Meu Giro/Script.html';
const backendPath = 'Meu Giro/PainelService.gs';
const frontend = fs.readFileSync(frontendPath, 'utf8');
const backend = fs.readFileSync(backendPath, 'utf8');
for(const [path,current] of [[frontendPath,frontend],[backendPath,backend]]){
  const published = cp.execFileSync('git',['show',BASE+':'+path],{encoding:'utf8'});
  assert.equal(current,published,'Código runtime modificado na branch de auditoria: '+path);
}
function section(source,start,next){
  const a=source.indexOf('function '+start+'(');
  const b=source.indexOf('function '+next+'(',a+1);
  assert.ok(a>=0 && b>a,'Não foi possível encontrar '+start);
  return source.slice(a,b);
}
const clientParts = [
  ['parseIsoDateLocal','calcularRitmoDesafio'],
  ['classificarDesafioV2_','renderGrupoDesafiosV2_'],
  ['obterTempoRestanteDesafioPainel_','atualizarResumoPrazoDesafioPainel_'],
  ['isDesafioAtivo_','isStatusFinalFoco_'],
  ['isStatusFinalFoco_','isDesafioHistoricoElegivel_']
].map(x=>section(frontend,...x)).join('\n');
const serverParts = [
  ['painelMG_isDesafioAtivoParaFoco_','painelMG_isStatusFinalFoco_'],
  ['painelMG_isStatusFinalFoco_','painelMG_compareFocoDesafiosAtivos_']
].map(x=>section(backend,...x)).join('\n');
let count=0;
function clock(iso){
  const frozen=new Date(iso+'T12:00:00-03:00').getTime();
  class MockDate extends Date {
    constructor(...args){super(...(args.length?args:[frozen]));}
    static now(){return frozen;}
  }
  const sandbox={Date:MockDate, console};
  sandbox.painelMG_norm_=value=>String(value||'').trim();
  sandbox.painelMG_normalizarStatus_=value=>String(value||'').trim().toUpperCase();
  sandbox.painelMG_normalizarDataISO_=value=>String(value||'').trim();
  vm.createContext(sandbox);
  vm.runInContext(clientParts+'\n'+serverParts,sandbox);
  return sandbox;
}
const inscricao={
  id_inscricao:'INS_FICTICIA_1000',id_desafio:'1000',
  tipo_meta:'PRAZO_DIAS',prazo_dias:45,
  periodo_inicio:'2026-09-01',periodo_fim:'2026-10-15',
  status_usuario_desafio:'EM_ANDAMENTO',status_apuracao:'STATUS_EM_ANALISE',
  status_desafio:'ATIVO',status_lista_desafios:'ATIVO',
  meta_km:1000,distancia_realizada:18
};
const encerramento=clock('2026-10-16');
const ultimoDia=clock('2026-10-15');
assert.equal(ultimoDia.obterTempoRestanteDesafioPainel_(inscricao.periodo_fim),'Último dia');
assert.equal(ultimoDia.classificarDesafioV2_(inscricao).grupo,'andamento');
count++;
assert.equal(encerramento.obterTempoRestanteDesafioPainel_(inscricao.periodo_fim),'Encerrado');
assert.equal(encerramento.classificarDesafioV2_(inscricao).grupo,'andamento',
  'Divergência conhecida: prazo vencido ainda classificado como andamento');
count++;
assert.equal(encerramento.isDesafioAtivo_(inscricao),true);
assert.equal(encerramento.painelMG_isDesafioAtivoParaFoco_(inscricao),true);
count++;
const outubro={...inscricao,id_inscricao:'INS_FICTICIA_OUT',tipo_meta:'KM',prazo_dias:0,
periodo_inicio:'2026-10-01',periodo_fim:'2026-10-31',meta_km:150};
assert.equal(encerramento.classificarDesafioV2_(outubro).grupo,'andamento');
assert.equal(encerramento.isDesafioAtivo_(outubro),true);
count++;
const concluido={...inscricao,status_usuario_desafio:'CONCLUIDO',status_apuracao:'CONCLUIDO'};
assert.equal(encerramento.classificarDesafioV2_(concluido).grupo,'concluidos',
  'Conclusão administrativa deve permanecer separada do vencimento temporal');
assert.equal(encerramento.isDesafioAtivo_(concluido),false);
count++;
assert.match(frontend,/var grupos = \{ andamento: \[\], concluidos: \[\] \}/,
  'A lista V2 não tem uma terceira seção para prazo encerrado');
assert.match(frontend,/if \(classificacao.grupo === 'andamento' \|\| classificacao.grupo === 'concluidos'\)/);
assert.equal(inscricao.distancia_realizada,18);
assert.equal(inscricao.status_usuario_desafio,'EM_ANDAMENTO');
count++;
console.log('DIAGNOSTICO_PRAZO_UI_PROD10='+count+'/6 PASS');
console.log('RISCO_REPRODUZIDO=tempo Encerrado mas grupos/foco EM_ANDAMENTO; sem escrita/sem mudança de código.');
