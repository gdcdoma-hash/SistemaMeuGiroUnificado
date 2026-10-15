'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

// Regressão da exceção antiga da PR #307: link já emitido NUNCA
// substitui 100% dos KM nem CONCLUIDO manual do administrador.
const root=path.resolve(__dirname,'..','Meu Giro');
const backend=fs.readFileSync(path.join(root,'CertificadoService.gs'),'utf8');
const uiSource=fs.readFileSync(path.join(root,'Script.html'),'utf8');

function extract(text,fn) {
  const start=text.indexOf('function '+fn+'(');
  assert.ok(start>=0,'missing function '+fn);
  const end=text.indexOf('\nfunction ',start+11);
  return text.slice(start,end<0?text.length:end);
}

const uiFns=['isDesafioElegivelCertificado_','getMensagemCertificado_']
  .map(n=>extract(uiSource,n)).join('\n');
const ui=new Function('toFiniteNumber',uiFns+
  '\nreturn {eligible:isDesafioElegivelCertificado_,message:getMensagemCertificado_};')(
    (v,fallback=0)=>{
      const n=Number(String(v==null?'':v).replace(',','.'));
      return Number.isFinite(n)?n:fallback;
    });

const backendFn=extract(backend,'gerarOuObterCertificadoDesafio');
function check({km=100,meta=100,admin='CONCLUIDO',status='CONCLUIDO',url='',reprovado=false}={}) {
  let generated=0,read=0;
  const deps={
    certificadoBuscarContextoDesafio_:()=>({
      ok:true,status_usuario_desafio:admin,status_apuracao:status,
      meta_atingida:meta>0 && km>=meta,
      link_certificado_existente:url,
      status_validacao_certificado:reprovado?'REPROVADO':''
    }),
    certificadoLerLinkPlanilha_:()=>{read++;return url;},
    certLinkValido_:x=>/^https?:\/\/\S+/i.test(String(x||'')),
    certificadoMontarExtrasImagem_:()=>({}),
    montarUrlDownloadPdfCertificado_:x=>x,
    gerarCertificadoDesafio_:()=>{generated++;return {ok:true,url:'https://example.test/new.pdf',pdfDownloadUrl:'https://example.test/new.pdf'};}
  };
  const fn=new Function(...Object.keys(deps),backendFn+
    '\nreturn gerarOuObterCertificadoDesafio;')(...Object.values(deps));
  const response=fn({});
  return {response,generated,read};
}
const existing='https://example.test/issued.pdf';

test('meta inferior a 100% + EM_ANDAMENTO = nenhum certificado',()=>{
  assert.equal(check({km:99,admin:'EM_ANDAMENTO'}).response.code,'META_NAO_ATINGIDA_CERTIFICADO');
  assert.equal(ui.eligible({meta_km:100,distancia_realizada:99,status_apuracao:'STATUS_EM_ANALISE',status_usuario_desafio:'EM_ANDAMENTO'}),false);
});

test('100% da meta + EM_ANDAMENTO = aguardando admin',()=>{
  assert.equal(check({km:100,admin:'EM_ANDAMENTO'}).response.code,'CONCLUSAO_ADMINISTRATIVA_PENDENTE');
  assert.equal(ui.eligible({meta_km:100,distancia_realizada:100,status_apuracao:'CONCLUIDO',status_usuario_desafio:'EM_ANDAMENTO'}),false);
});

test('menos de 100% + status manual CONCLUIDO = sem certificado',()=>{
  const a=check({km:99,admin:'CONCLUIDO'});
  assert.equal(a.response.code,'META_NAO_ATINGIDA_CERTIFICADO');
  assert.equal(a.generated,0);
});

test('100% e CONCLUIDO manual = certificado liberado',()=>{
  const a=check({km:100,admin:'CONCLUIDO'});
  assert.equal(a.response.ok,true);
  assert.equal(a.generated,1);
  assert.equal(ui.eligible({meta_km:100,distancia_realizada:100,status_apuracao:'CONCLUIDO',status_usuario_desafio:'CONCLUIDO'}),true);
});

test('mais de 100% e CONCLUIDO manual também é válido',()=>{
  assert.equal(check({km:101,admin:'CONCLUIDO'}).response.ok,true);
});

test('LINK_CERTIFICADO anterior não supera KM insuficientes',()=>{
  const a=check({km:99,admin:'CONCLUIDO',url:existing});
  assert.equal(a.response.code,'META_NAO_ATINGIDA_CERTIFICADO');
  assert.equal(a.generated,0);
  assert.equal(a.read,0);
  assert.equal(ui.eligible({meta_km:100,distancia_realizada:99,status_apuracao:'STATUS_EM_ANALISE',status_usuario_desafio:'CONCLUIDO',certificado_emitido:true}),false);
});

test('LINK_CERTIFICADO anterior não supera status do admin pendente',()=>{
  const a=check({km:100,admin:'EM_ANDAMENTO',url:existing});
  assert.equal(a.response.code,'CONCLUSAO_ADMINISTRATIVA_PENDENTE');
  assert.equal(a.generated,0);
});

test('sem meta válida não liberar, mesmo com apuração antiga CONCLUIDO',()=>{
  const a=check({km:100,meta:0,admin:'CONCLUIDO',status:'CONCLUIDO'});
  assert.equal(a.response.code,'META_NAO_ATINGIDA_CERTIFICADO');
});

test('status de apuração não concluído bloqueia mesmo com KM numéricos 100%',()=>{
  const a=check({km:100,admin:'CONCLUIDO',status:'STATUS_EM_ANALISE'});
  assert.equal(a.response.code,'DESAFIO_NAO_ELEGIVEL_CERTIFICADO');
  assert.equal(ui.eligible({meta_km:100,distancia_realizada:100,status_apuracao:'STATUS_EM_ANALISE',status_usuario_desafio:'CONCLUIDO'}),false);
});

test('certificado existente é reutilizado APENAS com ambos os requisitos presentes',()=>{
  const a=check({km:100,admin:'CONCLUIDO',status:'CONCLUIDO',url:existing});
  assert.equal(a.response.ok,true);
  assert.equal(a.response.reused,true);
  assert.equal(a.response.url,existing);
  assert.equal(a.generated,0);
});

test('CANCELADO nunca libera certificado',()=>{
  const a=check({km:100,admin:'CANCELADO',status:'CONCLUIDO',url:existing});
  assert.equal(a.response.code,'INSCRICAO_CANCELADA');
  assert.equal(a.generated,0);
});

test('mensagem da interface exige KM 100% e validação manual',()=>{
  const msg=ui.message({meta_km:100,distancia_realizada:99,status_usuario_desafio:'CONCLUIDO',status_apuracao:'CONCLUIDO'});
  assert.match(msg.texto,/100%/);
});
