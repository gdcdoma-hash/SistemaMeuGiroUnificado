'use strict';
const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');

const root=path.resolve(__dirname,'..','Meu Giro');
const backend=fs.readFileSync(path.join(root,'CertificadoService.gs'),'utf8');
const frontend=fs.readFileSync(path.join(root,'Script.html'),'utf8');
const utils=fs.readFileSync(path.join(root,'Utils.gs'),'utf8');
function extract(source,name) {
  const start=source.indexOf('function '+name+'(');
  assert.ok(start>=0,'missing '+name);
  const end=source.indexOf('\nfunction ',start+10);
  return source.slice(start,end<0?source.length:end);
}

const backendCode=['gerarOuObterCertificadoDesafio','certificadoRespostaLinkExistente_']
  .map(n=>extract(backend,n)).join('\n');
const uiCode=['isDesafioElegivelCertificado_','getMensagemCertificado_']
  .map(n=>extract(frontend,n)).join('\n');
const ui=new Function(uiCode+'\nreturn {eligible:isDesafioElegivelCertificado_,message:getMensagemCertificado_};')();

function scenario({admin='CONCLUIDO',status='CONCLUIDO',url='',rowValid=true,validacao=''}={}) {
  let writes=0,recalculations=0,generated=0,reusedCalls=0,read=0;
  const ctx={
    ok:true, sheet_name:'dgmbDesafios', rowNumber:100,idx_link_certificado:22,
    id_inscricao:'fake-A', id_dgmb:'fake-user',
    status_usuario_desafio:admin,status_apuracao:status,
    status_validacao_certificado:validacao,
    link_certificado_existente:url
  };
  const fakeSheet={};
  const deps={
    certificadoBuscarContextoDesafio_:()=>ctx,
    getSpreadsheet_:()=>({getSheetByName:()=>fakeSheet}),
    certificadoRevalidarLinhaInscricao_:()=>({ok:rowValid}),
    certificadoLerLinkPlanilha_:()=>{read++;return url;},
    certLinkValido_:link=>/^https?:\/\/\S+$/i.test(String(link||'')),
    certificadoMontarExtrasImagem_:()=>({}),
    montarUrlDownloadPdfCertificado_:link=>link,
    gerarCertificadoDesafio_:()=>{
      generated++;
      return {ok:true,url:'https://drive.google.com/new.pdf',pdfDownloadUrl:'https://drive.google.com/new.pdf'};
    },
    atualizarMeuGiroResumo_:()=>{recalculations++;throw Error('unexpected recomputation');},
    certificadoSalvarLinkPlanilha_:()=>{writes++;throw Error('unexpected write');}
  };
  const fn=new Function(...Object.keys(deps),backendCode+
    '\nreturn gerarOuObterCertificadoDesafio;')(...Object.values(deps));
  return {response:fn({id_dgmb:'fake-user',id_inscricao:'fake-A'}),
    state:{writes,recalculations,generated,reusedCalls,read}};
}

const validUrl='https://drive.google.com/file/d/issued/view';

test('certificado emitido e conclusão administrativa: re-download permitido apesar da divergência de KM',()=>{
  const result=scenario({url:validUrl,status:'STATUS_EM_ANALISE'});
  assert.equal(result.response.ok,true);
  assert.equal(result.response.reused,true);
  assert.equal(result.response.url,validUrl);
  assert.equal(result.state.generated,0);
  assert.equal(result.state.writes,0);
  assert.equal(result.state.recalculations,0);
});

test('documento não emitido: não gerar certificado novo quando KM divergirem',()=>{
  const result=scenario({url:'',status:'STATUS_EM_ANALISE'});
  assert.equal(result.response.ok,false);
  assert.equal(result.response.code,'DESAFIO_NAO_ELEGIVEL_CERTIFICADO');
  assert.equal(result.state.generated,0);
});

test('conclusão administrativa não validada: link antigo não deve liberar documento',()=>{
  const result=scenario({url:validUrl,status:'STATUS_EM_ANALISE',admin:'EM_ANDAMENTO'});
  assert.equal(result.response.ok,false);
  assert.equal(result.state.generated,0);
});

test('inscrição cancelada continua bloqueada mesmo com link',()=>{
  const result=scenario({url:validUrl,status:'CONCLUIDO',admin:'CANCELADO'});
  assert.equal(result.response.code,'INSCRICAO_CANCELADA');
});

test('identidade alterada impede exceção de re-download de conclusão divergente',()=>{
  const result=scenario({url:validUrl,status:'STATUS_EM_ANALISE',rowValid:false});
  assert.equal(result.response.ok,false);
  assert.equal(result.state.generated,0);
});

test('link inválido não é prova de certificado emitido',()=>{
  const result=scenario({url:'nota interna',status:'STATUS_EM_ANALISE'});
  assert.equal(result.response.ok,false);
  assert.equal(result.state.generated,0);
});

test('com apuração e conclusão administrativas OK, emissão normal permanece funcional',()=>{
  const result=scenario({url:'',status:'CONCLUIDO'});
  assert.equal(result.response.ok,true);
  assert.equal(result.response.reused,false);
  assert.equal(result.state.generated,1);
});

test('com apuração e conclusão OK e link existente, reaproveita normalmente',()=>{
  const result=scenario({url:validUrl,status:'CONCLUIDO'});
  assert.equal(result.response.ok,true);
  assert.equal(result.response.reused,true);
  assert.equal(result.state.generated,0);
});

test('frontend: aviso de documento já emitido apenas para inscrição oficialmente concluída',()=>{
  const emitted={status_usuario_desafio:'CONCLUIDO',
    status_apuracao:'STATUS_EM_ANALISE',certificado_emitido:true};
  assert.equal(ui.eligible(emitted),true);
  assert.match(ui.message(emitted).texto,/já emitido/);
  assert.equal(ui.eligible({...emitted,certificado_emitido:false}),false);
  assert.equal(ui.eligible({...emitted,status_usuario_desafio:'EM_ANDAMENTO'}),false);
  assert.equal(ui.eligible({...emitted,status_usuario_desafio:'CANCELADO'}),false);
});

test('frontend: certificado inexistente preserva regra exigindo as duas conclusões',()=>{
  const pending={status_usuario_desafio:'CONCLUIDO',status_apuracao:'STATUS_EM_ANALISE'};
  assert.equal(ui.eligible(pending),false);
  assert.equal(ui.eligible({...pending,status_apuracao:'CONCLUIDO'}),true);
});

test('o backend diferencia a disponibilidade do documento do status percentual',()=>{
  assert.match(utils,/certificado_emitido: statusDgmbResumo\.certificado_emitido === true/);
  assert.match(utils,/certificado_emitido: vinculo\.certificado_emitido === true/);
  assert.match(utils,/idxLinkCertificado > -1/);
});


test('certificado explicitamente REPROVADO não é recuperado pela exceção',()=>{
  const result=scenario({url:validUrl,status:'STATUS_EM_ANALISE',validacao:'REPROVADO'});
  assert.equal(result.response.ok,false);
  assert.equal(result.response.code,'DESAFIO_NAO_ELEGIVEL_CERTIFICADO');
  assert.equal(result.state.generated,0);
});

test('campo certificado_emitido propagado apenas quando validação não foi reprovada',()=>{
  assert.match(utils,/idxStatusValidacaoCertificado < 0 \|\| normalizeText_\(row\[idxStatusValidacaoCertificado\]\)/);
  assert.match(utils,/certificado_emitido: vinculo\.certificado_emitido === true &&/);
  assert.match(utils,/certificado_emitido: statusDgmbResumo\.certificado_emitido === true &&/);
});
