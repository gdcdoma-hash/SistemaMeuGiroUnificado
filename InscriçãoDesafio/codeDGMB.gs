/**
 * SistemaDGMBv3 - Versão USUARIOdgmb1
 */
const SPREADSHEET_ID = '1wjbEuF5ZHUeLfKUpbNvgB2yuXWqHC20H74JXu4SnnsQ';
const ABA_PESSOAL = 'DadosPessoais';
const ABA_LISTA = 'ListaDesafios';
const PASTA_SISTEMA_ID = '1bDrfsDoim-TN9i5lWJnfCPRXAV1OIKQP';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('formDGMB')
    .setTitle('Desafio Giro Motos Bikes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function criptografar(texto) { 
  if (!texto) return "";
  const limpo = texto.toString().replace(/\D/g, '');
  return Utilities.base64Encode(limpo); 
}

function verificarStatusSistema() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ABA_LISTA);
  return sheet.getRange("G2").getValue(); // Retorna "Aberto" ou "Fechado"
}

function processarUpload(e) {
  try {
    const pastaSistema = DriveApp.getFolderById(PASTA_SISTEMA_ID);
    let pastaDesafio;
    const pastasDesafioEx = pastaSistema.getFoldersByName(e.abaDesafio);
    pastaDesafio = pastasDesafioEx.hasNext() ? pastasDesafioEx.next() : pastaSistema.createFolder(e.abaDesafio);
    const nomeSub = (e.tipo === 'avatar') ? 'dgmbAvatares' : 'dgmbComprovantes';
    let subFolder;
    const subPastasEx = pastaDesafio.getFoldersByName(nomeSub);
    subFolder = subPastasEx.hasNext() ? subPastasEx.next() : pastaDesafio.createFolder(nomeSub);
    const bytes = Utilities.base64Decode(e.base64.split(",")[1]);
    const blob = Utilities.newBlob(bytes, e.mimeType, e.nomeArquivo);
    const arquivo = subFolder.createFile(blob);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const urlArquivo = arquivo.getUrl();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(e.abaDesafio);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1].toString().trim() === e.id_dgmb.toString().trim()) {
        const col = (e.tipo === 'avatar') ? 7 : 9;
        sheet.getRange(i + 1, col).setValue(urlArquivo);
        sheet.getRange(i + 1, col + 1).setValue('Enviado');
        return { status: 'OK' };
      }
    }
    return { status: 'ERRO', msg: 'ID não localizado.' };
  } catch (err) { return { status: 'ERRO', msg: err.toString() }; }
}

function verificarCPF(cpf) {
  const clean = cpf.replace(/\D/g, '');
  if (!isValidCPF(clean)) return { status: 'INVALIDO' };
  
  const statusGeral = verificarStatusSistema();
  const cpfCripto = criptografar(clean);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rowsP = ss.getSheetByName(ABA_PESSOAL).getDataRange().getValues();
  
  let user = null;
  for (let i = 1; i < rowsP.length; i++) {
    if (rowsP[i][3] === cpfCripto) {
      user = { status: 'CADASTRO_EXISTENTE', nome: rowsP[i][4], cpf: clean, id_dgmb: rowsP[i][1] };
      break;
    }
  }

  // Se não existe cadastro e o sistema está FECHADO
  if (!user && statusGeral === "Fechado") {
    return { status: 'SISTEMA_FECHADO' };
  }

  // Se não existe cadastro e está ABERTO
  if (!user) {
    return { status: 'NOVO_USUARIO', cpf: clean };
  }

  // Se existe cadastro, verificar se já está inscrito em algum desafio ativo
  const listaSheet = ss.getSheetByName(ABA_LISTA);
  const desafiosAtivos = listaSheet.getDataRange().getDisplayValues().slice(1).filter(l => l[3] === 'Ativo');
  
  for (let d of desafiosAtivos) {
    const s = ss.getSheetByName(d[1]); 
    if (!s) continue;
    const rowsD = s.getDataRange().getValues();
    for (let j = 1; j < rowsD.length; j++) {
      if (rowsD[j][1].toString() === user.id_dgmb.toString()) {
        return { 
          status: 'INSCRICAO_REALIZADA', 
          nome: rowsD[j][2], 
          id_dgmb: user.id_dgmb, 
          cpf: clean, 
          abaDesafio: d[1], 
          nomeExibicao: d[2], 
          meta: rowsD[j][3], 
          fotoStatus: rowsD[j][7], 
          pagtoStatus: rowsD[j][9], 
          fraseIA: rowsD[j][12] || "O vento no rosto é a recompensa da liberdade!", 
          periodo: d[5],
          totalAtivos: desafiosAtivos.length,
          sistemaStatus: statusGeral
        };
      }
    }
  }

  // Se tem cadastro mas não tem inscrição em desafio ativo
  if (statusGeral === "Fechado") {
    return { status: 'SISTEMA_FECHADO' };
  }
  
  return user;
}

function isValidCPF(cpf) {
  const c = cpf.replace(/\D/g, '');
  if (c.length !== 11 || !!c.match(/(\d)\1{10}/)) return false;
  const r = (n) => (c.split('').slice(0, n-12).reduce((s, e, i) => (s + e * (n - i)), 0) * 10) % 11 % 10;
  return r(10) == c[9] && r(11) == c[10];
}

function buscarCEP(cep) {
  const clean = cep.replace(/\D/g, '');
  if (clean.length !== 8) return null;
  try {
    const geocoder = Maps.newGeocoder().setRegion('br').geocode(clean);
    if (geocoder.status === 'OK' && geocoder.results.length > 0) {
      const result = geocoder.results[0];
      let cidade = ""; let estado = "";
      result.address_components.forEach(comp => {
        if (comp.types.includes("administrative_area_level_2")) cidade = comp.long_name;
        if (comp.types.includes("administrative_area_level_1")) estado = comp.short_name;
      });
      return { localidade: cidade, uf: estado };
    }
    return null;
  } catch (e) { return null; }
}

function listarDesafiosAtivos() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ABA_LISTA);
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  return data.filter(l => l[3] === 'Ativo').map(l => {
    return { aba: l[1], nome: l[2], prazo: l[5] };
  });
}

function gravarDadosPessoais(d) {
  if (verificarStatusSistema() === "Fechado") return { status: 'ERRO' };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ABA_PESSOAL);
  const data = sheet.getDataRange().getValues();
  let ultimoId = 1132;
  if (data.length > 1) { ultimoId = Math.max(...data.slice(1).map(r => parseInt(r[1]) || 0)); }
  const novoId = ultimoId + 1;
  sheet.appendRow([new Date(), novoId, "", criptografar(d.cpf), d.nome, d.cep, d.cidadeuf, d.whatsapp, d.nascimento]);
  return { status: 'OK', id_dgmb: novoId };
}

function gravarInscricaoDesafio(d) {
  if (verificarStatusSistema() === "Fechado") return { status: 'ERRO' };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const s = ss.getSheetByName(d.abaDesafio);
  const novaLinha = [new Date(), d.id_dgmb, d.avatar, d.distancia, d.obs, 'Sim', '', 'Pendente', '', 'Pendente', '', '', "Pedalar é superar limites!"];
  s.appendRow(novaLinha);
  const listaSheet = ss.getSheetByName(ABA_LISTA);
  const listaAtivos = listaSheet.getDataRange().getDisplayValues().slice(1).filter(l => l[3] === 'Ativo');
  let nEx = d.abaDesafio; let per = "";
  for(let i=0; i<listaAtivos.length; i++) { 
    if(listaAtivos[i][1] === d.abaDesafio) { nEx = listaAtivos[i][2]; per = listaAtivos[i][5]; break; } 
  }
  return { status: 'OK', nome: d.avatar, id_dgmb: d.id_dgmb, meta: d.distancia, abaDesafio: d.abaDesafio, nomeExibicao: nEx, cpf: d.cpf, pagtoStatus: 'Pendente', fotoStatus: 'Pendente', fraseIA: "Sua jornada rumo à meta começou!", periodo: per, totalAtivos: listaAtivos.length, jaExistia: false };
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }