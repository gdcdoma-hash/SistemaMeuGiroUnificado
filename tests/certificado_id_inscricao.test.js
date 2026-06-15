const assert = require('assert');
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const CERT = fs.readFileSync(path.join(ROOT, 'Meu Giro/CertificadoService.gs'), 'utf8');
const ADMIN = fs.readFileSync(path.join(ROOT, 'Meu Giro/AdminCertificadoService.gs'), 'utf8');

function normalize(value) {
  return String(value == null ? '' : value).trim();
}

function headerMap(headers) {
  const out = {};
  headers.forEach((value, index) => { out[normalize(value).toLowerCase()] = index; });
  return out;
}

function optionalIndex(map, aliases) {
  for (const alias of aliases) {
    const key = normalize(alias).toLowerCase();
    if (Object.prototype.hasOwnProperty.call(map, key)) return map[key];
  }
  return -1;
}

class FakeRange {
  constructor(sheet, row, col, numRows = 1, numCols = 1) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  getValue() { return this.sheet.values[this.row - 1][this.col - 1]; }
  setValue(value) { this.sheet.values[this.row - 1][this.col - 1] = value; return this; }
  getValues() {
    return Array.from({ length: this.numRows }, (_, r) =>
      Array.from({ length: this.numCols }, (_, c) => this.sheet.values[this.row - 1 + r][this.col - 1 + c])
    );
  }
}

class FakeSheet {
  constructor(values) { this.values = values.map(row => row.slice()); }
  getDataRange() { return { getValues: () => this.values.map(row => row.slice()) }; }
  getRange(row, col, numRows, numCols) { return new FakeRange(this, row, col, numRows, numCols); }
  getLastColumn() { return Math.max(...this.values.map(row => row.length)); }
}

function baseSandbox(sheet, summaries) {
  const spreadsheet = { getSheetByName: name => name === 'dgmbDesafios' ? sheet : null };
  return {
    console,
    SHEETS: { DESAFIO: 'dgmbDesafios', MEU_GIRO_RESUMO: 'MEU_GIRO_RESUMO' },
    normalizeText_: normalize,
    buildHeaderMap_: headerMap,
    getOptionalColumnIndex_: optionalIndex,
    getRequiredColumnIndex_: (map, aliases) => {
      const idx = optionalIndex(map, aliases);
      if (idx < 0) throw new Error(`Missing column ${aliases[0]}`);
      return idx;
    },
    getIdDesafioColumnIndex_: map => optionalIndex(map, ['id_desafio']),
    obterIdDesafioRegistro_: (row, idx) => normalize(row[idx]),
    getSpreadsheet_: () => spreadsheet,
    atualizarMeuGiroResumo_: () => summaries || [],
    SpreadsheetApp: { flush() {} },
    Logger: { log() {} },
    parseLocalizedNumber_: value => Number(String(value || 0).replace(',', '.')),
    firstFilledValue_: (row, aliases) => {
      for (const alias of aliases) if (normalize(row[alias])) return row[alias];
      return '';
    },
    getAllObjects_: () => [],
    Date,
  };
}

function loadCertificate(sheet, summaries) {
  const sandbox = baseSandbox(sheet, summaries);
  vm.createContext(sandbox);
  vm.runInContext(CERT, sandbox, { filename: 'CertificadoService.gs' });
  return sandbox;
}

function loadAdmin(sheet) {
  const sandbox = baseSandbox(sheet, []);
  vm.createContext(sandbox);
  vm.runInContext(ADMIN, sandbox, { filename: 'AdminCertificadoService.gs' });
  return sandbox;
}

const headers = [
  'id_inscricao', 'id_dgmb', 'id_desafio', 'id_item_estoque', 'status_apuracao',
  'status_usuario_desafio', 'status_validacao_certificado', 'print_strava_certificado',
  'link_print_strava', 'data_envio_print_strava', 'data_aprovacao_certificado',
  'obs_validacao_certificado', 'link_certificado'
];

function modernRow(inscricao, link = '') {
  return [inscricao, '10', 'D1', 'ITEM1', 'CONCLUIDO', 'CONCLUIDO', 'APROVADO', '', '', '', '', '', link];
}

// Cenário 1: a validação administrativa não participa da elegibilidade do certificado.
[
  ['PENDENTE', true],
  ['EM_ANALISE', true],
  ['APROVADO', true],
  ['REPROVADO', true],
  ['', true]
].forEach(([statusValidacao, esperado], index) => {
  const row = modernRow(`INS-1-${index}`);
  row[6] = statusValidacao;
  const sheet = new FakeSheet([headers, row]);
  const ctx = loadCertificate(sheet, [{ id_inscricao: row[0], id_dgmb: '10', status_apuracao: 'CONCLUIDO' }]);
  const result = ctx.certificadoBuscarContextoDesafio_({ id_inscricao: row[0], id_dgmb: '10', id_desafio: 'D1', id_item_estoque: 'ITEM1' });
  assert.equal(result.ok, true);
  assert.equal(result.id_inscricao, row[0]);
  assert.equal(result.status_validacao_certificado, statusValidacao);
  assert.equal(result.desafio_elegivel, esperado);
});

// Status de apuração ou do usuário ainda impedem a elegibilidade, independentemente da validação administrativa.
[
  ['ATIVO', 'CONCLUIDO'],
  ['CONCLUIDO', 'EM_ANDAMENTO']
].forEach(([statusApuracao, statusUsuario], index) => {
  const row = modernRow(`INS-NE-${index}`);
  row[5] = statusUsuario;
  row[6] = 'APROVADO';
  const sheet = new FakeSheet([headers, row]);
  const ctx = loadCertificate(sheet, [{ id_inscricao: row[0], id_dgmb: '10', status_apuracao: statusApuracao }]);
  const result = ctx.certificadoBuscarContextoDesafio_({ id_inscricao: row[0], id_dgmb: '10', id_desafio: 'D1', id_item_estoque: 'ITEM1' });
  assert.equal(result.desafio_elegivel, false);
});

// Cenário 2: duas inscrições iguais na chave legada selecionam, aprovam e gravam o link da inscrição indicada.
{
  const rows = [headers, modernRow('INS-A'), modernRow('INS-B')];
  rows[1][6] = 'PENDENTE';
  rows[2][6] = 'PENDENTE';
  const sheet = new FakeSheet(rows);
  const summaries = [
    { id_inscricao: 'INS-A', id_dgmb: '10', meta_km: 50, distancia_realizada: 10, percentual_concluido: 20, periodo_inicio: '2026-01-01', periodo_fim: '2026-01-31', status_apuracao: 'ATIVO' },
    { id_inscricao: 'INS-B', id_dgmb: '10', meta_km: 100, distancia_realizada: 100, percentual_concluido: 100, periodo_inicio: '2026-02-01', periodo_fim: '2026-02-28', status_apuracao: 'CONCLUIDO' }
  ];
  const cert = loadCertificate(sheet, summaries);
  const selected = cert.certificadoBuscarContextoDesafio_({ id_inscricao: 'INS-B', id_dgmb: '10', id_desafio: 'D1', id_item_estoque: 'ITEM1' });
  assert.equal(selected.rowNumber, 3);
  assert.equal(selected.status_apuracao, 'CONCLUIDO');
  const visualSource = cert.certificadoBuscarResumoDesafio_(selected);
  assert.equal(visualSource.meta_km, 100);
  assert.equal(visualSource.distancia_realizada, 100);
  assert.equal(visualSource.percentual_concluido, 100);
  assert.equal(visualSource.periodo_inicio, '2026-02-01');
  assert.equal(cert.certificadoSalvarLinkPlanilha_(selected, 'https://drive.google.com/file/d/1234567890123456789012345/view').ok, true);
  assert.equal(sheet.values[1][12], '');
  assert.match(sheet.values[2][12], /^https:/);

  const admin = loadAdmin(sheet);
  const approved = admin.atualizarStatusValidacaoCertificadoAdmin({
    admin_id_dgmb: '1133', id_inscricao: 'INS-B', id_dgmb: '10', id_desafio: 'D1',
    id_item_estoque: 'ITEM1', row_number: 2, novo_status: 'APROVADO'
  });
  assert.equal(approved.ok, true);
  assert.equal(approved.data.row_number, 3);
  assert.equal(sheet.values[1][6], 'PENDENTE');
  assert.equal(sheet.values[2][6], 'APROVADO');
}

// Cenário 3: registro legado sem coluna ID_INSCRICAO usa a chave legada exata.
{
  const legacyHeaders = headers.slice(1);
  const legacyRow = modernRow('').slice(1);
  const sheet = new FakeSheet([legacyHeaders, legacyRow]);
  const ctx = loadCertificate(sheet, [{ id_desafio: 'D1', id_item_estoque: 'ITEM1', status_apuracao: 'CONCLUIDO' }]);
  const result = ctx.certificadoBuscarContextoDesafio_({ id_dgmb: '10', id_desafio: 'D1', id_item_estoque: 'ITEM1' });
  assert.equal(result.ok, true);
  assert.equal(result.id_inscricao, '');
  assert.equal(result.status_apuracao, 'CONCLUIDO');
}

// Cenário 4: fallback legado ambíguo não escolhe a primeira linha.
{
  const legacyHeaders = headers.slice(1);
  const legacyRow = modernRow('').slice(1);
  const sheet = new FakeSheet([legacyHeaders, legacyRow, legacyRow]);
  const ctx = loadCertificate(sheet, []);
  const result = ctx.certificadoBuscarContextoDesafio_({ id_dgmb: '10', id_desafio: 'D1', id_item_estoque: 'ITEM1' });
  assert.equal(result.ok, false);
  assert.equal(result.code, 'CERTIFICADO_INSCRICAO_AMBIGUA');
}

// Cenário 5: Link_Certificado existente na inscrição correta é reutilizado sem gerar novo PDF.
{
  const existing = 'https://drive.google.com/file/d/1234567890123456789012345/view';
  const sheet = new FakeSheet([headers, modernRow('INS-5', existing)]);
  const ctx = loadCertificate(sheet, [{ id_inscricao: 'INS-5', id_dgmb: '10', status_apuracao: 'CONCLUIDO' }]);
  let generated = false;
  ctx.certificadoMontarExtrasImagem_ = () => ({});
  ctx.gerarCertificadoDesafio_ = () => { generated = true; return { ok: true }; };
  const result = ctx.gerarOuObterCertificadoDesafio({ id_inscricao: 'INS-5', id_dgmb: '10', id_desafio: 'D1', id_item_estoque: 'ITEM1' });
  assert.equal(result.ok, true);
  assert.equal(result.reused, true);
  assert.equal(result.url, existing);
  assert.equal(generated, false);
}

// Nomes novos incluem ID_INSCRICAO; o sufixo só é adicionado quando ele existe, preservando o legado.
assert.match(CERT, /if \(ctx\.id_inscricao\) nomeArquivoPartes\.push\(ctx\.id_inscricao\)/);
const imageSource = fs.readFileSync(path.join(ROOT, 'Meu Giro/CertificadoImagemService.gs'), 'utf8');
assert.match(imageSource, /if \(ctx\.id_inscricao\) nomeArquivoPartes\.push\(ctx\.id_inscricao\)/);

console.log('OK: regra de elegibilidade e cenários de certificado por ID_INSCRICAO validados.');
