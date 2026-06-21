const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');
const CERT = fs.readFileSync(path.join(ROOT, 'Meu Giro/CertificadoService.gs'), 'utf8');
const DEFAULT_TEMPLATE = 'template-padrao';
const HEADERS = ['ID_DESAFIO', 'ID_ITEM_ESTOQUE', 'DESCRICAO', 'ID_SLIDE_TEMPLATE', 'ATIVO'];
const HEADERS_SEM_DESCRICAO = ['ID_DESAFIO', 'ID_ITEM_ESTOQUE', 'ID_SLIDE_TEMPLATE', 'ATIVO'];

class FakeSheet {
  constructor(values) { this.values = values; }
  getDataRange() { return { getValues: () => this.values.map(row => row.slice()) }; }
}

function load(configValues) {
  const sheet = configValues === null ? null : new FakeSheet(configValues);
  const sandbox = {
    console,
    SHEETS: { CONFIG_CERTIFICADO_TEMPLATE: 'CONFIG_CERTIFICADO_TEMPLATE' },
    TEMPLATE_CERTIFICADO_SLIDES_ID_: DEFAULT_TEMPLATE,
    getSpreadsheet_: () => ({ getSheetByName: () => sheet }),
    Date
  };
  vm.createContext(sandbox);
  vm.runInContext(CERT, sandbox, { filename: 'CertificadoService.gs' });
  sandbox.TEMPLATE_CERTIFICADO_SLIDES_ID_ = DEFAULT_TEMPLATE;
  return sandbox;
}

function resolve(values, desafio = 'D1', item = 'ITEM1') {
  return load(values).certificadoResolverTemplateSlides_(desafio, item);
}

// 1. Aba ausente → padrão.
assert.equal(resolve(null).source, 'PADRAO');

// 2. Aba vazia → padrão.
assert.equal(resolve([]).source, 'PADRAO');

// 3. Cabeçalho incompleto → padrão.
assert.equal(resolve([HEADERS.slice(0, 4)]).source, 'PADRAO');

// Cabeçalho sem DESCRICAO continua válido.
assert.notEqual(resolve([HEADERS_SEM_DESCRICAO]).reason, 'CABECALHO_INCOMPLETO');

// Configurações por desafio e por item funcionam sem DESCRICAO.
assert.equal(resolve([HEADERS_SEM_DESCRICAO, ['D1', '', 'template-desafio', true]]).source, 'DESAFIO');
assert.equal(resolve([HEADERS_SEM_DESCRICAO, ['D1', 'ITEM1', 'template-item', true]]).source, 'ITEM');

// As quatro colunas funcionais continuam obrigatórias.
for (const obrigatorio of HEADERS_SEM_DESCRICAO) {
  const headersIncompletos = HEADERS_SEM_DESCRICAO.filter(header => header !== obrigatorio);
  assert.equal(resolve([headersIncompletos]).reason, 'CABECALHO_INCOMPLETO');
}

// 4. ID_DESAFIO vazio → padrão.
assert.equal(resolve([HEADERS], '  ').reason, 'ID_DESAFIO_VAZIO');

// 5. Configuração por desafio → usa DESAFIO.
assert.deepEqual(
  JSON.parse(JSON.stringify(resolve([HEADERS, ['D1', '', 'Geral', 'template-desafio', true]]))),
  { templateId: 'template-desafio', source: 'DESAFIO', fallback: false, reason: 'CONFIG_DESAFIO' }
);

// 6. Configuração por desafio + item → usa ITEM.
assert.equal(resolve([HEADERS, ['D1', 'ITEM1', 'Item', 'template-item', 'SIM']]).source, 'ITEM');

// 7. Item inexistente → cai para DESAFIO.
assert.equal(resolve([
  HEADERS,
  ['D1', '', 'Geral', 'template-desafio', true],
  ['D1', 'OUTRO', 'Outro', 'template-outro', true]
]).source, 'DESAFIO');

// 8. Nenhuma configuração → padrão.
assert.equal(resolve([HEADERS, ['D2', '', 'Outro desafio', 'template-d2', true]]).source, 'PADRAO');

// 9. Linha inativa → padrão.
assert.equal(resolve([HEADERS, ['D1', '', 'Inativo', 'template-desafio', false]]).source, 'PADRAO');

// 10. Template vazio → padrão.
assert.equal(resolve([HEADERS, ['D1', '', 'Sem template', '  ', true]]).source, 'PADRAO');

// 11. Duplicidade por desafio → padrão.
assert.equal(resolve([
  HEADERS,
  ['D1', '', 'Geral 1', 'template-1', true],
  ['D1', '', 'Geral 2', 'template-2', true]
]).reason, 'DUPLICIDADE_ATIVA_DESAFIO');

// 12. Duplicidade por desafio + item → padrão.
assert.equal(resolve([
  HEADERS,
  ['D1', 'ITEM1', 'Item 1', 'template-1', true],
  ['D1', 'ITEM1', 'Item 2', 'template-2', true]
]).reason, 'DUPLICIDADE_ATIVA_ITEM');

// 13. ATIVO TRUE/SIM/1 aceito.
for (const active of [true, 'TRUE', 'VERDADEIRO', 'SIM', '1', 1]) {
  assert.equal(resolve([HEADERS, ['D1', '', 'Ativo', `template-${active}`, active]]).source, 'DESAFIO');
}

// 14. ATIVO falso/texto estranho rejeitado.
for (const inactive of [false, 'FALSE', 'NAO', 'texto', 0, 2, '']) {
  assert.equal(resolve([HEADERS, ['D1', '', 'Inativo', 'template', inactive]]).source, 'PADRAO');
}

// 15. Espaços laterais nos IDs são normalizados.
assert.equal(resolve([HEADERS, ['  D1  ', '  ITEM1  ', 'Item', ' template-item ', true]]).templateId, 'template-item');

function generationSandbox(options = {}) {
  const calls = [];
  let generated = 0;
  const configuredFails = options.configuredFails !== false;
  const defaultFails = options.defaultFails === true;
  const folder = {
    getFilesByName: () => ({ hasNext: () => false }),
    createFile: () => ({
      setSharing() {},
      getUrl: () => 'https://drive.google.com/file/d/1234567890123456789012345/view',
      getId: () => '1234567890123456789012345'
    })
  };
  const sandbox = load([HEADERS, ['D1', 'ITEM1', 'Item', 'template-item', true]]);
  sandbox.certificadoBuscarDadosVisuais_ = () => ({});
  sandbox.certificadoGetOuCriarPastaDesafio_ = () => folder;
  sandbox.certificadoSalvarLinkPlanilha_ = () => ({ ok: true });
  sandbox.DriveApp = {
    Access: { ANYONE_WITH_LINK: 'ANYONE' },
    Permission: { VIEW: 'VIEW' },
    getFileById: id => {
      calls.push(id);
      if ((id === 'template-item' && configuredFails) || (id === DEFAULT_TEMPLATE && defaultFails)) {
        throw new Error(`falha ${id}`);
      }
      return {
        makeCopy: () => ({ getId: () => `tmp-${++generated}`, setTrashed() {}, getBlob: () => ({}) }),
        getBlob: () => ({ getAs: () => ({ setName: () => ({}) }) })
      };
    }
  };
  sandbox.SlidesApp = {
    openById: () => ({
      getSlides: () => [{ replaceAllText() {} }],
      saveAndClose() {}
    })
  };
  sandbox.Utilities = { formatDate: () => '15/06/2026' };
  sandbox.Session = { getScriptTimeZone: () => 'UTC' };
  sandbox.MimeType = { PDF: 'PDF' };
  sandbox.__calls = calls;
  return sandbox;
}

// 16. Falha do template configurado → tenta padrão.
{
  const ctx = generationSandbox();
  const result = ctx.gerarCertificadoDesafio_({ id_dgmb: '1', id_desafio: 'D1', id_item_estoque: 'ITEM1' });
  assert.equal(result.ok, true);
  assert.deepEqual(ctx.__calls.slice(0, 2), ['template-item', DEFAULT_TEMPLATE]);
}

// 17. Falha do configurado e do padrão → retorna erro.
{
  const ctx = generationSandbox({ defaultFails: true });
  const result = ctx.gerarCertificadoDesafio_({ id_dgmb: '1', id_desafio: 'D1', id_item_estoque: 'ITEM1' });
  assert.equal(result.ok, false);
  assert.equal(result.code, 'CERTIFICADO_GERACAO_SLIDES_ERROR');
  assert.deepEqual(ctx.__calls, ['template-item', DEFAULT_TEMPLATE]);
}

// 18. Certificado existente continua sendo reutilizado antes de resolver template.
{
  const ctx = load([HEADERS]);
  let resolved = false;
  ctx.certificadoBuscarDadosVisuais_ = () => ({});
  ctx.certificadoGetOuCriarPastaDesafio_ = () => ({});
  ctx.certificadoBuscarArquivoExistente_ = () => ({
    getUrl: () => 'https://drive.google.com/file/d/1234567890123456789012345/view',
    getId: () => '1234567890123456789012345'
  });
  ctx.certificadoSalvarLinkPlanilha_ = () => ({ ok: true });
  ctx.certificadoResolverTemplateSlides_ = () => { resolved = true; return {}; };
  const result = ctx.gerarCertificadoDesafio_({ id_dgmb: '1', id_desafio: 'D1', id_item_estoque: 'ITEM1' });
  assert.equal(result.reused, true);
  assert.equal(resolved, false);
}


const IMAGE = fs.readFileSync(path.join(ROOT, 'Meu Giro/CertificadoImagemService.gs'), 'utf8');
const TEMPLATE_IMAGEM_DEPRECATED = ['TEMPLATE', 'CERTIFICADO', 'IMAGEM', 'SLIDES', 'ID'].join('_') + '_';
assert.equal(IMAGE.includes(TEMPLATE_IMAGEM_DEPRECATED), false);
assert.match(IMAGE, /certificadoResolverTemplateSlides_\(\s*\(ctx \|\| \{\}\)\.id_desafio,\s*\(ctx \|\| \{\}\)\.id_item_estoque\s*\)/);
for (const placeholder of ['{{NOME}}', '{{DESAFIO}}', '{{META}}', '{{KM_REALIZADO}}', '{{STATUS}}', '{{PERIODO}}']) {
  assert.match(IMAGE, new RegExp(placeholder.replace(/[{}]/g, '\\$&')));
}

console.log('OK: 18 cenários e vínculos de imagem de template de certificado validados.');
