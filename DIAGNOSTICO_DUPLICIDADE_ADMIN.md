# Diagnóstico — Duplicidade de funções (área admin)

Data: 2026-04-21
Escopo: varredura do projeto inteiro para os nomes:
- `listarPendenciasValidacaoCertificado`
- `atualizarStatusValidacaoCertificadoAdmin`
- `carregarAdminValidacaoPendencias`
- `renderAdminValidacaoLista`
- `formatAdminStatusValidacao_`

Comando usado:

```bash
rg -n "listarPendenciasValidacaoCertificado|atualizarStatusValidacaoCertificadoAdmin|carregarAdminValidacaoPendencias|renderAdminValidacaoLista|formatAdminStatusValidacao_" /workspace/SistemaMeuGiroUnificado
```

## 1) Ocorrências por função e arquivo

### `listarPendenciasValidacaoCertificado`
- Definição: `Meu Giro/AdminCertificadoService.gs:23`
- Uso/chamada: `Meu Giro/Script.html:3313`

### `atualizarStatusValidacaoCertificadoAdmin`
- Definição: `Meu Giro/AdminCertificadoService.gs:138`
- Uso/chamada: `Meu Giro/Script.html:3410`

### `carregarAdminValidacaoPendencias`
- Definição: `Meu Giro/Script.html:3291`
- Uso/chamada interna: `Meu Giro/Script.html:3288`, `Meu Giro/Script.html:3404`
- Uso/chamada no HTML (onclick): `Meu Giro/Index.html:293`

### `renderAdminValidacaoLista`
- Definição: `Meu Giro/Script.html:3316`
- Uso/chamada: `Meu Giro/Script.html:3307`

### `formatAdminStatusValidacao_`
- Definição: `Meu Giro/Script.html:3421`
- Uso/chamada: `Meu Giro/Script.html:3331`, `Meu Giro/Script.html:3373`

## 2) Conclusão objetiva sobre conflito

Não foi encontrada duplicidade de definição para nenhum dos cinco nomes investigados.

Cada função aparece com **apenas uma definição ativa** no código-fonte atual:
- funções de backend em `AdminCertificadoService.gs`
- funções de frontend em `Script.html`

Logo, com base no repositório atual, **não há conflito por sobrescrita causado por redefinição duplicada desses nomes**.

## 3) Correção mínima proposta (se comportamento em produção segue incorreto)

Como não há duplicidade no código atual, a correção mínima recomendada é operacional/deploy, sem refatoração:

1. Confirmar o deployment ativo do Apps Script/Web App aponta para a versão mais recente do projeto.
2. Se necessário, publicar nova versão e atualizar o Web App para essa versão explicitamente.
3. Invalidar cache do navegador (hard reload) e repetir o teste em janela anônima.
4. (Opcional mínima de diagnóstico) adicionar um `console.log('build: <timestamp>')` no carregamento do `Script.html` para confirmar rapidamente se o cliente está recebendo o bundle novo.

Esses passos tratam o cenário clássico de "código correto no repo, comportamento antigo em runtime" sem alterar regra de negócio.
