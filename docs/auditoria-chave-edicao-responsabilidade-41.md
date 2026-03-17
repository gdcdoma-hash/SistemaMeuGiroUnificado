# Auditoria técnica — Responsabilidade 41 (chave_edicao)

## Escopo
Validação técnica da estabilidade da `chave_edicao` atual, sem religar funcionalidades e sem mudanças de comportamento.

## Onde a chave_edicao nasce
- O registro de atividade grava a primeira coluna (`Timestamp`) com `new Date()` no `appendRow` de `registrarAtividade(...)`.
- Portanto a origem material da chave é o valor da coluna de timestamp na aba `REGISTRO_KM`.

## Como a chave é formada e normalizada hoje
1. **Criação do registro**: `registrarAtividade(...)` grava `new Date()`.
2. **Leitura para painel**: `buscarAtividadesUsuario_(...)` lê `Timestamp` e chama `normalizarTimestampEdicao_(...)`.
3. **Normalização aplicada** (`normalizarTimestampEdicao_(...)`):
   - se for `Date`: formata em `dd/MM/yyyy HH:mm:ss` com `Session.getScriptTimeZone()`;
   - se não for `Date`: usa `String(valor).trim()` sem transformação adicional.
4. **Retorno para UI**: `chave_edicao` vai como string no payload do painel.
5. **Uso em edição/exclusão**: `editarAtividade(...)` e `excluirAtividade(...)` procuram a linha por comparação estrita de:
   - `normalizarTimestampEdicao_(timestamp_da_linha) === chave_edicao`
   - e `rowId === id_dgmb`.

## Estabilidade no fluxo (registro -> painel -> edição/exclusão)
- **Consistência interna**: painel, edição e exclusão usam a **mesma função de normalização** (`normalizarTimestampEdicao_`), reduzindo divergência local.
- **Timezone**: para valores `Date`, todos os pontos usam `Session.getScriptTimeZone()`, mantendo coerência enquanto a timezone do projeto não mudar entre emissão e uso da chave.
- **Tipo**: quando a célula vier como `Date`, a chave perde milissegundos por formatação em segundos.

## Riscos identificados
1. **Colisão por precisão de segundos**
   - `new Date()` tem ms, mas a chave é normalizada para `dd/MM/yyyy HH:mm:ss` (sem ms).
   - Dois lançamentos do mesmo `id_dgmb` no mesmo segundo podem compartilhar a mesma `chave_edicao`.
   - Em edição/exclusão, o código para no primeiro match (timestamp + id), podendo operar na linha errada em cenário de colisão.
2. **Dependência de representação textual para não-Date**
   - Se timestamp estiver como texto com formato diferente, a normalização só faz `trim()`.
   - Não há parse/canonicalização de strings de data/hora para um padrão único.
3. **Sensibilidade a mudança de timezone do script**
   - Se a timezone do projeto mudar depois, a mesma data pode formatar em string diferente.
   - Isso pode invalidar chaves emitidas anteriormente (janela entre carregar painel e acionar edição/exclusão, ou uso tardio).

## Conclusão objetiva
- A `chave_edicao` atual é **funcional em cenário comum**, mas **não é robusta o suficiente** como identificador estável para edição/exclusão futura em ambiente com concorrência/alta cadência.
- Principal fragilidade: **colisão por truncar para segundos**.
- Recomendação técnica (fora do escopo desta responsabilidade): adotar identificador imutável e único por linha (ex.: UUID/ID próprio), mantendo timestamp apenas como metadado.

## Arquivos inspecionados
- `Meu Giro/RegistroService.gs`
- `Meu Giro/PainelService.gs`
- `Meu Giro/Script.html`
- `Meu Giro/Utils.gs`
