# Responsabilidade 22 — Mapeamento de consumidores do payload do painel

## Escopo
Mapeamento objetivo dos consumidores de `getPainelUsuario()` e das propriedades lidas de `data` no frontend atual, com foco nos aliases temporários de ranking.

## Arquivos analisados
- `Meu Giro/PainelService.gs`
- `Meu Giro/Script.html`
- `Meu Giro/Index.html`
- `Meu Giro/FrasesService.gs`

## Onde `getPainelUsuario()` é consumido/encadeado
1. **Frontend (consumo direto):**
   - `Meu Giro/Script.html` chama `google.script.run...getPainelUsuario(currentUser.id_dgmb)`.
2. **Backend (encadeamento interno):**
   - `Meu Giro/FrasesService.gs` chama `getPainelUsuario(idDgmb)` dentro de `getOutraFraseMotivacional`.

## Propriedades do payload `data` com leitura no frontend atual
Leituras encontradas em `renderPainel(p)` de `Meu Giro/Script.html`:
- `nome`
- `cidade_uf`
- `id_dgmb`
- `meta`
- `realizado`
- `restante`
- `percentual`
- `diaAtual`
- `kmIdealAtual`
- `diasRestantes`
- `kmPorDiaRestante`
- `frase`
- `atividades`
- `totalPedalado` (com fallback para `realizado`)

## Aliases temporários com evidência real de uso
- **`totalPedalado`**: usado no frontend (`formatNumber(p.totalPedalado || p.realizado || 0)`).

## Aliases temporários sem evidência clara de uso no consumo encontrado
No frontend atual mapeado, **não foi encontrada leitura** para:
- `posicao_ranking`
- `posicaoRanking`
- `total_participantes`
- `totalParticipantes`

Observação: os elementos HTML para posição existem (`painel-posicao-ranking`, `painel-posicao-detalhe`), mas não há código JS no frontend atribuindo valores desses campos.

## Snake_case, camelCase e fallback
- Há uso misto de convenções no frontend:
  - snake_case: `cidade_uf`, `id_dgmb`
  - camelCase: `diaAtual`, `kmIdealAtual`, `totalPedalado`
- Fallback identificado no frontend:
  - `totalPedalado || realizado`
- **Não foi encontrado fallback** entre os pares de alias de ranking:
  - `posicao_ranking` ↔ `posicaoRanking`
  - `total_participantes` ↔ `totalParticipantes`

## Conclusão
A decisão correta nesta responsabilidade é **diff funcional zero**:
- não há evidência de dependência atual no frontend para os aliases temporários de ranking;
- missão solicitada é de mapeamento e diagnóstico, sem refatorar/remover compatibilidade.
