# Arquitetura — Meu Giro

Atualizado em: 2026-09-16

## Papel no sistema
O Meu Giro é um módulo do ecossistema Desafio Giro. Para o atleta, deve parecer parte do mesmo sistema do Portal Giro, embora atualmente esteja em projeto Apps Script/repositório separado.

## Integração
- Entrada principal do atleta: Portal Giro.
- Identidade deve vir da sessão/handoff autenticado.
- O boot integrado deve prevalecer sobre restaurações antigas do navegador.
- Não expor CPF/ID como mecanismo de autenticação na URL.
- Navegação embutida: `Portal Giro | Início | Desafios | Registrar | Conquistas`.
- Ranking permanece fora até reconstrução.

## Estabilidade
Não reintroduzir sem nova análise:
- altura dinâmica/ResizeObserver/READY do iframe;
- cache integral do iframe para retorno instantâneo.
Essas estratégias já produziram regressões e foram revertidas.

## Administração
Meu Giro Admin será reorganizado somente dentro da revisão administrativa registrada em `CONTINUIDADE.md` e na documentação principal do Portal. A revisão deve preservar cálculos, registros, auditorias e identidade do atleta.

## Fonte de dados
Planilha oficial conhecida: `18UCv96cqQMShaSabAhnX0AOLILAOcujCj-yBLUT1eJs`.

## Continuidade
GitHub é a fonte oficial. Produção não deve ser atualizada automaticamente.