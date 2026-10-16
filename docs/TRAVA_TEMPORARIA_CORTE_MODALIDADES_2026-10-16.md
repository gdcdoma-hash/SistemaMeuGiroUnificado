# Trava de escrita temporária — SOMENTE PREPARAÇÃO, sem PROD

Sistema: **Meu Giro**. Branch base: `prep/corte-seletivo-meugiro-modalidades-2026-10-16`.

**Motivo:** a base final de PROD é compartilhada por dois projetos Apps Script. Não há suporte existente para congelar novas inscrições e registros de KM de modo coordenado.

O helper `Meu Giro/CorteManutencaoService.gs` adiciona leitura de uma única chave na planilha compartilhada, sem cache. A aba `DGMB_CORTE_MANUTENCAO` deve ter `A1=DGMB_CORTE_MODALIDADES_20261016`, `B1=PAUSADO` durante o corte e `B1=LIBERADO` após verificação. Aba ausente mantém rotina normal; conteúdo inválido ou erro de leitura bloqueia novas gravações por precaução.

API(s) protegida(s): `function registrarAtividade(idDgmb, dataAtividade, km, force, modalidadeEsportiva) {`, `function editarAtividade(payload) {`, `function excluirAtividade(payload) {`.

**Limites e pré-condições:**
- Estas funções não bloqueiam toda atividade administrativa ou operações de backend externas. Os administradores precisam suspender edições, conclusões, estoque e pagamentos durante o corte.
- Somente implantar a trava após teste STAGING, e **antes** da migração. Atualizar ambos os Web Apps versionados, preservar URLs e validar a pausa pelo backend (não somente pela interface).
- Enquanto a trava não estiver presente nos **deployments em uso**, a aba de controle NÃO protege escritas. Nunca definir PAUSADO e assumir que o PROD já está bloqueado sem conferir versões e retornos das APIs.
- Aguarde as operações iniciadas antes da ativação terminarem antes de modificar o schema (Apps Script pode ter execuções em andamento). O ideal é anunciar manutenção e realizar o corte em período de pouco movimento.
- Após publicar ambos os módulos e validar esquema, retirar pausa com `LIBERADO`. No rollback, verificar se ambas as versões em uso também respeitam a trava; as versões antigas @62/@8 **não a conhecem**.
- Esta branch é preparação. **Não publicar diretamente no PROD** nem promover à branch automatizada antes da homologação e do checkpoint.

Controle: https://github.com/gdcdoma-hash/dgmb-inscricoes-repescagem/issues/398
