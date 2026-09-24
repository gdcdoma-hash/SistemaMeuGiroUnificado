MEU GIRO — ESTADO OPERACIONAL ATUAL
Atualizado em 24/09/2026

Este arquivo substitui as instruções antigas de "MEU GIRO V1".

FONTE OFICIAL
- Repositório: gdcdoma-hash/SistemaMeuGiroUnificado
- Branch: main
- Código ativo: pasta "Meu Giro"
- GitHub é canônico.

AMBIENTE DEV
- Script ID: 1N-10MQMYaq_91O75W_TFF837vPR6KfSDpJwVQJNS7776IwGo56NuppzG
- Deployment: AKfycbxCV-6fDri2y2ppdPC1JPLqCTxEOwXLuSameujgVkoYokm-sfQgPtDY4oQQ0Z_uVCmRKg
- Planilha oficial: 1sFxmiWmIrPlwXgSb56M8maHF5niRYVLWRMFSUQYXgHM

ABAS PRINCIPAIS
- DadosPessoais
- dgmbDesafios
- REGISTRO_KM
- MEU_GIRO_RESUMO
- ListaDesafios
- FRASES
- CONFIG_CERTIFICADO_TEMPLATE
- _PORTAL_HANDOFF

INTEGRAÇÃO COM PORTAL GIRO
- Portal e Meu Giro são Web Apps separados.
- O Portal cria handoff temporário para o atleta.
- O Meu Giro pode funcionar embutido no Portal.
- O doGet() deve manter XFrameOptionsMode.ALLOWALL.
- Não remover essa configuração: sem ela o iframe é bloqueado.

PRAZO_DIAS
- Prazo_Dias > 0 = janela individual.
- Após data_consolidacao, usar data_inicio_desafio e data_fim_desafio.
- Não substituir a data fim individual pelo fim mensal.

DEPLOY
- Não executar deploy manual como fluxo padrão.
- O workflow .github/workflows/sync-gas.yml faz push e atualiza o deployment DEV.
- Só testar após o Action concluir.

DOCUMENTAÇÃO COMPLETA
Ver:
docs/ESTADO_ATUAL_2026-09-24.md
