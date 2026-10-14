# Sistema Meu Giro Unificado

O código operacional ativo está na pasta `Meu Giro/`.

## Estado atual
A referência mais recente de continuidade é:

`docs/ESTADO_ATUAL_2026-10-14.md`

O histórico técnico de configuração, autenticação, integração e infraestrutura permanece em:

`docs/ESTADO_ATUAL_2026-09-24.md`

**Homologado em 14/10:** na tela Registrar, preservar `Modalidade | Data da atividade | KM realizados` alinhados na mesma linha, com dropdown obrigatório para atleta multimodal, popup padrão e espaço para o histórico. Antes de alterar sessão, período, PRAZO_DIAS, integração com Portal Giro ou deployment, leia ambos os documentos.

## Ambiente DEV
- Branch operacional: `main`
- Script ID GAS: `1N-10MQMYaq_91O75W_TFF837vPR6KfSDpJwVQJNS7776IwGo56NuppzG`
- Deployment DEV: `AKfycbxCV-6fDri2y2ppdPC1JPLqCTxEOwXLuSameujgVkoYokm-sfQgPtDY4oQQ0Z_uVCmRKg`
- Planilha LAB em uso pelo Meu Giro DEV (resolvida por `EnvironmentConfig.gs`): `1ivfrue_toKUiopYq1Xlx4UR_yFrKHFpyIgVn19xqEUA`
- ID histórico de fallback DEV: `1sFxmiWmIrPlwXgSb56M8maHF5niRYVLWRMFSUQYXgHM` (não assumir como base atualmente em execução)

O GitHub Actions sincroniza o código e atualiza o deployment DEV automaticamente.

## Integração
Meu Giro e Portal Giro são Web Apps separados. A integração usa handoff temporário e, no modo embutido, iframe + `postMessage`.

O `doGet()` do Meu Giro deve manter `HtmlService.XFrameOptionsMode.ALLOWALL`; sem isso o Portal não consegue embutir o Meu Giro.

## Regra crítica de PRAZO_DIAS
Para desafios de prazo individual consolidados, usar `data_inicio_desafio` e `data_fim_desafio`. Não aplicar término mensal da `ListaDesafios`.

## Histórico
Auditorias específicas permanecem em `docs/`. Informações antigas de branch, planilha ou deployment devem ser comparadas com o estado atual antes de uso.
