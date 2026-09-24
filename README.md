# Sistema Meu Giro Unificado

O código operacional ativo está na pasta `Meu Giro/`.

## Estado atual
A referência principal de continuidade é:

`docs/ESTADO_ATUAL_2026-09-24.md`

Antes de alterar sessão, período, PRAZO_DIAS, integração com Portal Giro ou deployment, leia esse documento.

## Ambiente DEV
- Branch operacional: `main`
- Script ID GAS: `1N-10MQMYaq_91O75W_TFF837vPR6KfSDpJwVQJNS7776IwGo56NuppzG`
- Deployment DEV: `AKfycbxCV-6fDri2y2ppdPC1JPLqCTxEOwXLuSameujgVkoYokm-sfQgPtDY4oQQ0Z_uVCmRKg`
- Planilha oficial: `1sFxmiWmIrPlwXgSb56M8maHF5niRYVLWRMFSUQYXgHM`

O GitHub Actions sincroniza o código e atualiza o deployment DEV automaticamente.

## Integração
Meu Giro e Portal Giro são Web Apps separados. A integração usa handoff temporário e, no modo embutido, iframe + `postMessage`.

O `doGet()` do Meu Giro deve manter `HtmlService.XFrameOptionsMode.ALLOWALL`; sem isso o Portal não consegue embutir o Meu Giro.

## Regra crítica de PRAZO_DIAS
Para desafios de prazo individual consolidados, usar `data_inicio_desafio` e `data_fim_desafio`. Não aplicar término mensal da `ListaDesafios`.

## Histórico
Auditorias específicas permanecem em `docs/`. Informações antigas de branch, planilha ou deployment devem ser comparadas com o estado atual antes de uso.
