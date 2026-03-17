MEU GIRO V1

1. Cole todos os arquivos no projeto Apps Script.
2. Em Config.gs, troque SPREADSHEET_ID pelo ID real da sua planilha.
3. Garanta que as abas existentes estejam com estes nomes exatos:
   - dadosPessoais
   - dgmgCamisa
4. Crie a nova aba REGISTRO_KM com cabeçalho na linha 1:
   Timestamp | ID_DGMB | Data_Atividade | KM | activity_id
5. Crie a nova aba FRASES com cabeçalho na linha 1:
   Fase | Situacao | Frase
6. Publique como Web App.

Observações:
- Login apenas por CPF.
- Não há cadastro novo.
- Não há inscrição.
- O sistema atualiza Distancia_Realizada na aba dgmgCamisa.
