# Documentação Fluxo Atualização Base de Contratos DCCC

## Objetivo
Este fluxo monitora a chegada de um e-mail específico do B.O que contém a planilha de consulta aos contratos da SEPLAG, processa seus anexos e os salva em um diretório do SharePoint, antes de excluir o e-mail original. Ele automatiza o arquivamento e evita retrabalho manual.

## Gatilho
- **Conector**: Office 365 Outlook
- **Operação**: OnNewEmailV3
- **Parâmetros Principais**:
  - **From**: bimg@prodemge.gov.br
  - **IncludeAttachments**: true
  - **SubjectFilter**: Planilha periódica de consulta aos Contratos SEPLAG.
  - **Importance**: Any
  - **FolderPath**: ID da pasta configurada no Outlook
  - **Autenticação**: Conexão shared_office365-1

## Fluxo de Execução
1. Inicializar variável Sucesso (boolean) com valor false.
2. Loop **Do until** Sucesso == true (limitado a 60 iterações ou 1 hora):
   1. **Apply to each** anexo recebido:
      1. **Get attachment (V2)**
      2. **Create file** no SharePoint
   2. Definir variável Sucesso = true
3. **Delete Email (V2)**: exclui o e-mail original.

## Ações

### Inicializar Variável

Tipo: InitializeVariable
Inputs:
  name: Sucesso
  type: boolean
  value: false


### Do until

Tipo: Until
Condição: @equals(variables(Sucesso), true)
Limites:
  count: 60
  timeout: PT1H


#### Apply to each

Foreach: @triggerOutputs()?body/attachments

1. **Get attachment (V2)**
   - Conector: Office 365 Outlook (GetAttachment_V2)
   - Inputs:
     - messageId: @triggerOutputs()?body/id
     - attachmentId: @items(Aplicar_a_cada)?id

2. **Create File**
   - Conector: SharePoint Online (CreateFile)
   - Inputs:
     - dataset: https://cecad365.sharepoint.com/sites/DCCCSeplag
     - folderPath: /Documentos Compartilhados/Planilhas de Controle e BI/Contratos/backup
     - name: DCCC-Paineis1.xlsx
     - body: @outputs(Get_attachment_(V2))?body/contentBytes
     - ContentTransfer: Chunked

#### Set Variable

Tipo: SetVariable
Inputs:
  name: Sucesso
  value: true


### Delete Email (V2)

Tipo: DeleteEmail_V2
Conector: Office 365 Outlook
Inputs:
  messageId: @triggerOutputs()?body/id


## Variáveis Utilizadas
- **Sucesso** (boolean): sinaliza se os anexos foram processados e salvos com êxito.

## Conexões Utilizadas
- Office 365 Outlook: shared_office365-1
- SharePoint Online: shared_sharepointonl-78f3c2f7-e6f3-4af0-8660-06606e421a67

## Tratamento de Falhas
Sem configurações específicas de Configure run after; todas as ações dependem de Succeeded.

## Resultados Esperados
- Arquivos salvos em SharePoint no caminho configurado.
- E-mail original excluído após processamento.

## Observações e Pontos de Atenção
- O arquivo salvo (DCCC-Paineis1.xlsx) é nome fixo e sobrescreverá versões anteriores.
- Verificar permissões na biblioteca do SharePoint e na caixa de e-mail.
- Credenciais de conexão devem estar ativas e válidas para ambas as APIs.
