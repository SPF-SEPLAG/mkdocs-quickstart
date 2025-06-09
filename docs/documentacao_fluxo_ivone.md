
# Fluxo Power Automate – Registro de Pagamento de Notas Fiscas

## Objetivo

Este fluxo automatiza o registro de ordens de pagamento no sistema SIAFI e no SEI, percorrendo registros de uma planilha de controle e anexando os relatórios de pagamento no processo SEI correspondente.

## Documentação

O fluxo é composto por 1 fluxo principal (Main).

1. **Inicialização e Leitura da Planilha**
   - **Excel.LaunchExcel**  
     - Path: `%file_path_fluxo%\\planilha_controle_pagamento.xlsx`  
     - Visible: False  
     - Instance: `planilha`  
     - Descrição: Abre o Microsoft Excel em segundo plano e carrega a planilha de controle de pagamentos para leitura dos dados.
   - **Excel.GetFirstFreeColumnRow**  
     - Instance:`planilha`  
     - FirstFreeColumn: `coluna_livre`  
     - FirstFreeRow: `linha_livre`  
     - Descrição: Identifica a primeira coluna e a primeira linha vazias na planilha, determinando o alcance do intervalo de dados.
   - **Excel.ReadFromExcel.ReadCells**  
     - Instance: `planilha`  
     - StartColumn: `A`  
     - StartRow: `1`  
     - EndColumn: `coluna_livre - 1`  
     - EndRow: `linha_livre - 1`  
     - ReadAsText: False  
     - FirstLineIsHeader: True  
     - RangeValue: `dados_pagamento`  
     - Descrição: Lê todas as células preenchidas no intervalo especificado, convertendo-as em uma tabela de dados chamada `dados_pagamento`, com cabeçalhos.
   - **Excel.CloseExcel.Close**  
     - Instance: `planilha`  
     - Descrição: Fecha a instância do Excel, garantido que o arquivo não fique aberto em segundo plano.

2. **Loop de Processamento por Registro**  
   - LOOP FOREACH `linha` IN `dados_pagamento`  
   - Descrição: Itera sobre cada registro de pagamento lido da planilha para processar individualmente.

   1. **Acesso ao SIAFI e Geração do Relatório**
      - **WebAutomation.LaunchChrome.LaunchChrome**  
        - Url: `https://www.siafi.mg.gov.br/fnad/Relatorios.mtw`  
        - BrowserInstance: `siafi_branco`  
        - Descrição: Abre o Google Chrome e acessa a página de relatórios do SIAFI para gerar o documento de pagamento.
      - **IF WebAutomation.IfWebPageContains.WebPageContainsText**  
        - BrowserInstance: `siafi_branco`  
        - Text: `Login Certificação Digital`  
        - THEN: Preenche e envia credenciais de certificação digital.  
        - Descrição: Verifica se é necessário autenticar via certificado digital e realiza o login automaticamente, se solicitado.
      - **WebAutomation.Click.Click**  
        - *Element: Anchor 'Consulta / Impressão Documento Específico'  
        - Descrição: Seleciona a opção de consulta específica de documentos para aplicar filtros avançados.
      - **WAIT** 2 segundos  
        - Descrição: Aguarda o carregamento completo da página antes de continuar.
      - **WebAutomation.ExecuteJavascript**  
        - Script: `document.getElementById('cmbUE').value = '%linha['Unidade Executora']%';`  
        - **Descrição:** Define a Unidade Executora no filtro de acordo com o valor da planilha.
      - **SendKeys**  
        -*Keys: `%linha['OP']%{Return}`  
        - **Descrição:** Insere o número da Ordem de Pagamento e confirma, gerando o relatório.
      - **MouseAndKeyboard.SendKeys.FocusAndSendKeysByTitleClass**  
        - Keys: `{Control}S`  
        - **Descrição:** Aciona o atalho para salvar o relatório como PDF na localização definida.
      - **WebAutomation.CloseWebBrowser** e **SendKeys** `{Alt}F4`  
        - **Descrição:** Fecha a janela do navegador, encerrando a sessão no SIAFI.

   2. **Registro do Pagamento no SEI**
      - **WebAutomation.LaunchChrome.LaunchChrome**  
        - Url:`https://www.sei.mg.gov.br/sip/login.php?sigla_orgao_sistema=GOVMG&sigla_sistema=SEI&infra_url=L3NlaS8=`  
        - BrowserInstance: `sei`  
        - Descrição: Abre o Chrome e acessa o sistema SEI para registro de documentos.
      - **SendKeys**  
        - Keys: `%sei_usuario%`, `%sei_senha%`, `SEPLAG`, `{Return}`  
        - **Descrição:** Insere credenciais e seleciona o órgão, efetuando o login no SEI.
      - **WAIT** 2 segundos  
        - **Descrição:** Aguarda o login ser processado e o painel principal carregar.
      - **SendKeys** `{Escape}`  
        - **Descrição:** Fecha notificações ou pop-ups iniciais.
      - **WebAutomation.Click.Click**  
        - Element: Campo `txtPesquisaRapida`  
        - **Descrição:** Posiciona o cursor para pesquisa rápida de processos.
      - **SendKeys**  
        - Keys: `%linha['Processo SEI']%{Return}`  
        - **Descrição:** Digita o número do processo e inicia a pesquisa.
      - **WAIT UNTIL** "Processo aberto"  
        - **Descrição:** Aguarda confirmação de que o processo foi carregado.
      - **WebAutomation.Click.Click**  
        - Element: Image 'Incluir Documento'  
        - **Descrição:** Abre a interface para anexar um novo documento ao processo.
      - **IF UIAutomation.IfImage.ExistsOnScreen** `[imgrepo['mais_verde']]` **THEN**  
        - WebAutomation.Click.Click: Image 'Exibir todos os tipos'  
        - **Descrição:** Se necessário, expande a lista completa de tipos de documento.
      - **SendKeys**  
        - Keys: `Externo{Tab}{Control}Enter`  
        - **Descrição:** Seleciona a natureza 'Externo' e confirma.
      - **WebAutomation.LaunchChrome.AttachToChromeByTitle**  
        - Title: `sei`  
        - **Descrição:** Conecta novamente à janela do SEI para upload.
      - **WebAutomation.ExecuteJavascript**  
        - Script: `selSerie = "263";` + seleção de opções  
        - **Descrição:** Configura série e visibilidade do documento.
      - **DateTime.GetCurrentDateTime.Local** → `data_atual`  
        - **Descrição:** Captura data e hora atuais.
      - **Text.ConvertDateTimeToText.FromDateTime** → `data_atual`  
        - **Descrição:** Converte data/hora em texto formatado.
      - **SendKeys**  
        - Keys: `data_atual{Tab}%linha['OP']%{Tab}pagamento.pdf{Return}`  
        - **Descrição:** Preenche campos de data, número da OP e seleciona o PDF para upload.
      - **File.Delete**  
        - Path: `%file_path_fluxo%\\pagamento.pdf`  
        - **Descrição:** Remove o arquivo PDF temporário após upload.
      - **SendKeys** `{Alt}F4`  
        - **Descrição:** Fecha a janela de registro de documento, concluindo o processo.

## Passos Para Execução do Fluxo no Power Automate Desktop

**Caso já tenha o ambiente preparado (Power Automate instalado, fluxo importado e variáveis definidas), pode-se passar para o passo 5.**

1. **Download e Instalação do Power Automate Desktop**  
   - **Descrição:** Acesse o site da Microsoft, faça o download e instale o PAD para começar a automação.  
2. **Login no PAD**  
   - **Descrição:** Abra o Power Automate Desktop e autentique-se com sua conta corporativa.  
3. **Importação do Fluxo**  
   - **Descrição:** Crie um novo fluxo chamado **Registro de Pagamento Ivone** e importe o conteúdo de `pgto fluxo ivone.txt`.  
4. **Definição de Variáveis de Entrada**  
   - **Descrição:** Configure as variáveis necessárias:  
     - `file_path_fluxo`: caminho da pasta com `planilha_controle_pagamento.xlsx`.  
     - `usuario_siafi` e `senha_siafi`: credenciais do SIAFI.  
     - `sei_usuario` e `sei_senha`: credenciais do SEI.  
5. **Preparação da Planilha de Controle**  
   - **Descrição:** Certifique-se de que `planilha_controle_pagamento.xlsx` esteja presente em `file_path_fluxo` e contenha as colunas `Unidade Executora`, `OP` e `Processo SEI`.  
6. **Pré-requisitos de Navegador e Internet**  
   - **Descrição:** Verifique se o Google Chrome está instalado e se há conexão estável com a internet.  
7. **Execução do Fluxo**  
   - **Descrição:** No PAD, selecione o fluxo **Registro de Pagamento Ivone** e clique em **Executar**.  
8. **Acompanhamento e Verificação**  
   - **Descrição:** Monitore o progresso e confirme que os relatórios foram gerados no SIAFI e registrados no SEI com sucesso.  

