# Fluxo Power Automate – Integração de Planilha ISSQN e Registro de Códigos de Barras

## Objetivo

Este fluxo automatiza a extração de dados filtrados da planilha **CONTROLE DCF 2024.xlsx** (navegador) para a planilha local **planilha_controle_issqn.xlsx**, e, em seguida, realiza o registro de cada item no SIAFI via aplicação **InputMinot**, incluindo a captura e inserção do código de barras da guia de ISSQN.

## Documentação

O fluxo principal chama, em sequência, dois subfluxos:

1. **Planilha Liliane**
2. **Registrar Código de Barras Liliane**

### Subfluxo: Planilha Liliane

1. Lançar Chrome  
   - Ação: `WebAutomation.LaunchChrome.LaunchChrome`  
   - Parâmetros:  
     - `Url`: `=vUrlPlanilhaDcf`  
     - `WindowState`: Maximizado  
     - `ClearCache`: False  
     - `ClearCookies`: False  
     - `BrowserInstance` ⇒ `vPlanilhaDcf`  
   - O que faz: Esta etapa inicializa o navegador Chrome e abre a planilha online “CONTROLE DCF 2024.xlsx” usando a URL configurada na variável `vUrlPlanilhaDcf`. Permite ao fluxo acessar e interagir com a planilha em ambiente web.

2. Anexar à aba da planilha  
   - Ação: `WebAutomation.LaunchChrome.AttachToChromeByTitle`  
   - Parâmetros: 
     - `TabTitle`: `CONTROLE DCF 2024.xlsx`  
     - `BrowserInstance` ⇒ `PlanilhaDcfChrome`  
   - O que faz: Conecta o fluxo à aba já aberta no Chrome cujo título corresponde à planilha de controle, garantindo que as ações seguintes sejam direcionadas ao navegador correto.

3. Solicitar seleção de dados  
   - Ação: `Display.ShowMessageDialog.ShowMessage`  
   - Parâmetros:  
     - `Title`: `Atenção!`  
     - `Message`: Instruções para aplicar filtros e selecionar registros, clicar em `A2` e confirmar.  
     - `Icon`: Warning  
   - O que faz: Exibe uma caixa de diálogo com instruções para o usuário aplicar filtros na planilha web e selecionar manualmente os dados de interesse antes de prosseguir.

4. Selecionar e copiar registros  
   - Ações:  
     - `MouseAndKeyboard.SendKeys.FocusAndSendKeysByInstanceOrHandle` ⇒ Seleção (Shift+Ctrl+Down e Right×20)  
     - `MouseAndKeyboard.SendKeys.FocusAndSendKeysByInstanceOrHandle` ⇒ `Ctrl+C`  
   - O que faz: Seleciona a coluna de dados filtrados na planilha web a partir da célula A2, abrange as colunas necessárias e copia o conteúdo para a área de transferência.

5. Abrir planilha local  
   - Ação: `Excel.LaunchExcel.LaunchAndOpenUnderExistingProcess`  
   - Parâmetros:  
     - `Path`: `%InputFolderPath%\planilha_controle_issqn.xlsx`  
     - `Visible`: True  
     - `Instance` ⇒ `PlanilhaExcel`  
   - O que faz: Abre no Excel a planilha local `planilha_controle_issqn.xlsx` a partir da pasta definida em `InputFolderPath`, deixando-a visível para receber os dados copiados.

6. Posicionar e ativar célula  
   - Ações:  
     - `MouseAndKeyboard.SendKeys` ⇒ `{Down}`  
     - `Excel.ActivateCellInExcel` ⇒ Coluna `A`, Linha `2`  
   - O que faz: Move o cursor para a primeira célula livre (A2) na planilha local, preparando o local de colagem dos dados.

7. Formatar Texto para Colunas e colar dados  
   - Ações:  
     - `MouseAndKeyboard.SendKeys` ⇒ Atalho do assistente de Texto para Colunas (`Ctrl+T`, `Alt+C`, `K`, `L`)  
     - Ajuste de direção com `{Left}`  
     - `MouseAndKeyboard.SendKeys` ⇒ `Ctrl+V`  
   - O que faz: Executa o assistente de “Texto para Colunas” para separar corretamente os campos copiados, ajusta o delimitador e cola os dados formatados na planilha.

8. Salvar e fechar planilha  
   - Ação: `Excel.CloseExcel.CloseAndSave`  
   - O que faz: Salva as alterações na planilha local e fecha a instância do Excel, garantindo que os dados estejam gravados no arquivo.

---

### Subfluxo: Registrar Código de Barras Liliane

1. **Abrir planilha local**  
   - Ação: `Excel.LaunchExcel.LaunchAndOpenUnderExistingProcess`  
   - Parâmetros:  
     - `Path`: `%InputFolderPath%\planilha_controle_issqn.xlsx`  
     - `Visible`: False  
     - Instance ⇒ `Planilha`  
   - O que faz: Abre a planilha de controle em modo invisível para leitura dos dados já formatados.

2. **Obter primeira célula livre**  
   - Ação: `Excel.GetFirstFreeColumnRow`  
   - Saídas:  
     - `FirstFreeColumn` ⇒ `ColLivre`  
     - `FirstFreeRow` ⇒ `LiLivre`  
   - O que faz: Identifica a primeira coluna e linha vazias na planilha, delimitando o intervalo de dados existente para leitura.

3. **Ler dados para tabela**  
   - Ação: `Excel.ReadFromExcel.ReadCells`  
   - Intervalo: `A1` até `ColLivre-1` / `LiLivre-1`  
   - Parâmetros:  
     - `FirstLineIsHeader`: True  
     - Saída ⇒ `DadosIssqn`  
   - O que faz: Lê todos os registros existentes na planilha local e armazena-os em uma variável de tabela (`DadosIssqn`) para processamento.

4. **Fechar planilha local**
   - Ação: `Excel.CloseExcel.Close`  
   - O que faz: Fecha o Excel sem salvar, pois não houve alteração no arquivo nesta etapa.

5. **Inicializar contagem de linha** 
   - Ação: `SET SheetLineCount TO 2`  
   - O que faz: Define o contador inicial de linha (`SheetLineCount`) como 2, indicando que o processamento começará a partir da segunda linha da tabela de dados.

6. **Conectar à planilha web**  
   - Ação: `WebAutomation.LaunchChrome.AttachToChromeByTitle`  
   - Parâmetros:  
     - `TabTitle`: `CONTROLE DCF 2024.xlsx`  
     - `BrowserInstance` ⇒ `PlanilhaDcfChrome`  
   - O que faz: Reconecta o fluxo à aba da planilha web aberta anteriormente, pronta para informar o código de barras.

7. **Solicitar seleção da célula de OP**  
   - Ação: `Display.ShowMessageDialog.ShowMessage`  
   - Parâmetros: instruções para aplicar filtros e selecionar a primeira célula vazia da coluna de OP.  
   - O que faz: Exibe instruções ao usuário para filtrar e selecionar manualmente a célula onde será registrada a OP no ambiente web.

8. **Carregar tabelas de contas**  
   - Ações:  
     - `Variables.GenerateDataTableFromCSV` ⇒ `Contas1501`  
     - `Variables.GenerateDataTableFromCSV` ⇒ `ContasOutrasUos`  
   - O que faz: Importa tabelas auxiliares de contas orçamentárias a partir de arquivos CSV, usadas nas etapas de cálculo e conferência de conta.

9. **Inicializar variáveis de exceção**  
   - Ações: `SET NumberException, LengthException, FloatException, RangeException, SameDayException, WeekendException` para `''`  
   - O que faz: Reinicia variáveis de controle de erro, assegurando que eventuais exceções sejam tratadas corretamente durante o loop.

10. **Loop de processamento para cada registro**  
    - Descrição: Itera sobre cada linha de `DadosIssqn`, extraindo informações necessárias e informando-as no sistema SIAFI via InputMinot.  
    - **Passos principais dentro do loop:**  
      - ID: `SET ID TO LiDadosIssqn['ID']`  
      - Processo SEI: `Scripting.RunPythonScript` (substituir `/` por `.`) ⇒ `ProcessoSEI`  
      - Número da NF: `SET NumeroNF TO LiDadosIssqn['Número da Nota Fiscal']`  
      - CNPJ: `Scripting.RunPythonScript` (remover pontuação) ⇒ `CnpjPref`  
      - Prefixo Município: `Scripting.RunPythonScript` (tratar acentuação e nome) ⇒ `Pref`  
      - Dados Orçamentários:  
        - `GmiFp` e `Fp` via script Python  
        - `ValorIssqn` via script Python de conversão  
        - `Guia`  
      - **Subfluxos auxiliares:**  
        - `CALL 'Extrair Uo'`  
        - `CALL 'Informar Código de Barras'`  
        - `CALL 'Conferir Conta'`  
        - `CALL 'Informar Multa'`  
        - `CALL 'Informar Linha'`  
        - `CALL 'Informar Data Pagamento'`  
        - `CALL 'Informar Op'`  
      - **Interações com SIAFI (InputMinot):**  
        - `MouseAndKeyboard.SendKeys` para preencher campos em sequência  
        - `Display.ShowMessageDialog` para conferência de texto  
        - `MouseAndKeyboard.SendKeys {F5}` (simulação) e desativação da execução real  
    - O que faz: Em cada iteração, o fluxo coleta dados da planilha local, processa informações orçamentárias via scripts Python, e preenche automaticamente os campos no sistema SIAFI, incluindo a inserção do código de barras e demais detalhes.

---

## Passos Para Execução

1. **Pré-requisitos**  
   - Instalar Power Automate Desktop e efetuar login corporativo  
   - Definir as variáveis de ambiente:  
     - `vUrlPlanilhaDcf`: URL da planilha web  
     - `InputFolderPath`: caminho da pasta local onde está `planilha_controle_issqn.xlsx`  
   - O que faz: Garante que todos os componentes necessários estejam configurados antes de iniciar o fluxo.

2. **Ambiente**  
   - Ter a planilha `CONTROLE DCF 2024.xlsx` aberta no Chrome  
   - Garantir que `planilha_controle_issqn.xlsx` esteja em `%InputFolderPath%`  
   - Disponibilizar os arquivos CSV para contas: `InputContas1501Csv`, `InputContasOutrasUosCsv`  
   - O que faz: Prepara o ambiente de execução com os arquivos e instâncias corretas.

3. **Execução**  
   - Executar o subfluxo **Planilha Liliane**  
   - Executar o subfluxo **Registrar Código de Barras Liliane**  
   - Seguir as instruções das caixas de diálogo para seleção de células, inserção de códigos de barras e conferências  
   - O que faz:Executa de forma sequencial os dois subfluxos, automatizando completamente o processo de extração, formatação e registro dos dados no sistema.
