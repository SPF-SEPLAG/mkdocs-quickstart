# Fluxo Power Automate – Liquidação

## Objetivo

Este fluxo automatiza a liquidação, registrando no sistema SIAFI os dados das notas fiscais a serem liquidadas.

## Documentação

O fluxo é composto por um fluxo principal (Main) e três subfluxos: Inicializar Variáveis de Fluxo, Executar Siafi Terminal e Siafi Web.

1. O fluxo principal (Main) inicializa as variáveis chamando o subfluxo Inicializar Variáveis.
2. Após a inicialização, o fluxo define a variável `FilePath`, que armazena o caminho da planilha auxiliar (`dados.xlsx`) localizada no diretório especificado. Esse arquivo contém os dados necessários para o processo de liquidação no SIAFI.
3. Após acessar a planilha auxiliar, o fluxo define quatro variáveis essenciais para o processo:
   - **Minot**: armazena a janela em instância do SIAFI aberta.
   - **CnpjInss**: contém o CNPJ do INSS.
   - **Uo**: guarda o número da unidade orçamentária (padrão: 1501).
   - **RowIndex**: armazena o valor `1`, utilizada para navegação na planilha.
4. Após a definição das variáveis, o fluxo inicia o Excel usando `FilePath` e obtém a primeira coluna e linha livre, armazenando-as nas variáveis `ColunaLivre` e `LinhaLivre`.
5. O fluxo lê os dados da planilha do intervalo `A1` até a última célula não vazia, armazenando-os em `DadosLiquidacao`.
6. O fluxo itera sobre cada linha em `DadosLiquidacao` (loop for each), armazenando a linha atual em `Linha`.
7. Para cada `Linha`, o fluxo envia os dados para o SIAFI (instância `Minot`), navegando no sistema.
8. **Condições IF**:
   - Se `Linha["Conformidade"]` ≠ `N/A`, lança o valor; caso contrário, pula com TAB.
9. Verifica se `Linha["Valor INSS"]` e `Linha["Valor ISSQN"]` estão vazias:
   - Se vazias: preenche com `s`; senão: preenche com `n`.
10. Obtém a data atual (`DataAtual`) e formata como mês de competência.
11. Converte o valor da NF em texto sem separador decimal.
12. Define `Ordenador` com o valor de `Linha["MASP"]`.
13. Remove dígito final de `MASP` se iniciar com `m` e envia para o SIAFI.
14. Se `Valor INSS` não está vazio, formata e armazena em `ValorInss` e `CompetenciaInss`.
15. Verifica `Valor ISSQN`: se não vazio, formata e armazena em `ValorIssqn`. Valida CNPJ Município, interrompendo com erro se ausente ou tratando e armazenando em `CnpjIssqn`.
16. **Lógica composta**:
    - Ambos INSS e ISSQN: envia `CnpjInss`, `CompetenciaInss`, `ValorInss`, `ValorIssqn`.
    - Apenas ISSQN: envia `CnpjIssqn`, `ValorIssqn`.
    - Apenas INSS: envia `CnpjInss`, `CompetenciaInss`, `ValorInss`.
17. Lança `Nome Credor`, `Número Nota Fiscal`, `ID` e `Processo SEI`, removendo barras `/` se presentes.
18. Formata `Processo SEI`, removendo `/` e envia.
19. Envia F5 para realizar liquidação.
20. Copia o conteúdo da tela (Ctrl + A, Ctrl + C) para `SeqLiquidacao`, processa com Python e atualiza.
21. Reabre a planilha e grava `SeqLiquidacao` na coluna `U`, usando `RowIndex`.
22. Envia F3 três vezes para retornar ao início.
23. Após concluir as etapas no terminal do SIAFI, o fluxo inicia as etapas do SIAFI Web.
24. Abre a planilha auxiliar (`FilePath`) no Power Automate.
25. Lê os dados em `DadosLiquidacao` novamente.
26. Itera sobre cada linha em `DadosLiquidacao`.
27. Abre o Google Chrome e acessa o portal SIAFI Web (instância `SiafiBranco`).
28. Procura pelo texto "Login Certificação Digital" e navega até os campos de login.
29. Preenche as credenciais com `UsuarioSiafi` e `SenhaSiafi`.
30. Clica em "Consulta / Impressão Documento Específico".
31. Executa JavaScript para selecionar a Unidade Executora (elemento "UE").
32. Navega com TAB e Down para preencher No.Empenho e Seq.Liquidação.
33. Pressiona Ctrl + S para salvar o relatório (`GeraRelatorioEmpenhoServlet`).
34. Na janela de diálogo, define o caminho usando `FilePath + "\\liquidação.pdf"` e confirma com Enter.
35. Fecha as janelas do SIAFI Web e do gerador de relatórios.
36. Abre nova instância do Chrome e acessa o portal SEI (variável `Sei`).
37. Faz login no SEI com `SeiSenha`.
38. Clica na barra de pesquisa e busca `Processo SEI` (valor em `Linha`).
39. Pressiona Enter para localizar o processo.
40. Clica em "Incluir Documento".
41. Seleciona "Externo", preenche dados (DataAtual, Nº Empenho, Seq. Liquidação, Id) e realiza upload.
42. Exclui o arquivo local após inserção no SEI.

## Passos Para Execução do Fluxo no Power Automate Desktop

**Pré-requisitos**  
- Power Automate Desktop instalado  
- Fluxo criado  
- Planilha auxiliar (`dados.xlsx`) configurada na variável `FilePath`  
- Siafi Terminal instalado  

**Fluxo de execução**  
1. Fazer download do Power Automate Desktop (site da Microsoft).  
2. Instalar e fazer login com conta corporativa (CAMG).  
3. Criar novo fluxo (caso não exista):  
   - Criar fluxo  
   - Nomear fluxo  
   - Criar subfluxos  
4. Colar os arquivos `Main.txt`, `Inicializar variáveis de Fluxo.txt`, `Executar Siafi Terminal.txt` e `Executar Siafi Web.txt` nos subfluxos correspondentes.  
5. Criar a planilha auxiliar com as colunas necessárias e definir seu caminho em `FilePath`.  
6. Executar fluxo.  
'''

