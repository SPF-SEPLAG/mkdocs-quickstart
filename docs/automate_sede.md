# Fluxo Extração de Ordens de Pagamento

## Objetivo

Este fluxo foi desenvolvido para uma demanda pontual da SEDE, onde deveriam ser extraídas cerca de 4400 ordens de pagamento referentes aos anos de 2018 a 2020.

## Documentação

O fluxo é composto por um fluxo principal (Main) e dois subfluxos:

- **Inicializar Variáveis de Fluxo**: define as variáveis utilizadas ao longo do processo.  
- **Navegador**: contém as ações realizadas no portal do SIAFI.

### Variáveis definidas em Inicializar Variáveis de Fluxo

- `Ano`: valor inicial 2018, correspondente à primeira aba da planilha e também usado para login no SIAFI.  
- `Pasta`: caminho do diretório onde as ordens de pagamento extraídas serão salvas.  
- `PrimeiraLinha`: número da primeira linha da planilha a partir da qual os dados serão lidos.  
- `LoginSIAFI`: usuário para login no SIAFI.  
- `SenhaSIAFI`: senha para login no SIAFI.  
- `UnidadeExecutoraSIAFI`: unidade executora usada no login do SIAFI.

### Fluxo Principal (Main)

1. Chama o subfluxo **Inicializar Variáveis de Fluxo**.  
2. Inicia o loop **LoopAbas**, que percorre as abas de 2018 a 2020 na planilha:  
   1. Atualiza a aba ativa usando a variável `Ano`.  
   2. Executa o subfluxo **Navegador**.  
   3. Obtém a primeira linha vazia da coluna G (ordens de pagamento) e armazena em `PrimeiraLinhaVazia`.  
3. Inicia o loop **LoopDownload**, que percorre as linhas da planilha de `PrimeiraLinha` até `PrimeiraLinhaVazia – 1`:  
   1. Lê o valor da coluna G na linha atual e armazena em `OrdemPagamento`.  
   2. Se `OrdemPagamento` estiver vazia, sai para o **LoopAbas** (próxima aba).  
   3. Define `NomeArquivo` igual a `OrdemPagamento`.  
   4. No portal SIAFI, preenche o campo de pesquisa com `OrdemPagamento` e clica em “Confirmar”.  
   5. Captura a janela do documento baixado e armazena em `JanelaDOC`.  
   6. Aguarda 2 segundos para carregar, envia TAB 8 vezes e ENTER para iniciar o download.  
   7. Aguarda 2 segundos e envia o caminho de destino: `%Pasta%\{Ano}\{OrdemPagamento}`.  
   8. Fecha a janela do documento, volta ao portal e aguarda 5 segundos.  
   9. Incrementa `PrimeiraLinha` em 1.  
4. Após finalizar o **LoopDownload**, clica em “Voltar” e em “Sair” para encerrar a sessão.  
5. Incrementa `Ano` em 1 e repete até 2020.  

## Passos para Execução no Power Automate Desktop

> **Caso já tenha o ambiente preparado** (Power Automate instalado, fluxo criado e estrutura de pastas):  
> - Pule para o passo 7.  
> **Caso já tenha Power Automate instalado e fluxo criado**:  
> - Pule para o passo 6.  
> **Caso não tenha o ambiente preparado**:  
> - Siga desde o passo 1.

1. Faça o download do Power Automate Desktop no site da Microsoft.  
2. Instale e faça login com sua conta corporativa (CAMG).  
3. Se necessário, crie um novo fluxo:  
   - Criar fluxo  
   - Nomear fluxo  
   - Criar subfluxos  
4. Copie e cole os arquivos `Main.txt`, `Inicializar variáveis de fluxo.txt` e `Navegador.txt` nos subfluxos correspondentes.  
5. Crie a estrutura de pastas:  
   - Pasta raiz: **Ordens de Pagamento**  
   - Subpastas: **2018**, **2019** e **2020**  
6. Defina as variáveis no Power Automate Desktop:  
   - `Pasta`: caminho da pasta raiz  
   - `PlanilhaExcel`: caminho do arquivo Excel com dados  
7. Execute o fluxo.
