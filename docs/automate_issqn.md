# Fluxo Power Automate – ISSQN

## Objetivo

Este fluxo automatiza a emissão das guias de ISSQN, a inserção de guias de ISSQN no SEI e o registro de serviços tomados.

## Documentação

O fluxo é composto por 1 fluxo principal (Main) e 6 subfluxos (Inicializa Fluxo, Registrar Serviço, Emitir Guia, Inserir Guia no SEI, Inserir Valor NF, Registrar Serviço Tomado).

1. O fluxo principal (Main) inicia definindo a variável `ListaProcessos` e, conforme escolha do usuário, chama o subfluxo **Registrar Serviço** ou o subfluxo **Emitir Guia**.
2. Caso o usuário escolha executar o subfluxo **Registrar Serviço**, serão chamados dois subfluxos: **Inicializa Fluxo** e **Registrar Serviço Tomado**.
3. Caso o usuário escolha executar o subfluxo **Emitir Guia**, serão chamados dois subfluxos: **Inicializa Fluxo** e **Emitir Guia**.
4. Para realizar o registro de serviços, o fluxo segue o seguinte passo a passo:
   - Executa o subfluxo **Inicializa Fluxo**.
   - O subfluxo **Inicializa Fluxo** define as variáveis de fluxo, recebendo o caminho da pasta “Documentos”.
   - Inicia o Excel na planilha auxiliar e armazena a instância na variável `Planilha`.
   - Obtém a primeira coluna e primeira linha livre, armazenando em `ColunaLivre` e `LinhaLivre`.
   - Lê todo o intervalo de dados preenchidos na planilha.
   - Fecha o Excel e coloca em foco a janela da DES.
   - Executa o subfluxo **Registrar Serviço Tomado**, que funciona como um loop, iterando cada linha na variável `DadosPlanilha`.
   - Chama o subfluxo **Inicializa Loop**, onde são inicializadas as variáveis com os dados de cada linha.
   - Define a variável `DataEmissaoOP` com o valor do campo “Data de emissão OP” da coleção `DadosPlanilha`, convertendo-o para Data.
   - Valida se o credor é MGS ou PRODEMGE e prossegue apenas nesses casos.
   - Envia as informações de emissão, processo SEI, credor e demais campos exigidos.
5. Após criar o fluxo seguindo os passos indicados, utilize os arquivos `Main.txt`, `Inicializar variáveis de fluxo.txt`, `Emitir Guia.txt` e `Registrar Serviço Tomado.txt` para copiá-los e colá-los no Power Automate, criando os subfluxos correspondentes.
6. Para o fluxo ser executado, crie uma planilha auxiliar (`dados.xlsx`) no caminho `%FilePathFluxo%\\dados.xlsx`.
7. A planilha auxiliar deve conter as seguintes colunas:
   *(defina as colunas conforme o layout de entrada dos dados)*
8. Para utilizar o fluxo, é necessário ter o terminal da DES instalado. A instalação pode ser solicitada ao setor de T.I.
9. Após o terminal da DES estar instalado, realize o login no sistema.
10. Com o ambiente preparado, a DES deve permanecer na tela inicial após o login.
11. Em seguida, execute normalmente o fluxo.

## Passos Para Execução do Fluxo no Power Automate Desktop

**Caso já tenha o ambiente preparado** (Power Automate instalado, fluxo criado e terminal da DES instalado), pode pular para o passo 6.  
**Caso já tenha o Power Automate instalado e o fluxo criado**, pode pular para o passo 5.  
**Caso não tenha o ambiente preparado**, siga desde o primeiro passo.

1. Fazer download do Power Automate Desktop no site da Microsoft.  
2. Instalar e realizar login com sua conta corporativa (CAMG).  
3. Criar novo fluxo (caso não exista):  
   - Criar fluxo  
   - Nomear fluxo  
   - Criar subfluxos  
4. Copiar e colar os arquivos `Main.txt`, `Inicializar variáveis de fluxo.txt`, `Emitir Guia.txt` e `Registrar Serviço Tomado.txt` nos subfluxos correspondentes.  
5. Criar a planilha auxiliar (`dados.xlsx`) com as colunas necessárias e definir seu caminho na variável `FilePathFluxo`.  
6. Executar o fluxo.  

