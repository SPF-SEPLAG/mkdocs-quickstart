## Objetivo

Este fluxo automatiza a leitura de notas fiscais (NFs) dentro de uma estrutura de pastas, percorrendo subpastas organizadas por unidades e processos. O objetivo é garantir que apenas as unidades MGS e PRODEMGE sejam processadas.

## Documentação

O fluxo é composto por um fluxo principal (Main) e dois subfluxos (Inicializar Variáveis de Fluxo e Ler NFs).

1. O fluxo principal (Main) inicializa as variáveis chamando o subfluxo Inicializar Variáveis de Fluxo.
2. Após a inicialização, ele obtém as subpastas dentro do diretório definido na variável FilePath. Essa variável contém o caminho onde os documentos a serem analisados estão armazenados. As subpastas encontradas são armazenadas na variável Unidades.
3. O fluxo então itera sobre cada subpasta armazenada na variável Unidades. Dentro de cada subpasta, ele obtém os processos existentes e os armazena na variável Processos.
4. Para cada Processo dentro de uma Unidade, o fluxo chama o subfluxo Ler NFs, responsável por realizar a leitura das notas fiscais.
5. Após a execução do subfluxo Ler NFs, são definidas variáveis contendo as informações extraídas das notas fiscais, tais como:
   - Processo SEI
   - Doc.SeiNF
   - Doc.SeiAteste
   - Doc.SeiConformidade
   - DtVencimento
   - InicioCompetencia
   - FimCompetencia
   - ValorNF
   - Banco (valor fixo)
   - Agência (valor fixo)
   - Conta (valor fixo)

   As variáveis Banco, Agência e Conta possuem valores fixos, enquanto as demais são extraídas com base em delimitadores de texto das notas fiscais.
6. Após a extração dos dados, o fluxo se anexa a uma aba aberta no navegador, onde a planilha CONTROLE DCF 2024.xlsx deve estar carregada na aba Notas Fiscais.
7. O fluxo exibe uma mensagem solicitando ao usuário que selecione a primeira célula vazia da planilha antes de continuar. Após essa seleção, ele utiliza a função Enviar Teclas para colar os valores extraídos diretamente na planilha do Excel.

## Passos Para Execução do Fluxo no Power Automate Desktop

**Caso já tenha o ambiente preparado (Power Automate instalado, fluxo criado e estrutura de pastas criada), pode-se passar para o passo 7.**

**Caso já tenha o ambiente preparado (Power Automate instalado e fluxo criado), pode-se passar para o passo 6.**

**Caso não tenha o ambiente preparado, seguir desde o primeiro passo.**

1. Fazer download da ferramenta no site da Microsoft (caso não tenha).
2. Após o download da ferramenta, realizar a instalação.
3. Após a instalação, a ferramenta solicitará que faça login com a sua conta corporativa (conta da própria CAMG).
4. Após a instalação e o login, caso o fluxo não esteja disponível em sua ferramenta, será necessário criá-lo. Para isso, siga os passos abaixo:
   - Criar Novo Fluxo
   - Nomear Novo Fluxo
   - Criar Subfluxos
5. Após criar o fluxo seguindo os passos indicados, você pode utilizar os arquivos Main.txt, Inicializar variáveis de fluxo.txt e Ler NFs.txt, para copiá-los e colá-los no Power Automate, criando os respectivos subfluxos.
6. Após a criação do fluxo, é necessário configurar a estrutura de pastas e subpastas para a leitura das notas fiscais. Para isso, acesse o diretório “Documentos” no seu computador e siga estas etapas:
   - Crie a pasta raiz onde as subpastas das unidades e processos serão armazenadas. Exemplo: Registrar NFs.
   - Dentro dela, crie as subpastas: MGS e PRODEMGE.
   - Armazene o caminho da pasta raiz na variável FilePath.
7. Após as subpastas serem criadas, será necessário fazer o download e descompactação dos processos do SEI onde estão as notas fiscais. Basta seguir os passos abaixo:
   - Fazer Login no SEI
   - Pesquisar Processo no SEI
   - Selecionar arquivo ZIP
   - Abrir tela de pesquisa de documentos
   - Escolher quais documentos estarão na pasta zipada: esses documentos serão as Notas Fiscais a serem extraídas, o documento de Ateste de Nota Fiscal e os Dados para Conformidade do mês corrente.
   - Gerar arquivo ZIP
   - Descompactar pasta zipada para subpasta referente a cada empresa
8. Agora, basta abrir a planilha CONTROLE DCF 2024.xlsx na aba Notas Fiscais (o fluxo está estruturado para que ela esteja aberta no navegador Microsoft Edge) e executar o fluxo.
9. Quando o fluxo iniciar sua execução, ele exibirá uma mensagem solicitando ao usuário que selecione a primeira linha vazia da planilha e pressione OK para que o fluxo realize a extração das notas e o preenchimento na planilha.
'''