# Fluxo Power Automate – Cadastro de Empenho no Portal Compras / Sistema DES / SEI

## Objetivo

Este fluxo automatiza o preenchimento dos dados de empenho no portal de execução de despesas do governo de Minas Gerais, realiza o envio dos dados no terminal do sistema DES, integra informações com o SEI e realiza verificações e preenchimentos automáticos com base nos dados extraídos da interface web e manipulação via Excel.

## Documentação

1. Abertura do portal e captura de dados
- O fluxo inicia conectando ao navegador Google Chrome já aberto na aba com o título “Dados de identificação” e armazena a instância como PortalCompras.
- Captura elementos como UE, UO, Dotação, UPG, CNPJ, Tipo, Valor, Ordenador, Número da Especificação, Dados da Compra, e Processo SEI por meio da extração de texto direto da interface com GetDetailsOfElement.

2. Envio de dados para o Terminal DES
- Utiliza o título da janela do emulador 3270 para envio de dados via comando SendKeys, simulando o uso do teclado no sistema legado.
- Envia comandos sequenciais contendo códigos de UE, UO, Dotação, UPG, CNPJ, Tipo, e demais informações financeiras.

3. Manipulação de dados no Excel
- Abre uma planilha Excel e escreve o valor capturado da interface na célula A1 e calcula um valor multiplicado por 100 na célula A2.
- Lê esse valor de volta e envia ao sistema DES.

4. Tratamento de Ordenador
- Verifica a posição do hífen no texto completo do nome do ordenador e extrai corretamente o nome a ser enviado ao sistema.

5. Envio de dados do processo de compra
- Extrai informações do processo de compra como Unidade, Número e Ano, e envia para o sistema.
- Envia também o número da especificação e finaliza com comandos de confirmação.

6. Tratamento Condicional com Grupo
- Lê o Grupo orçamentário da Dotação.
- Caso o grupo seja 3, envia sequência específica de comandos; caso contrário, outra sequência é utilizada.

7. Descrição Final do Empenho
- Captura parte do texto do serviço e insere a descrição completa do empenho no campo apropriado.

8. Integração com SEI
- Anexa à aba do Chrome com título do processo SEI.
- Extrai o número do processo do título da página, envia ao terminal e recorta o texto relevante.
- Obtém o número do empenho a partir da área de transferência e envia ao sistema.

9. Finalização
- Realiza comandos finais de navegação e limpeza de dados.
- Etapas com o sistema SIAFI estão desativadas, indicando que são opcionais ou descontinuadas.

## Passos Para Execução do Fluxo no Power Automate Desktop

1. Ter o Power Automate Desktop instalado.

2. Criar um novo fluxo com o nome sugerido “Cadastro de Empenho”.

3. Criar os subfluxos ou blocos de ação conforme os comandos presentes no arquivo Main.txt.

4. Preparar a planilha Excel (caso ainda não exista) para receber e manipular valores.

5. Abrir o terminal DES com a janela ativa e logado.

6. Abrir o navegador Google Chrome com as abas do portal de execução de despesas e SEI já carregadas.

7. Executar o fluxo.

