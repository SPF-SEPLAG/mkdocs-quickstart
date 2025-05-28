# Documentação Rotina de Atualização e Tratamento de Dados para o Dashboard BI da DPO

## Introdução

Esse projeto foi desenvolvido com a intenção de automatizar o tratamento e a atualização dos dados do Dasbhoard de BI (Business Intelligence) da Diretoria de Planejamento e Orçamento (DPO) da Superintendência de Planejamento e Finanças (SPF) da Secretaria de Planejamento e Gestão de Minas Gerais (SEPLAG – MG).

Para isso é utilizada uma série de scripts desenvolvidos na linguagem de programação Python que realizam a extração dos dados pelo Sharepoint e o seu tratamento em sequência, por meio das bibliotecas msal, Office365-REST-Python-Client e pandas.

Além disso, para a automação desta rotina, é utilizado o Github Actions, que permite a execução automatizada de scripts de programação, funcionalidade presente no Github, plataforma para controle e versionamento de código, amplamente utilizada por desenvolvedores.

## Ferramentas necessárias:

- Python
- Github
- Vscode (recomendado)
- Gitbash (recomendado)

## Passo a passo:
**Observação: o repositório do github deve ser privado, pois o Script utiliza as credenciais da organização do usuário para gerar os Tokens de autenticação.**

**Além disso, o github não aceita push de arquivos com Token em repos públicos.**

## Setup inicial:

1. Instalar o python.

2. Criar repositório privado no github.

3. Criar repo local.

## Inicializar git:
- Criar o README.md;
- Criar .gitignore;
- Realizar primeiro commit;
- Alterar nome da branch para main;
- Criar origin e definir URL a partir do remote;
- Realizar o primeiro push;

## Comandos:

- git init
- git add .
- git commit -m "first commit"
- git branch -M main
- git remote add origin git push origin main

**Observação: alterar user e repo-name.**

- Adicionar scripts .py na raiz do repo local.

- Adicionar script .yml no repo local para execução do Actions do github.

**Criar na raiz do repo: .github/workflows/run-script.yml**

## Instalar ambiente virtual python:

### Criar ambiente virtual:
- python –m venv venv

### Ativar o ambiente virtual:

- source venv/Scripts/activate

### Instalar os requirements do projeto.

- pip install –r requirements.txt

### Commitar e fazer o push das alterações para o github.

- git add .
- git commit -m "mensagem do commit"
- git push origin main

**Observação: alterar a mensagem do commit.**

## Criar o Secret e o Personal Access Token (PAT) no Github para correta execução dos Actions do Github.

### Criar PAT:

 - Conta > Settings > Developer settings > Personal access tokens > Tokens (classic) > Generate new token > Generate new token (classic) > Note: gh-action-push > Select scopes: repo (marcar checkbox) > Generate token > Copiar o hash do token (sugestão: copiar e salvar para um txt, pois você não conseguirá ver ele novamente).

### Criar Secret:

- Repositório > Settings > Secrets and variables > Actions > New repository secret > Name: PERSONAL_TOKEN, Secret: colar a hash do PAT que você copiou > Add secret

## Execução manual

### De 3 em 3 meses será necessário executar os scripts manualmente para gerar um novo Refresh Token.

### Executar git pull do repo do github para baixar os arquivos mais recentes.

- git pull origin main

### Ativar o ambiente virtual python.

- source venv/Scripts/activate

### Executar o script de download.

- python download.py

### Executar os scripts de tratamento.

python process_data_num.py

python process_data_txt.py

### Commitar e fazer o push das alterações para o github.

- git add .
- git commit -m "mensagem do commit"
- git push origin main

**Observação: alterar a mensagem do commit.**

# Documentação:

## download.py

Chama a função download_sharepoint_file(), que recebe os seguintes parâmetros:

- base_url: URL da equipe no Sharepoint.

- folder_path: o caminho da pasta do arquivo que será baixado.

- file_name: o nome do arquivo para baixar (nome do arquivo no Sharepoint).

- local_filename: o novo nome do arquivo que será salvo (nome do arquivo na máquina local).

Além disso, a função download_sharepoint_file() chama por baixo dos panos a função get_sharepoint_token(), presente no script sharepoint_utils.py.

Pode-se dizer que essa função é a mais importante para o funcionamento do projeto como um todo, pois é através dela que é realizada a autenticação do usuário no Sharepoint e a geração do Refresh Token e do Access Token, conforme será detalhado abaixo.

## sharepoint_utils.py

### get_sharepoint_token()

É a função mais importante do projeto: a partir dela é gerado o arquivo cache que armazenará o Access Token e o Refresh Token que permitirão a autenticação do usuário e realização do download ou upload do arquivo para o Sharepoint que será tratado posteriormente para extração dos dados e informações relevantes.

Na primeira execução, inicialmente é gerado o cache em que serão armazenadas as informações de acesso do usuário.

O app é então carregado com, além das credenciais da organização, as informações presentes no cache, se existirem. É isso que permite a autenticação do usuário silenciosamente (sem interação com o navegador) em execuções futuras, conforme será detalhado abaixo.

Seguindo a lógica anterior, a próxima etapa é checar se a credencial do usuário corresponde a algum dos usuários listados no cache da aplicação. Ou seja, se aquele usuário tentando realizar o download ou upload do arquivo é autenticado e possui algum Token correspondente. Se for, o Token será gerado silenciosamente, conforme detalhado abaixo. Na primeira execução essa variável accounts será vazia.

Em seguida, ao não conseguir gerar o Token silenciosamente, o navegador é aberto e é pedido ao usuário para logar na sua conta e, portanto, gerar o Token interativamente.

Essas novas informações presentes no cache que foi gerado previamente são então salvas no arquivo msal_cache.bin (se ele não existia previamente na pasta do projeto ele é criado aqui).

Dessa forma, a partir da primeira execução e geração do arquivo de cache, com as informações de acesso, incluindo o Token, nas execuções posteriores, o Token poderá ser adquirido silenciosamente, não sendo necessário que o usuário realize qualquer tipo de interação com o navegador: o Token será apenas lido do arquivo com o cache (msal_cache.bin).

Ao final, a função retorna a variável result. É essa variável que conterá as informações relativas ao Token.

Por fim, deve-se detalhar a diferença entre Access Token e Refresh Token e como eles estão presentes neste projeto. O Access Token é o que de fato é utilizado para autenticação de usuário. O problema do Access Token é que ele expira após cerca de 1 hora, ou seja, possui um prazo de validade muito curto. Como solução para esse problema, existe o Refresh Token, que é utilizado para gerar um novo Access Token recorrentemente, tendo em vista que possui um prazo de validade mais longo (por volta de 3 meses num geral).

No contexto desse projeto, quando o usuário executa o script pela primeira vez, ao realizar o login no navegador, são gerados interativamente o Access Token e o Refresh Token, na chamada da função acquire_token_interactive(). Ao final do código, conforme detalhado anteriormente, esses Tokens são salvos no cache e no arquivo correspondente, msal_cache.bin.

A partir daí, em novas execuções o Token será gerado silenciosamente por meio da função acquire_token_silent(). Essa função lê o Access Token presente no arquivo de cache e passa o valor para a variável; no entanto, caso ele tenha expirado, ela utiliza o Refresh Token e gera um novo Access Token, permitindo novamente a autenticação.

Na prática, pela possibilidade de geração do Access Token silenciosamente a partir do Refresh Token, pode-se dizer que o uso desse script permite o desenvolvimento de rotinas de automação. No entanto, a cada 3 meses será necessária a execução do script manualmente por parte do usuário, para geração de um novo Refresh Token.

