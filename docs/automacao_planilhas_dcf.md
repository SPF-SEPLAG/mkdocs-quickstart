# Documentação do Fluxo de Automação de Planilhas

## Objetivo
Este aplicativo em Python automatiza o processo de downloads de planilhas do SharePoint, atualização de dados em arquivos de controle e oferece uma interface gráfica para execução das operações.  
O fluxo principal consiste em três etapas:
1. **Download**: obtém os arquivos fonte do SharePoint.
2. **Atualização**: faz backup dos arquivos de destino e atualiza abas específicas com dados de origem, preservando formatação e células mescladas.
3. **Upload**: (em placeholder) envia os arquivos atualizados ao SharePoint.

## Pré-requisitos
- Python 3.7 ou superior
- Bibliotecas Python:
  - `pandas`
  - `openpyxl`
  - `tkinter` (incluso na maioria das instalações Python)
  - `msal`
  - `requests`

Instale as dependências com:
```bash
pip install pandas openpyxl msal requests
```

## Estrutura de Arquivos
```
.
├── att.py                # Aplicativo GUI principal
├── download.py           # Script de download simples
├── sharepoint_utils.py   # Funções de autenticação e download/upload no SharePoint
├── treatment.py          # Lógica de leitura, limpeza e atualização das planilhas
└── arquivos/             # Pasta onde os arquivos são salvos localmente
    ├── Dados dos empenhos.xlsx
    ├── Dados OP BO.xlsx
    ├── CONTROLE DCF 2024 - Copia.xlsx
    └── CONTROLE DCF 2025 - Copia.xlsx
```

## Módulos e Funções

###  sharepoint_utils.py
- `get_sharepoint_token()`  
  Autentica no Azure AD via MSAL e retorna um token de acesso.
- `download_sharepoint_file(base_url, folder_path, file_name, local_filename)`  
  Baixa um arquivo do SharePoint usando o token.
- `upload_sharepoint_file(base_url, folder_path, file_name, local_filename)`  
  Envia (ou sobrescreve) um arquivo no SharePoint (função implementada como futura).

###  download.py
Script de uso direto para baixar os arquivos fixos:
```python
from sharepoint_utils import download_sharepoint_file

# Configuração de URL e pasta
base_url = "https://cecad365.sharepoint.com/sites/AutomatizacoesSPF"
folder_path = "/sites/AutomatizacoesSPF/Documentos Compartilhados/General/teste automação vitor"

# Chamadas de download
download_sharepoint_file(base_url, folder_path, "Dados OP BO.xlsx", "Dados OP BO.xlsx")
# ...
```

### treatment.py
Contém a lógica de atualização das planilhas:
1. Carrega planilhas de origem e destino com `openpyxl`.
2. Lê dados das abas de origem (`Dados dos Empenhos`, `Relatório 1`) via pandas.
3. Limpa valores antigos nas abas de destino.
4. Copia os valores, preservando formatação e células mescladas.
5. Salva as planilhas de destino.

### att.py
Interface gráfica com `tkinter`:
- Classe `AutomacaoApp` com botões:
  - **Download** → `executar_download()`
  - **Atualização** → `atualizar_planilhas()`
  - **Upload** → `executar_upload()` (placeholder)
  - **Executar Tudo** → executa as três etapas em sequência
- Redireciona `stdout` para uma caixa de texto para exibir logs em tempo real.

## Fluxo de Execução

1. **Inicialização**  
   O usuário roda `python att.py` e a janela é exibida.
2. **Download**  
   Cria pasta `arquivos/` e baixa os arquivos configurados.
3. **Backup e Atualização**  
   - Cria backups em `backup_planilhas/` com timestamp.  
   - Lê dados das planilhas de origem (`Dados dos Empenhos.xlsx`, `Dados OP BO.xlsx`).  
   - Atualiza abas correspondentes em `CONTROLE DCF 2024 - Copia.xlsx` e `CONTROLE DCF 2025 - Copia.xlsx`.
4. **Upload**  
   Placeholder que poderá ser substituído pela função real de upload no SharePoint.

## Uso

Execute:
```bash
python att.py
```
Na interface, clique conforme desejado:
- **Download**: baixa fontes.
- **Atualização**: processa e atualiza planilhas.
- **Upload**: envia (quando implementado).
- **Executar Tudo**: realiza todas as etapas automaticamente.

## Logs e Tratamento de Erros
- Mensagens de status e erros são exibidas na área de log da GUI.
- Cada função envolve `try/except` para capturar exceções e registrar erros.

## Personalização
- Edite os nomes de arquivo e caminhos em constantes no início de `att.py`.
- Ajuste `SHAREPOINT_BASE_URL` e `SHAREPOINT_FOLDER_PATH` em `sharepoint_utils.py`.
- Implemente `executar_upload()` com lógica de upload real conforme necessidades.

## Observações
- **Backup**: arquiva versões anteriores para segurança.
- **Formatação**: a função de atualização preserva estilos e mesclagens.
- **Segurança**: mantenha o cache do MSAL (`msal_cache.bin`) protegido.
