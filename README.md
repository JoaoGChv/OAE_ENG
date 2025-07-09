# OAE Engineering Tools

Este repositório inclui utilitários para gerenciar entregas de projetos. Uma interface web simples baseada em Flask pode ser usada para acionar as rotinas de entrega e testar a geração de planilhas, e um aplicativo baseado em Tkinter ajuda a organizar e verificar os entregáveis do projeto.

-----

## Configurando Seu Ambiente Python (Setting Up Your Python Environment)

Antes de começar, certifique-se de ter o **Python 3** instalado em seu sistema. Você pode baixá-lo no [site oficial do Python](https://www.python.org/downloads/).

### 1\. Instalar `pip` (Install `pip`)

`pip` é o instalador de pacotes para Python. Geralmente, ele já vem pré-instalado com o Python 3. Para verificar se você tem o `pip` e instalá-lo caso não tenha, abra seu terminal ou prompt de comando e execute:

```bash
python -m ensurepip --default-pip
```

### 2\. Criar um Ambiente Virtual (Create a Virtual Environment)

É altamente recomendável usar um ambiente virtual para gerenciar as dependências do projeto. Isso isola as dependências do projeto de outros projetos Python em seu sistema.

Navegue até o diretório raiz deste repositório (o arquivo `gerenciar_entregas.py` deve estar neste diretório) e execute:

```bash
python -m venv myvenv
```

Este comando cria um novo diretório chamado `myvenv` (você pode escolher um nome diferente se preferir) que contém um ambiente Python novo.

### 3\. Ativar o Ambiente Virtual (Activate the Virtual Environment)

Antes de instalar as dependências ou executar os aplicativos, você precisa ativar seu ambiente virtual.

**No Windows (PowerShell):**

```bash
.\myvenv\Scripts\Activate.ps1
```

**No Windows (Prompt de Comando):**

```bash
.\myvenv\Scripts\activate.bat
```

**No macOS/Linux:**

```bash
source myvenv/bin/activate
```

Uma vez ativado, seu prompt de terminal geralmente mostrará `(myvenv)` indicando que você está operando dentro do ambiente virtual.

-----

## Instalando Dependências (Installing Dependencies)

Após ativar seu ambiente virtual, você pode instalar todos os pacotes necessários usando o arquivo `requirements.txt`.

### 1\. Baixar `requirements.txt` (Download `requirements.txt`)

Se você ainda não tiver um arquivo `requirements.txt`, precisará criar um. Este arquivo lista todos os pacotes Python e suas versões necessárias para o projeto. Para este projeto, um `requirements.txt` básico seria assim:

```
flask
openpyxl
send2trash
```

Salve este conteúdo em um arquivo chamado `requirements.txt` no diretório raiz do seu projeto.

### 2\. Instalar Pacotes Necessários (Install Required Packages)

Com o `requirements.txt` no lugar e seu ambiente virtual ativado, instale as dependências executando:

```bash
pip install -r requirements.txt
```

Este comando fará o download e instalará todos os pacotes listados.

-----

## Configurando Variáveis de Ambiente (Configuring Environment Variables)

Os locais padrão para os arquivos JSON usados pelo aplicativo podem ser substituídos usando variáveis de ambiente. Essas variáveis especificam os caminhos para os arquivos de configuração que o script usa.

Para definir uma variável de ambiente, você normalmente faria o seguinte:

**No Windows (Prompt de Comando):**

```bash
set OAE_PROJETOS_JSON="C:\caminho\para\seu\diretorios_projetos.json"
set OAE_NOMENCLATURAS_JSON="C:\caminho\para\seu\nomenclaturas.json"
set OAE_ULTIMO_DIR_JSON="C:\caminho\para\seu\ultimo_diretorio_arqs.json"
```

**No macOS/Linux:**

```bash
export OAE_PROJETOS_JSON="/caminho/para/seu/diretorios_projetos.json"
export OAE_NOMENCLATURAS_JSON="/caminho/para/seu/nomenclaturas.json"
export OAE_ULTIMO_DIR_JSON="/caminho/para/seu/ultimo_diretorio_arqs.json"
```

O script define esses padrões internamente se as variáveis de ambiente não forem definidas:

```python
PROJETOS_JSON = _resolve_json_path(
    "OAE_PROJETOS_JSON",
    r"G:\Drives compartilhados\OAE-JSONS\diretorios_projetos.json",
)
NOMENCLATURAS_JSON = _resolve_json_path(
    "OAE_NOMENCLATURAS_JSON",
    r"G:\Drives compartilhados\OAE-JSONS\nomenclaturas.json",
)
ARQ_ULTIMO_DIR = _resolve_json_path(
    "OAE_ULTIMO_DIR_JSON",
    r"G:\Drives compartilhados\OAE-JSONS\ultimo_diretorio_arqs.json",
)
```

-----

## Rodando o Servidor Web (Running the Web Server)

Este repositório inclui uma interface web simples baseada em Flask que pode ser usada para acionar rotinas de entrega e testar a geração de planilhas.

Certifique-se de que seu ambiente virtual esteja ativado e as dependências instaladas.

1.  **Inicie o servidor** a partir da raiz do repositório (onde `web_app/app.py` está localizado, ou simplesmente do diretório onde `gerenciar_entregas.py` reside):

    ```bash
    flask --app web_app.app run
    ```

    O aplicativo estará disponível em `http://127.0.0.1:5000/`.

Acesse `http://127.0.0.1:5000/` e escolha **New Delivery** (Nova Entrega) para selecionar um projeto e fazer upload de arquivos. Os arquivos carregados são copiados para uma nova pasta de entrega, e uma planilha GRD é gerada automaticamente. O link de demonstração ainda cria uma planilha fictícia em seu diretório temporário.

**Páginas adicionais (Additional pages):**

  * `/history?folder=<ENTREGAS_PATH>&tipo=AP` – visualize os arquivos de entrega mais recentes para uma disciplina.
  * `/nomenclature?folder=<PASTA>&num=<PROJETO>` – verifique os tokens de nomenclatura dos arquivos usando as regras de nomenclatura JSON.

-----

## Rodando o Aplicativo Principal (Interface Tkinter) (Running the Main Application (Tkinter Interface))

`gerenciar_entregas.py` é uma ferramenta baseada em Tkinter que ajuda a organizar e verificar os entregáveis do projeto. Ele lê informações do projeto e de nomenclatura de arquivos JSON, auxilia na seleção de arquivos para cada entrega e gera ou atualiza um GRD (`GRD_ENTREGAS.xlsx`).

Para executar o script principal, certifique-se de que seu ambiente virtual esteja ativado e todas as dependências instaladas. Em seguida, a partir do diretório raiz do repositório:

```bash
python gerenciar_entregas.py
```

Uma interface gráfica irá guiá-lo na seleção do projeto, escolha de arquivos e atualização da planilha GRD.

-----

## Gerando um Executável (.exe) (Generating an Executable (.exe))

Você pode converter o script `gerenciar_entregas.py` em um arquivo executável autônomo (`.exe`) usando o PyInstaller.

1.  **Instalar PyInstaller:**

    ```bash
    pip install pyinstaller
    ```

2.  **Gerar o executável:**

    Navegue até o diretório raiz do repositório em seu terminal e execute:

    ```bash
    pyinstaller --onefile --windowed gerenciar_entregas.py
    ```

    O executável será criado no diretório `dist/`.

Certifique-se de ter instalado previamente todas as dependências listadas em `requirements.txt` para que a criação do executável ocorra sem erros.

-----