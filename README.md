# OAE Engineering Tools

Este repositório inclui utilitários para gerenciar entregas de projetos. Uma interface web simples baseada em Flask pode ser usada para acionar as rotinas de entrega e testar a geração de planilhas, e um aplicativo baseado em Tkinter ajuda a organizar e verificar os entregáveis do projeto.

-----

## Configurando Seu Ambiente Python (Setting Up Your Python Environment)

Antes de começar, é necessário ter instalado em sua máquina, o python, idealmente a versão 3.12.3 que foi a utilizada no código, caso opte por usar de uma versão mais nova será necessário também atualizar as demais bibliotecas utilizadas no código. Caso não tenho o python instalado você pode baixá-lo no:

 [site oficial do Python](https://www.python.org/downloads/).

 [Vídeo Tutotial para instalação e configuração das variáveis de ambiente](https://www.youtube.com/watch?v=WgFqLVRh0Y0).

 Feito isso, podemos partir para as instalações de bibliotecas:

 -----

### 1\. Instalar `pip` (Install `pip`)

`pip` é o instalador de pacotes para Python. Geralmente, ele já vem pré-instalado com o Python 3. Para verificar se você tem o `pip` e instalá-lo caso não tenha, abra seu terminal ou prompt de comando e execute:

```bash
pip --version
```

### 2\. Criar um Ambiente Virtual (Create a Virtual Environment)

É altamente recomendável usar um ambiente virtual para gerenciar as dependências do projeto. Isso isola as dependências do projeto de outros projetos Python em seu sistema. No caso, trazendo um contexto venv ou ambiente virtual é uma ferramenta integrada ao Python que permite criar ambientes virtuais isolados. Esses ambientes são importantes porque permitem que cada projeto tenha suas próprias dependências, sem interferir em outros projetos ou na instalação global do Python. Tendo isso entendido, navegue até o diretório raiz deste repositório (o arquivo `gerenciar_entregas.py` deve estar neste diretório) e execute:

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

## 4\. Instalar Pacotes Necessários (Install Required Packages)

Com o `requirements.txt` no lugar e seu ambiente virtual ativado, instale as dependências executando:

```bash
pip install -r requirements.txt
```

Este comando fará o download e instalará todos os pacotes listados.

## 5\. Configurando Variáveis de Ambiente (Configuring Environment Variables)

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
## 6\. Rodando o Aplicativo Principal (Interface Tkinter) (Running the Main Application (Tkinter Interface))

Após tudo o que foi feito, podemos rodar nossa aplicação para tanto, para executar o script principal, certifique-se de que seu ambiente virtual esteja ativado e todas as dependências instaladas. Em seguida, a partir do diretório raiz do repositório:

```bash
python gerenciar_entregas.py
```

Uma interface gráfica irá guiá-lo na seleção do projeto, escolha de arquivos e atualização da planilha GRD.

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

