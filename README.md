# Repositório da Aplicação TKinter para gerenciamento de arquivos da OAE - ENG.

-----

*Está é uma interface Tkinter que faz o gerenciamento de arquivos para um escritório de projetos em Engenharia. O projeto tem especificações próprias mas é modular o suficiente para que possa ser aplicado em outros contextos. Ele permite:*

  - Seleção de Projetos: Através de uma interface gráfica, o usuário pode selecionar um projeto a partir de uma lista carregada de um arquivo JSON.
  - Seleção de Disciplinas: Dentro de um projeto, o usuário escolhe uma disciplina específica para a qual os arquivos serão entregues.
  - Seleção de Arquivos: O usuário seleciona os arquivos relevantes para a entrega dentro da pasta de entregas da disciplina escolhida.
  - Extração de Dados: O código extrai informações relevantes dos nomes dos arquivos, como status, fase, tipo e revisão.
  - Validação de Nomenclatura: Uma interface permite ao usuário revisar e corrigir a nomenclatura dos arquivos com base em regras predefinidas.
  - Verificação de Revisão: O sistema identifica e separa arquivos revisados e obsoletos com base na nomenclatura.
  - Organização de Arquivos: Os arquivos são movidos para pastas "Revisados" e "Obsoletos" para melhor organização.

-----

# Para conseguir rodar a aplicação em qualquer máquina, faça o seguinte processo:

# OAE Engineering Tools

Este repositório inclui utilitários para gerenciar entregas de projetos. Uma interface web simples baseada em Flask pode ser usada para acionar as rotinas de entrega e testar a geração de planilhas, e um aplicativo baseado em Tkinter ajuda a organizar e verificar os entregáveis do projeto.

-----

## 1\. Configurando o Git e o GitHub

Para colaborar com o projeto, gerenciar versões do seu código e sincronizar seu trabalho, é essencial usar o Git e o GitHub.

### 1.1. Instalar o Git

O primeiro passo é instalar o Git em sua máquina. O Git é um sistema de controle de versão distribuído que rastreia as alterações no código-fonte durante o desenvolvimento de software.

  - **Download:** Baixe o Git no [site oficial do Git](https://git-scm.com/downloads).
  - **Instalação:** Siga as instruções do instalador para o seu sistema operacional (Windows, macOS ou Linux). Na maioria dos casos, as configurações padrão são suficientes.

Para verificar se a instalação foi bem-sucedida, abra seu terminal ou prompt de comando e execute:

```bash
git --version
```

Você deverá ver a versão do Git instalada.

### 1.2. Configurar o Git (git config)

Após a instalação, configure seu nome de usuário e e-mail. Isso é importante porque cada "commit" (salvamento) que você fizer usará essas informações.

Abra o terminal e execute os seguintes comandos, substituindo "Seu Nome" e "seu.email@example.com" pelos seus dados:

```bash
git config --global user.name "Seu Nome"
git config --global user.email "seu.email@example.com"
```

### 1.3. Conectar o Git à sua conta do GitHub

Para conectar seu computador ao GitHub, você precisa de uma forma de autenticação segura. A maneira mais comum e segura é usando SSH.

1.  **Crie uma conta no GitHub:** Se você ainda não tiver uma, crie uma conta em [github.com](https://github.com).
2.  **Gere uma chave SSH:** Siga o [guia oficial do GitHub](https://docs.github.com/pt/authentication/connecting-to-github-with-ssh/generating-a-new-ssh-key-and-adding-it-to-the-ssh-agent) para gerar uma chave SSH e adicioná-la ao agente SSH.
3.  **Adicione a chave SSH à sua conta do GitHub:** Copie sua chave SSH pública e adicione-a às configurações da sua conta no GitHub, seguindo [este guia](https://docs.github.com/pt/authentication/connecting-to-github-with-ssh/adding-a-new-ssh-key-to-your-github-account).

### 1.4. Como Subir Projetos para um Repositório

Para enviar um projeto local para um novo repositório no GitHub:

1.  **Crie um novo repositório no GitHub:** Vá ao GitHub e crie um novo repositório. Não o inicialize com um arquivo README, .gitignore ou licença.
   
2.  **Inicialize o repositório local:** Navegue até a pasta do seu projeto no terminal e execute:
   
    ```bash
    git init
    ```
    
3.  **Adicione os arquivos ao stage:** Para adicionar todos os arquivos do projeto para serem monitorados pelo Git, use:
   
    ```bash
    git add .
    ```
  
4.  **Faça o commit dos arquivos:** O "commit" é como um instantâneo do seu projeto. Salve suas alterações com uma mensagem descritiva:
   
    ```bash
    git commit -m "Primeiro commit: início do projeto"
    ```
    
5.  **Conecte seu repositório local ao repositório remoto (GitHub):**
   
    ```bash
    git remote add origin URL_DO_SEU_REPOSITORIO.git
    ```
    
    *Substitua `URL_DO_SEU_REPOSITORIO.git` pela URL que você copiou do seu repositório no GitHub.*
    
6.  **Envie suas alterações (push):**

    ```bash
    git push -u origin main
    ```
    
    *Se sua branch principal não se chamar `main`, substitua pelo nome correto (ex: `master`).*

### 1.5. Como Clonar um Repositório (git clone)

Para baixar uma cópia de um projeto que já existe no GitHub (como este), você usa o `git clone`.

1.  **Copie a URL do repositório:** No GitHub, clique no botão verde "Code" e copie a URL (HTTPS ou SSH).
2.  **Clone o repositório:** Abra o terminal, navegue até o diretório onde deseja salvar o projeto e execute:
   
    ```bash
    git clone URL_DO_REPOSITORIO_COPIADA
    ```

Isso criará uma pasta com o nome do repositório contendo todos os arquivos do projeto.

-----

## 2\. Configurando Seu Ambiente Python (Setting Up Your Python Environment)

Antes de começar, é necessário ter instalado em sua máquina, o python, idealmente a versão 3.12.3 que foi a utilizada no código, caso opte por usar de uma versão mais nova será necessário também atualizar as demais bibliotecas utilizadas no código. Caso não tenho o python instalado você pode baixá-lo no:

[site oficial do Python](https://www.python.org/downloads/).

[Vídeo Tutorial para instalação e configuração das variáveis de ambiente](https://www.youtube.com/watch?v=WgFqLVRh0Y0).

Feito isso, podemos partir para as instalações de bibliotecas:

-----

### 2.1. Instalar `pip` (Install `pip`)

`pip` é o instalador de pacotes para Python. Geralmente, ele já vem pré-instalado com o Python 3. Para verificar se você tem o `pip` e instalá-lo caso não tenha, abra seu terminal ou prompt de comando e execute:

```bash
pip --version
```

### 2.2. Criar um Ambiente Virtual (Create a Virtual Environment)

É altamente recomendável usar um ambiente virtual para gerenciar as dependências do projeto. Isso isola as dependências do projeto de outros projetos Python em seu sistema. No caso, trazendo um contexto venv ou ambiente virtual é uma ferramenta integrada ao Python que permite criar ambientes virtuais isolados. Esses ambientes são importantes porque permitem que cada projeto tenha suas próprias dependências, sem interferir em outros projetos ou na instalação global do Python. Tendo isso entendido, navegue até o diretório raiz deste repositório (o arquivo `gerenciar_entregas.py` deve estar neste diretório) e execute:

```bash
python -m venv myvenv
```

Este comando cria um novo diretório chamado `myvenv` (você pode escolher um nome diferente se preferir) que contém um ambiente Python novo.

### 2.3. Ativar o Ambiente Virtual (Activate the Virtual Environment)

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

### 2.4. Instalar Pacotes Necessários (Install Required Packages)

Com o `requirements.txt` no lugar e seu ambiente virtual ativado, instale as dependências executando:

```bash
pip install -r requirements.txt
```

Este comando fará o download e instalará todos os pacotes listados.

## 3\. Rodando o Aplicativo Principal (Interface Tkinter) (Running the Main Application (Tkinter Interface))

Após tudo o que foi feito, podemos rodar nossa aplicação para tanto, para executar o script principal, certifique-se de que seu ambiente virtual esteja ativado e todas as dependências instaladas. Em seguida, a partir do diretório raiz do repositório:

```bash
python main.py
```

Uma interface gráfica irá guiá-lo na seleção do projeto, escolha de arquivos e atualização da planilha GRD.