---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Repositório da Aplicação TKinter para gerenciamento de arquivos da OAE - ENG.
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

*Está é uma interface Tkinter que faz o gerenciamento de arquivos para um escritório de projetos em Engenharia. O projeto tem especificações próprias mas é modular o suficiente para que possa ser aplicado em outros contextos. Ele permite:*

- Seleção de Projetos: Através de uma interface gráfica, o usuário pode selecionar um projeto a partir de uma lista carregada de um arquivo JSON.
- Seleção de Disciplinas: Dentro de um projeto, o usuário escolhe uma disciplina específica para a qual os arquivos serão entregues.
- Seleção de Arquivos: O usuário seleciona os arquivos relevantes para a entrega dentro da pasta de entregas da disciplina escolhida.
- Extração de Dados: O código extrai informações relevantes dos nomes dos arquivos, como status, fase, tipo e revisão.
- Validação de Nomenclatura: Uma interface permite ao usuário revisar e corrigir a nomenclatura dos arquivos com base em regras predefinidas.
- Verificação de Revisão: O sistema identifica e separa arquivos revisados e obsoletos com base na nomenclatura.
- Organização de Arquivos: Os arquivos são movidos para pastas "Revisados" e "Obsoletos" para melhor organização.
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Para consguir rodar a aplicação em qualquer máquina, faça o seguinte processo: 

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

## 6\. Rodando o Aplicativo Principal (Interface Tkinter) (Running the Main Application (Tkinter Interface))

Após tudo o que foi feito, podemos rodar nossa aplicação para tanto, para executar o script principal, certifique-se de que seu ambiente virtual esteja ativado e todas as dependências instaladas. Em seguida, a partir do diretório raiz do repositório:

```bash
python gerenciar_entregas.py
```

Uma interface gráfica irá guiá-lo na seleção do projeto, escolha de arquivos e atualização da planilha GRD.
