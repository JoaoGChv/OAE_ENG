from flask import (
    Flask, # Classe principal para criar a aplicação Flask
    render_template, # Função para renderizar templates HTML
    redirect, # Função para redirecionar o usuário para outra URL
    url_for, # Função para construir URLs para funções de view
    request, # Objeto que contém os dados da requisição HTTP atual
    flash, # Função para exibir mensagens flash (mensagens temporárias para o usuário)
)
from utils.planilha_gerador import criar_ou_atualizar_planilha # Importa a função para criar/atualizar planilhas Excel
from oae.file_ops import (
    PROJETOS_JSON, # Caminho para o arquivo JSON de configuração de projetos
    criar_pasta_entrega_ap_pe, # Função para criar a estrutura de pastas de entrega
    listar_arquivos_no_diretorio, # Função para listar arquivos em um diretório
    identificar_obsoletos_custom, # Função para identificar arquivos obsoletos
    carregar_nomenclatura_json, # Função para carregar regras de nomenclatura
    split_including_separators, # Função para dividir strings mantendo os separadores
    verificar_tokens, # Função para verificar tokens de nomenclatura
    folder_mais_recente # Função para encontrar a pasta de entrega mais recente
)
from pathlib import Path # Módulo para manipular caminhos de arquivo de forma orientada a objetos
import os # Módulo para interagir com o sistema operacional (caminhos de arquivo, variáveis de ambiente)
import json # Módulo para trabalhar com dados JSON
import tempfile # Módulo para criar arquivos e diretórios temporários
import re # Módulo para expressões regulares

app = Flask(__name__) # Inicializa a aplicação Flask
app.secret_key = "oae-secret-key" # Define uma chave secreta para segurança (usada para sessões e mensagens flash)

@app.route('/') # Decorador que define a rota para a URL raiz
def index():
    # Rota para a página inicial da aplicação.
    return render_template('index.html') # Renderiza o template 'index.html'

def _load_projects():
    # Função interna para carregar a lista de projetos do arquivo PROJETOS_JSON.
    if not os.path.exists(PROJETOS_JSON): # Verifica se o arquivo JSON de projetos existe
        return [] # Retorna lista vazia se não existir
    try:
        with open(PROJETOS_JSON, "r", encoding="utf-8") as f: # Abre o arquivo JSON
            return list(json.load(f).items()) # Carrega o JSON e retorna uma lista de tuplas (código, caminho)
    except Exception: # Captura qualquer erro ao carregar o JSON
        return [] # Retorna lista vazia em caso de erro

@app.route('/select_project', methods=['GET', 'POST']) # Define a rota para a seleção de projetos, permitindo GET e POST
def select_project():
    # Rota para a página de seleção de projeto.
    projects = _load_projects() # Carrega a lista de projetos disponíveis
    if request.method == 'POST': # Se a requisição for um POST (formulário enviado)
        project = request.form.get('project') # Obtém o projeto selecionado do formulário
        tipo = request.form.get('tipo', 'AP') # Obtém o tipo de entrega (AP ou PE), padrão 'AP'
        if not project: # Se nenhum projeto foi selecionado
            flash('Selecione um projeto') # Exibe uma mensagem de erro
        else:
            # Redireciona para a página de upload, passando o projeto e o tipo como parâmetros na URL
            return redirect(url_for('upload_files', project=project, tipo=tipo))
    # Se a requisição for GET ou o POST falhou, renderiza a página de seleção de projeto
    return render_template('select_project.html', projects=projects)

def _next_delivery_dir(base: str, tipo: str) -> str:
    # Determina o caminho do próximo diretório de entrega sequencial (ex: "1.AP - Entrega-1").
    prefixo = '1.AP - Entrega-' if tipo == 'AP' else '2.PE - Entrega-' # Define o prefixo da pasta de entrega
    subdir = 'AP' if tipo == 'AP' else 'PE' # Define o subdiretório (AP ou PE)
    pasta_base = os.path.join(base, subdir) # Monta o caminho da pasta base (ex: /path/to/project/AP)
    os.makedirs(pasta_base, exist_ok=True) # Cria a pasta base se ela não existir

    # Lista os diretórios de entrega existentes que correspondem ao prefixo e não são obsoletos
    entregas = [
        d for d in os.listdir(pasta_base)
        if d.startswith(prefixo) and not d.endswith('-OBSOLETO')
    ]
    if entregas: # Se houver entregas existentes
        # Extrai os números das entregas e encontra o maior
        nums = [int(re.search(r"(\d+)$", d).group(1)) for d in entregas]
        n_prox = max(nums) + 1 # O próximo número será o maior + 1
    else:
        n_prox = 1 # Se não houver entregas, a próxima é a número 1
    # Retorna o caminho completo para o próximo diretório de entrega
    return os.path.join(pasta_base, f"{prefixo}{n_prox}")

@app.route('/upload', methods=['GET', 'POST']) # Define a rota para o upload de arquivos, permitindo GET e POST
def upload_files():
    # Rota para a página de upload de arquivos.
    project = request.args.get('project') # Obtém o caminho do projeto da URL
    tipo = request.args.get('tipo', 'AP') # Obtém o tipo de entrega da URL, padrão 'AP'

    if request.method == 'POST' and project: # Se a requisição for POST e um projeto estiver definido
        files = request.files.getlist('files') # Obtém a lista de arquivos enviados do formulário
        temp_saved = [] # Lista para armazenar informações dos arquivos temporariamente salvos

        with tempfile.TemporaryDirectory() as tmp_dir: # Cria um diretório temporário para salvar os arquivos recebidos
            for f in files: # Itera sobre cada arquivo enviado
                if not f.filename: # Se o arquivo não tiver um nome (vazio), pula
                    continue
                dest = os.path.join(tmp_dir, f.filename) # Define o caminho de destino temporário
                f.save(dest) # Salva o arquivo no diretório temporário
                size = os.path.getsize(dest) # Obtém o tamanho do arquivo
                temp_saved.append(('', f.filename, size, dest, '')) # Adiciona informações do arquivo à lista

            if temp_saved: # Se houver arquivos para processar
                try:
                    # Cria a estrutura de pastas de entrega e move os arquivos
                    criar_pasta_entrega_ap_pe(project, tipo, temp_saved)
                    # Determina o caminho da pasta de entrega recém-criada
                    nova_pasta = _next_delivery_dir(project, tipo) # Nota: esta chamada aqui é para obter o nome da pasta de *destino*,
                                                                  # mas 'criar_pasta_entrega_ap_pe' já cria a pasta.
                                                                  # É importante que 'nova_pasta' reflita o diretório exato onde os arquivos foram movidos.
                    
                    arquivos = listar_arquivos_no_diretorio(nova_pasta) # Lista os arquivos na nova pasta de entrega
                    excel_path = Path(tempfile.gettempdir()) / 'grd_web.xlsx' # Define o caminho para a planilha GRD temporária
                    
                    # Cria ou atualiza a planilha Excel GRD
                    criar_ou_atualizar_planilha(
                        excel_path, # Caminho da planilha
                        tipo, # Tipo de entrega
                        1, # Número da entrega (fixo como 1 neste contexto da GRD web)
                        nova_pasta, # Diretório base para a entrega atual
                        arquivos, # Arquivos a serem incluídos na planilha
                    )
                    # Redireciona para a página de resultado, mostrando o caminho da planilha
                    return render_template('result.html', path=excel_path)
                except Exception as exc: # Captura exceções durante o processo
                    flash(str(exc)) # Exibe a mensagem de erro para o usuário

    # Se a requisição for GET ou o POST falhou, renderiza a página de upload
    return render_template('upload_files.html', project=project, tipo=tipo)

@app.route('/history') # Define a rota para o histórico de entregas
def delivery_history():
    """Mostra os arquivos da entrega mais recente."""
    folder = request.args.get('folder') # Obtém o diretório do projeto da URL
    tipo = request.args.get('tipo', 'AP') # Obtém o tipo de entrega, padrão 'AP'
    if not folder:
        return 'Missing folder', 400 # Retorna erro se o diretório não for fornecido
    
    pasta = folder_mais_recente(folder, tipo) # Encontra a pasta da entrega mais recente para o projeto/tipo
    if not pasta: # Se nenhuma pasta de entrega recente for encontrada
        arquivos = [] # Lista de arquivos vazia
    else:
        todos = listar_arquivos_no_diretorio(pasta) # Lista todos os arquivos na pasta da entrega
        # Filtra arquivos, removendo os obsoletos identificados
        arquivos = [a for a in todos if a not in set(identificar_obsoletos_custom(todos))]
        arquivos.sort(key=lambda x: x[1].lower()) # Ordena os arquivos pelo nome (ignorando maiúsculas/minúsculas)
    # Renderiza a página de histórico, passando os arquivos e o tipo
    return render_template('history.html', files=arquivos, tipo=tipo)

@app.route('/nomenclature') # Define a rota para a verificação de nomenclatura
def nomenclature_check():
    """Página simples de validação de nomenclatura."""
    folder = request.args.get('folder') # Obtém o diretório do projeto da URL
    num = request.args.get('num') # Obtém o número da regra de nomenclatura da URL
    if not folder or not num:
        return 'Missing parameters', 400 # Retorna erro se parâmetros ausentes
    
    nomen = carregar_nomenclatura_json(num) # Carrega as regras de nomenclatura com base no número fornecido
    arquivos = listar_arquivos_no_diretorio(folder) # Lista os arquivos no diretório
    results = [] # Lista para armazenar os resultados da verificação de nomenclatura
    for _rv, nome, _tam, _cam, _dt in arquivos: # Itera sobre cada arquivo
        base, _ = os.path.splitext(nome) # Separa o nome base da extensão
        tokens = split_including_separators(base, nomen) # Divide o nome base em tokens usando as regras de nomenclatura
        tags = verificar_tokens(tokens, nomen) # Verifica cada token em relação às regras
        results.append((nome, list(zip(tokens, tags)))) # Adiciona o nome do arquivo e os resultados da verificação
    # Renderiza a página de nomenclatura, passando os resultados
    return render_template('nomenclature.html', results=results)

@app.route('/start_delivery') # Define uma rota de exemplo para iniciar uma entrega
def start_delivery():
    """Rota de exemplo que gera uma planilha GRD dummy."""
    temp_dir = tempfile.gettempdir() # Obtém o diretório temporário do sistema
    excel_path = Path(temp_dir) / 'grd_example.xlsx' # Define o caminho para uma planilha de exemplo temporária
    #