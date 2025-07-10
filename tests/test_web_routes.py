import sys # Módulo para interagir com o sistema Python (ex: sys.modules)
import types # Módulo para criar tipos e módulos dinamicamente
import os # Módulo para interagir com o sistema operacional (ex: variáveis de ambiente)
import json # Módulo para trabalhar com dados JSON
from pathlib import Path # Módulo para manipular caminhos de arquivo de forma orientada a objetos
import io # Módulo para trabalhar com fluxos de I/O (usado para simular upload de arquivo)
import importlib # Módulo para recarregar módulos (importlib.reload)
import tempfile # Módulo para criar arquivos e diretórios temporários

import pytest # Framework de testes para Python

pytest.importorskip("flask") # Ignora os testes se o módulo 'flask' não estiver instalado

# Stub openpyxl like other tests so the modules import without the real package
# Este bloco de código "stub" (simula) o módulo openpyxl.
# É usado em testes para evitar a dependência real do openpyxl,
# permitindo que o código de teste seja executado mesmo sem o openpyxl instalado,
# ou para controlar o comportamento de suas funções durante o teste.
if 'openpyxl' not in sys.modules: # Verifica se openpyxl já foi importado
    openpyxl = types.ModuleType('openpyxl') # Cria um módulo dummy para openpyxl
    openpyxl.Workbook = object # Define Workbook como um objeto genérico
    openpyxl.load_workbook = lambda *a, **k: None # Simula a função load_workbook (não faz nada)
    class _Dummy: # Classe dummy para simular objetos de estilo e alinhamento
        def __init__(self, *a, **k):
            pass
    openpyxl.styles = types.ModuleType('openpyxl.styles') # Módulo dummy para estilos
    openpyxl.styles.Alignment = _Dummy # Classe dummy para Alinhamento
    openpyxl.styles.Font = _Dummy # Classe dummy para Fonte
    openpyxl.styles.PatternFill = _Dummy # Classe dummy para Preenchimento
    openpyxl.utils = types.ModuleType('openpyxl.utils') # Módulo dummy para utilitários
    openpyxl.utils.get_column_letter = lambda n: str(n) # Simula a função get_column_letter (retorna o número como string)
    sys.modules['openpyxl'] = openpyxl # Adiciona o módulo dummy ao sys.modules
    sys.modules['openpyxl.styles'] = openpyxl.styles # Adiciona o módulo dummy de estilos
    sys.modules['openpyxl.utils'] = openpyxl.utils # Adiciona o módulo dummy de utilitários

@pytest.fixture # Decorador do pytest que define uma função como uma fixture (recurso de setup/teardown para testes)
def client(tmp_path, monkeypatch):
    # Fixture que configura o ambiente para os testes da aplicação web.
    # tmp_path: Diretório temporário fornecido pelo pytest.
    # monkeypatch: Fixture para modificar o ambiente e o comportamento de funções.

    project_dir = tmp_path / "proj1" # Define um diretório de projeto temporário
    project_dir.mkdir() # Cria o diretório do projeto
    projects_json = tmp_path / "projects.json" # Define um arquivo JSON de projetos temporário
    # Escreve um JSON com um projeto dummy e seu diretório
    projects_json.write_text(json.dumps({"001": str(project_dir)}))
    # Define a variável de ambiente OAE_PROJETOS_JSON para apontar para o arquivo JSON dummy
    monkeypatch.setenv("OAE_PROJETOS_JSON", str(projects_json))

    from utils import planilha_gerador # Importa o módulo planilha_gerador

    def dummy_excel(path, *a, **k):
        # Função dummy que simula a criação/atualização de uma planilha Excel.
        # Apenas cria um arquivo vazio no caminho especificado.
        Path(path).write_text("dummy")

    # Substitui a função real 'criar_ou_atualizar_planilha' pela dummy_excel no módulo planilha_gerador
    monkeypatch.setattr(planilha_gerador, "criar_ou_atualizar_planilha", dummy_excel)

    import oae.file_ops # Importa o módulo de operações de arquivo
    import web_app.app # Importa o módulo principal da aplicação web
    importlib.reload(oae.file_ops) # Recarrega o módulo file_ops para aplicar as configurações do monkeypatch
    importlib.reload(web_app.app) # Recarrega o módulo da aplicação web para aplicar as configurações do monkeypatch

    # Retorna o cliente de teste da aplicação Flask e o diretório do projeto dummy
    return web_app.app.app.test_client(), project_dir

def test_select_project_page(client):
    # Testa se a página de seleção de projeto carrega corretamente.
    c, _ = client # Desempacota o cliente de teste da fixture
    resp = c.get("/select_project") # Faz uma requisição GET para a rota /select_project
    assert resp.status_code == 200 # Verifica se o status HTTP da resposta é 200 (OK)
    assert b"Select Project" in resp.data # Verifica se o conteúdo da página contém o texto "Select Project"

def test_upload_creates_folder_and_spreadsheet(client):
    # Testa se o upload de um arquivo cria a pasta de entrega e a planilha corretamente.
    c, project_dir = client # Desempacota o cliente de teste e o diretório do projeto da fixture
    excel_path = Path(tempfile.gettempdir()) / "grd_web.xlsx" # Define um caminho temporário para a planilha Excel
    if excel_path.exists():
        excel_path.unlink() # Remove o arquivo Excel temporário se ele já existir (para garantir um teste limpo)
    
    # Prepara os dados para a requisição POST, simulando um upload de arquivo.
    # io.BytesIO(b"dummy"): Simula o conteúdo binário do arquivo.
    # "sample.txt": O nome do arquivo a ser "uploadado".
    data = {"files": (io.BytesIO(b"dummy"), "sample.txt")}
    
    # Faz uma requisição POST para a rota /upload com os dados do arquivo e parâmetros de projeto/tipo.
    resp = c.post(f"/upload?project={project_dir}&tipo=AP", data=data, content_type="multipart/form-data")
    
    assert resp.status_code == 200 # Verifica se o status HTTP da resposta é 200 (OK)
    
    # Define o caminho esperado da pasta de entrega criada.
    delivery_dir = project_dir / "AP" / "1.AP - Entrega-1"
    assert delivery_dir.exists() # Verifica se a pasta de entrega foi criada
    assert (delivery_dir / "sample.txt").exists() # Verifica se o arquivo "sample.txt" foi copiado para a pasta de entrega
    assert excel_path.exists() # Verifica se a planilha Excel temporária foi criada