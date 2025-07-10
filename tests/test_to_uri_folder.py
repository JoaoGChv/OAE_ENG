import sys # Módulo para acessar parâmetros e funções específicas do sistema
import types # Módulo para criar tipos dinamicamente
from pathlib import Path # Módulo para manipular caminhos de arquivo de forma orientada a objetos
import pytest # Framework para testes Python

# Stub openpyxl so the module can be imported without the dependency installed
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

# Adiciona o diretório pai do diretório pai do arquivo atual ao sys.path.
# Isso permite que os módulos do projeto (como utils.planilha_gerador) sejam importados.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.planilha_gerador import _to_uri_folder # Importa a função a ser testada

@pytest.mark.parametrize('inp,expected', [ # Decorador do pytest para parametrizar o teste com múltiplos conjuntos de dados
    ('C:\\Folder\\Sub', 'file:///C:/Folder/Sub/?open'), # Caso de teste 1: Caminho padrão Windows
    ('C:\\Folder\\Sub\\', 'file:///C:/Folder/Sub/?open'), # Caso de teste 2: Caminho Windows com barra invertida no final
    ('G://Folder/Sub', 'file:///G:/Folder/Sub/?open'), # Caso de teste 3: Caminho com barras mistas
    ('E:\\Entrega- 2\\Docs', 'file:///E:/Entrega-2/Docs/?open'), # Caso de teste 4: Caminho com espaço e hífens
    ('C:\\My Folder\\Stuff', 'file:///C:/My%20Folder/Stuff/?open'), # Caso de teste 5: Caminho com espaços, esperando codificação URL (%20)
])
def test_to_uri_folder(inp: str, expected: str) -> None:
    # Função de teste que verifica a conversão de caminhos de pasta para URIs.
    # inp: Caminho de entrada (string)
    # expected: URI esperada como saída (string)
    assert _to_uri_folder(inp) == expected # Compara a saída da função com o valor esperado