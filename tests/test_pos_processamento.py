# Se quiser voltar ao que era originalmente esse arquivo: test_pos_processamente precisa ser apagado ou inteiramente comentado.

import os # Módulo para interagir com o sistema operacional
from pathlib import Path # Módulo para manipular caminhos de arquivo de forma orientada a objetos
import types # Módulo para criar tipos dinamicamente
import sys # Módulo para acessar parâmetros e funções específicas do sistema
import pytest # Framework para testes Python

# stub openpyxl similar to other tests
# Este bloco de código "stub" (simula) o módulo openpyxl.
# É usado em testes para evitar a dependência real do openpyxl,
# permitindo que o código de teste seja executado mesmo sem o openpyxl instalado,
# ou para controlar o comportamento de suas funções durante o teste.
if 'openpyxl' not in sys.modules: # Verifica se openpyxl já foi importado
    openpyxl = types.ModuleType('openpyxl') # Cria um módulo dummy para openpyxl
    openpyxl.Workbook = object # Define Workbook como um objeto genérico
    openpyxl.load_workbook = lambda *a, **k: None # Simula a função load_workbook
    class _Dummy: # Classe dummy para estilos e utilitários
        def __init__(self, *a, **k):
            pass
    openpyxl.styles = types.ModuleType('openpyxl.styles') # Módulo dummy para estilos
    openpyxl.styles.Alignment = _Dummy # Classe dummy para Alinhamento
    openpyxl.styles.Font = _Dummy # Classe dummy para Fonte
    openpyxl.styles.PatternFill = _Dummy # Classe dummy para Preenchimento
    openpyxl.utils = types.ModuleType('openpyxl.utils') # Módulo dummy para utilitários
    openpyxl.utils.get_column_letter = lambda n: str(n) # Simula a função get_column_letter
    sys.modules['openpyxl'] = openpyxl # Adiciona o módulo dummy ao sys.modules
    sys.modules['openpyxl.styles'] = openpyxl.styles # Adiciona o módulo dummy de estilos
    sys.modules['openpyxl.utils'] = openpyxl.utils # Adiciona o módulo dummy de utilitários

# Adiciona o diretório pai do diretório pai do arquivo atual ao sys.path.
# Isso permite que os módulos do projeto (como oae.file_ops) sejam importados.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from oae.file_ops import (
    pos_processamento, # Importa a função principal a ser testada
    GRD_MASTER_AP_NOME, # Importa o nome da planilha mestra AP
    GRD_MASTER_PE_NOME, # Importa o nome da planilha mestra PE
)
from utils import planilha_gerador # Importa o módulo planilha_gerador

def test_pos_processamento_selects_correct_master(tmp_path, monkeypatch):
    # Função de teste que verifica se pos_processamento seleciona a planilha mestra correta (AP ou PE)
    # tmp_path: Fixture do pytest que fornece um diretório temporário para testes
    # monkeypatch: Fixture do pytest para modificar módulos, classes ou funções durante o teste

    called = {} # Dicionário para armazenar informações sobre chamadas de funções mockadas

    def dummy_excel(*args, **kwargs):
        # Função dummy que simula criar_ou_atualizar_planilha do planilha_gerador.
        # Ela apenas registra o caminho do arquivo Excel que seria criado/atualizado.
        path = kwargs.get('caminho_excel') or args[0] # Obtém o caminho do argumento
        called['path'] = str(path) # Armazena o caminho
        Path(path).write_text('x') # Cria um arquivo dummy no caminho especificado

    # Substitui a função original 'criar_ou_atualizar_planilha' pela dummy_excel no módulo planilha_gerador
    monkeypatch.setattr(planilha_gerador, 'criar_ou_atualizar_planilha', dummy_excel)
    import oae.file_ops as file_ops # Reimporta o módulo para garantir que as alterações do monkeypatch sejam aplicadas
    # Substitui a função 'criar_ou_atualizar_planilha' pela dummy_excel no módulo oae.file_ops
    monkeypatch.setattr(file_ops, 'criar_ou_atualizar_planilha', dummy_excel)

    # Teste para entrega AP (Apresentação?)
    ap_file = tmp_path / 'a.txt' # Cria um caminho de arquivo dummy para AP
    ap_file.write_text('x') # Cria um arquivo dummy
    data_ap = {} # Dados anteriores vazios para a primeira entrega AP
    pos_processamento(
        True, # primeira_entrega = True
        str(tmp_path), # Diretório da entrega
        data_ap, # Dados anteriores
        [('', ap_file.name, ap_file.stat().st_size, str(ap_file), '')], # Arquivos novos
        [], # Arquivos revisados
        [], # Arquivos alterados
        [], # Obsoletos
        'AP', # Tipo de entrega
    )
    # Verifica se a planilha mestra AP foi criada no diretório temporário
    assert Path(tmp_path, GRD_MASTER_AP_NOME).exists()
    # Verifica se o caminho da planilha que seria atualizada termina com o nome da planilha mestra AP
    assert called['path'].endswith(GRD_MASTER_AP_NOME)

    # Teste para entrega PE (Produção? Projeto?) em uma nova pasta
    pe_dir = tmp_path / 'pe' # Cria um subdiretório dummy para PE
    pe_dir.mkdir() # Cria o diretório
    pe_file = pe_dir / 'b.txt' # Cria um caminho de arquivo dummy para PE
    pe_file.write_text('y') # Cria um arquivo dummy
    called.clear() # Limpa o registro de chamadas para o próximo teste
    data_pe = {} # Dados anteriores vazios para a primeira entrega PE
    pos_processamento(
        True, # primeira_entrega = True
        str(pe_dir), # Diretório da entrega (o novo subdiretório)
        data_pe, # Dados anteriores
        [('', pe_file.name, pe_file.stat().st_size, str(pe_file), '')], # Arquivos novos
        [], # Arquivos revisados
        [], # Arquivos alterados
        [], # Obsoletos
        'PE', # Tipo de entrega
    )
    # Verifica se a planilha mestra PE foi criada no subdiretório temporário
    assert Path(pe_dir, GRD_MASTER_PE_NOME).exists()
    # Verifica se o caminho da planilha que seria atualizada termina com o nome da planilha mestra PE
    assert called['path'].endswith(GRD_MASTER_PE_NOME)