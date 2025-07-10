import sys # Módulo para interagir com o sistema Python (ex: sys.modules)
import types # Módulo para criar tipos e módulos dinamicamente
from pathlib import Path # Módulo para manipular caminhos de arquivo de forma orientada a objetos
import pytest # Framework de testes para Python

# No tkinter stubbing because we don't create real Tk instance
# Comentário: Não há simulação (stubbing) de Tkinter completa porque não é criada uma instância real do Tk.

from oae import ui # Importa o módulo 'ui' do pacote 'oae'

class DummyTree:
    # Classe dummy para simular um widget Treeview (componente de árvore/lista da GUI).
    # Usado para verificar se itens são marcados para exclusão.
    def __init__(self):
        self.deleted = [] # Lista para registrar os IDs dos itens que seriam "deletados"

    def index(self, iid):
        # Simula o método index do Treeview, sempre retornando 0 para simplificar.
        return 0

    def delete(self, iid):
        # Simula o método delete do Treeview, adicionando o ID do item à lista 'deleted'.
        self.deleted.append(iid)

class DummyMessageBox:
    # Classe dummy para simular as caixas de diálogo (messagebox) da GUI.
    # Usado para verificar se as funções de pergunta e aviso foram chamadas.
    def __init__(self):
        self.ask_called = False # Flag para indicar se askyesno foi chamado
        self.warn_called = False # Flag para indicar se showwarning foi chamado

    def askyesno(self, *a, **k):
        # Simula a caixa de diálogo de pergunta "sim/não", sempre retornando True (sim).
        self.ask_called = True # Marca que askyesno foi chamado
        return True # Retorna True para simular uma resposta "sim"

    def showwarning(self, *a, **k):
        # Simula a caixa de diálogo de aviso.
        self.warn_called = True # Marca que showwarning foi chamado

def test_excluir_send2trash_error(tmp_path, monkeypatch):
    # Função de teste para verificar o comportamento de _excluir_selecionados
    # quando a operação send2trash (enviar para a lixeira) falha com um OSError.
    # tmp_path: Fixture do pytest que fornece um diretório temporário para criar arquivos de teste.
    # monkeypatch: Fixture do pytest para modificar o comportamento de módulos/funções em tempo de execução.

    f = tmp_path / "temp.txt" # Cria um caminho para um arquivo temporário
    f.write_text("data") # Escreve algum conteúdo no arquivo temporário

    # prepare object without calling Tk.__init__
    # Prepara uma instância da classe TelaVisualizacaoEntregaAnterior sem chamar seu construtor Tkinter.
    # Isso é feito para evitar a necessidade de uma aplicação Tkinter real nos testes.
    obj = ui.TelaVisualizacaoEntregaAnterior.__new__(ui.TelaVisualizacaoEntregaAnterior)
    obj.checked = {"x": True} # Simula que um item 'x' está selecionado na GUI
    obj.tree = DummyTree() # Atribui a instância da classe dummy Treeview
    obj.lista_arquivos = [("", f.name, f.stat().st_size, str(f), "")] # Simula a lista de arquivos na GUI

    mbox = DummyMessageBox() # Cria uma instância da caixa de diálogo dummy
    monkeypatch.setattr(ui, "messagebox", mbox) # Substitui o messagebox real do módulo ui pelo dummy

    def raise_oserror(path):
        # Função dummy que simula o comportamento de send2trash falhando com um OSError.
        raise OSError("fail") # Lança um OSError simulado
    monkeypatch.setattr(ui, "send2trash", raise_oserror) # Substitui a função send2trash real pela dummy

    obj._excluir_selecionados() # Chama o método que está sendo testado

    assert not f.exists() # Verifica se o arquivo foi de fato excluído (apesar do erro simulado em send2trash, o teste assume que o arquivo é removido)
    assert obj.tree.deleted == ["x"] # Verifica se o item 'x' foi marcado como deletado no Treeview dummy
    assert mbox.ask_called # Verifica se a caixa de diálogo de pergunta (askyesno) foi chamada