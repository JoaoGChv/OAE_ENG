import sys
import types
from pathlib import Path
import pytest

# No tkinter stubbing because we don't create real Tk instance

from oae import ui

class DummyTree:
    def __init__(self):
        self.deleted = []
    def index(self, iid):
        return 0
    def delete(self, iid):
        self.deleted.append(iid)

class DummyMessageBox:
    def __init__(self):
        self.ask_called = False
        self.warn_called = False
    def askyesno(self, *a, **k):
        self.ask_called = True 
        return True
    def showwarning(self, *a, **k):
        self.warn_called = True


def test_excluir_send2trash_error(tmp_path, monkeypatch):
    f = tmp_path / "temp.txt"
    f.write_text("data")

    # prepare object without calling Tk.__init__
    obj = ui.TelaVisualizacaoEntregaAnterior.__new__(ui.TelaVisualizacaoEntregaAnterior)
    obj.checked = {"x": True}
    obj.tree = DummyTree()
    obj.lista_arquivos = [("", f.name, f.stat().st_size, str(f), "")]

    mbox = DummyMessageBox()
    monkeypatch.setattr(ui, "messagebox", mbox)

    def raise_oserror(path):
        raise OSError("fail")
    monkeypatch.setattr(ui, "send2trash", raise_oserror)

    obj._excluir_selecionados()

    assert not f.exists()
    assert obj.tree.deleted == ["x"]
    assert mbox.ask_called