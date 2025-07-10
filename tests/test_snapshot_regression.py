import sys
import types
from pathlib import Path
import pytest

# Stub openpyxl so the module can be imported without the dependency installed
if 'openpyxl' not in sys.modules:
    openpyxl = types.ModuleType('openpyxl')
    openpyxl.Workbook = object
    openpyxl.load_workbook = lambda *a, **k: None
    class _Dummy:
        def __init__(self, *a, **k):
            pass

    openpyxl.styles = types.ModuleType('openpyxl.styles')
    openpyxl.styles.Alignment = _Dummy
    openpyxl.styles.Font = _Dummy
    openpyxl.styles.PatternFill = _Dummy
    openpyxl.utils = types.ModuleType('openpyxl.utils')
    openpyxl.utils.get_column_letter = lambda n: str(n)
    sys.modules['openpyxl'] = openpyxl
    sys.modules['openpyxl.styles'] = openpyxl.styles
    sys.modules['openpyxl.utils'] = openpyxl.utils

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.planilha_gerador import _key, _carregar_snapshot, _determinar_status


class DummyWS:
    def __init__(self, mapping, max_row):
        self.mapping = mapping
        self.max_row = max_row

    def cell(self, row, column):
        class Cell:
            pass
        c = Cell()
        c.value = self.mapping.get(row)
        return c


def test_key_normalizes_base():
    assert _key("Doc_A1-R01.dwg") == ("doc-a1", ".dwg")


def test_snapshot_recognizes_unchanged():
    estado = {"doc-a1|.dwg": {"revisao": "R01", "tamanho": 10, "timestamp": 1}}
    ws = DummyWS({2: "Doc_A1-R01.dwg"}, max_row=2)
    snap = _carregar_snapshot(ws, linha_titulo=1, col_prev=1, estado_anterior=estado)
    assert ("doc-a1", ".dwg") in snap
    info = {"rev": "R01", "tam": 10, "ts": 1}
    status = _determinar_status(info, snap[("doc-a1", ".dwg")])
    assert status == "inalterado"