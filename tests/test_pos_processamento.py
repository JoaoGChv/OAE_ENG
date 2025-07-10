# Se quiser voltar ao que era originalmente esse arquivo: test_pos_processamente precisa ser apagado ou inteiramente comentado.

import os
from pathlib import Path
import types
import sys
import pytest

# stub openpyxl similar to other tests
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

from oae.file_ops import (
    pos_processamento,
    GRD_MASTER_AP_NOME,
    GRD_MASTER_PE_NOME,
)
from utils import planilha_gerador


def test_pos_processamento_selects_correct_master(tmp_path, monkeypatch):
    called = {}
    def dummy_excel(*args, **kwargs):
        path = kwargs.get('caminho_excel') or args[0]
        called['path'] = str(path)
        Path(path).write_text('x')
    monkeypatch.setattr(planilha_gerador, 'criar_ou_atualizar_planilha', dummy_excel)
    import oae.file_ops as file_ops
    monkeypatch.setattr(file_ops, 'criar_ou_atualizar_planilha', dummy_excel)

    # AP delivery
    ap_file = tmp_path / 'a.txt'
    ap_file.write_text('x')
    data_ap = {}
    pos_processamento(
        True,
        str(tmp_path),
        data_ap,
        [('', ap_file.name, ap_file.stat().st_size, str(ap_file), '')],
        [],
        [],
        [],
        'AP',
    )
    assert Path(tmp_path, GRD_MASTER_AP_NOME).exists()
    assert called['path'].endswith(GRD_MASTER_AP_NOME)

    # PE delivery in new folder
    pe_dir = tmp_path / 'pe'
    pe_dir.mkdir()
    pe_file = pe_dir / 'b.txt'
    pe_file.write_text('y')
    called.clear()
    data_pe = {}
    pos_processamento(
        True,
        str(pe_dir),
        data_pe,
        [('', pe_file.name, pe_file.stat().st_size, str(pe_file), '')],
        [],
        [],
        [],
        'PE',
    )
    assert Path(pe_dir, GRD_MASTER_PE_NOME).exists()
    assert called['path'].endswith(GRD_MASTER_PE_NOME)