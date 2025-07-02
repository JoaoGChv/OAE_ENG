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
from utils.planilha_gerador import _to_uri_folder

@pytest.mark.parametrize('inp,expected', [
    ('C:\\Folder\\Sub', 'file:///C:/Folder/Sub/?open'),
    ('C:\\Folder\\Sub\\', 'file:///C:/Folder/Sub/?open'),
    ('G://Folder/Sub', 'file:///G:/Folder/Sub/?open'),
    ('E:\\Entrega-  2\\Docs', 'file:///E:/Entrega-2/Docs/?open'),
    ('C:\\My Folder\\Stuff', 'file:///C:/My%20Folder/Stuff/?open'),
])
def test_to_uri_folder(inp: str, expected: str) -> None:
    assert _to_uri_folder(inp) == expected