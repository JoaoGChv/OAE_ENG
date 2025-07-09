import sys
import types
import os
import json
from pathlib import Path
import io
import importlib
import tempfile

import pytest

pytest.importorskip("flask")

# Stub openpyxl like other tests so the modules import without the real package
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

@pytest.fixture
def client(tmp_path, monkeypatch):
    project_dir = tmp_path / "proj1"
    project_dir.mkdir()
    projects_json = tmp_path / "projects.json"
    projects_json.write_text(json.dumps({"001": str(project_dir)}))
    monkeypatch.setenv("OAE_PROJETOS_JSON", str(projects_json))

    from utils import planilha_gerador

    def dummy_excel(path, *a, **k):
        Path(path).write_text("dummy")

    monkeypatch.setattr(planilha_gerador, "criar_ou_atualizar_planilha", dummy_excel)

    import oae.file_ops
    import web_app.app
    importlib.reload(oae.file_ops)
    importlib.reload(web_app.app)

    return web_app.app.app.test_client(), project_dir

def test_select_project_page(client):
    c, _ = client
    resp = c.get("/select_project")
    assert resp.status_code == 200
    assert b"Select Project" in resp.data

def test_upload_creates_folder_and_spreadsheet(client):
    c, project_dir = client
    excel_path = Path(tempfile.gettempdir()) / "grd_web.xlsx"
    if excel_path.exists():
        excel_path.unlink()
    data = {"files": (io.BytesIO(b"dummy"), "sample.txt")}
    resp = c.post(f"/upload?project={project_dir}&tipo=AP", data=data, content_type="multipart/form-data")
    assert resp.status_code == 200
    delivery_dir = project_dir / "AP" / "1.AP - Entrega-1"
    assert delivery_dir.exists()
    assert (delivery_dir / "sample.txt").exists()
    assert excel_path.exists()