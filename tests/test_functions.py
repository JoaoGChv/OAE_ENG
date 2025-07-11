import pytest
from utils.file_operations import listar_arquivos_no_diretorio
from utils.json_operations import carregar_nomenclatura_json

def test_listar_arquivos_no_diretorio():
    arquivos = listar_arquivos_no_diretorio("test_dir")
    assert isinstance(arquivos, list)

def test_carregar_nomenclatura_json():
    nomenclatura = carregar_nomenclatura_json("123", "nomenclaturas.json")
    assert nomenclatura is not None 
 