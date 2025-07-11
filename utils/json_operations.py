import json
import os

def carregar_nomenclatura_json(numero_projeto, caminho_json):
    if not os.path.exists(caminho_json):
        return None
    try:
        with open(caminho_json, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data.get(numero_projeto, None)
    except Exception as e:
        print(f"Erro ao carregar nomenclatura JSON: {e}")
        return None


def salvar_dados(caminho, dados):
    try:
        with open(caminho, 'w', encoding='utf-8') as f:
            json.dump(dados, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Erro ao salvar dados: {e}") 
