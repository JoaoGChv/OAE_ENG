from utils.validation import identificar_nome_com_revisao
import os
import datetime

def listar_arquivos_no_diretorio(diretorio):
    ignorar = {"dados_execucao_anterior.json"}
    for f in os.listdir(diretorio):
        if f.startswith("GRD-ENTREGA."):
            ignorar.add(f)
    saida = []
    for raiz, dirs, files in os.walk(diretorio):
        for a in files:
            if a in ignorar:
                continue
            nb, rv, ex = identificar_nome_com_revisao(a)
            if ex in ['.jpg', '.jpeg', '.dwl', '.dwl2', '.png', '.ini']:
                continue
            cam = os.path.join(raiz, a)
            tam = os.path.getsize(cam)
            dmod_ts = os.path.getmtime(cam)
            dmod = datetime.datetime.fromtimestamp(dmod_ts).strftime("%d/%m/%Y %H:%M")
            saida.append((rv, a, tam, cam, dmod))
            
    return saida

def carregar_ultimo_diretorio():
    try:
        with open("ultimo_diretorio.json", 'r', encoding='utf-8') as f:
            data = json.load(f)
            return data.get("ultimo_diretorio", None)
    except Exception as e:
        print(f"Erro ao carregar o último diretório: {e}")
        return None

def salvar_ultimo_diretorio(ultimo_dir):
    try:
        with open("ultimo_diretorio.json", 'w', encoding='utf-8') as f:
            json.dump({"ultimo_diretorio": ultimo_dir}, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Erro ao salvar o último diretório: {e}")
