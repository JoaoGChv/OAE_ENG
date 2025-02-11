import re
import os

def identificar_nome_com_revisao(nome_arquivo):
    nome_sem_extensao, extensao = os.path.splitext(nome_arquivo)
    padrao = re.compile(r'^(.*?)-R(\d{2})$', re.IGNORECASE)
    match = padrao.match(nome_sem_extensao)
    if match:
        nome_base = match.group(1)
        revisao = 'R' + match.group(2)
        return nome_base, revisao, extensao.lower()
    return nome_sem_extensao, '', extensao.lower()

def verificar_tokens(tokens, nomenclatura):
    if not nomenclatura:
        return ['mismatch'] * len(tokens)
    campos_cfg = nomenclatura.get("campos", [])
    tokens_esperados = []
    for cinfo in campos_cfg:
        tokens_esperados.append(cinfo.get("tipo", "livre"))
    result = []
    for i, token in enumerate(tokens):
        if i < len(tokens_esperados) and tokens_esperados[i] != "livre" and tokens_esperados[i] != token:
            result.append("mismatch")
        else:
            result.append("ok")
    return result
