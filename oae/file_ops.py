"""File processing utilities for OAE."""
from __future__ import annotations
import datetime
import json
import os
import re
import shutil
from typing import Dict, List, Sequence, Tuple

from utils.planilha_gerador import criar_ou_atualizar_planilha

try:
    from openpyxl import Workbook
except Exception as exc:  # pragma: no cover - runtime dependency check
    raise ImportError("openpyxl is required") from exc

# Paths can be overridden by environment variables

def _resolve_json_path(env_var: str, default_path: str) -> str:
    return os.getenv(env_var, default_path)

PROJETOS_JSON: str = _resolve_json_path(
    "OAE_PROJETOS_JSON",
    r"G:\Drives compartilhados\OAE-JSONS\diretorios_projetos.json",
)

NOMENCLATURAS_JSON: str = _resolve_json_path(
    "OAE_NOMENCLATURAS_JSON",
    r"G:\Drives compartilhados\OAE-JSONS\nomenclaturas.json",
)

ARQ_ULTIMO_DIR: str = _resolve_json_path(
    "OAE_ULTIMO_DIR_JSON",
    r"G:\Drives compartilhados\OAE-JSONS\ultimo_diretorio_arqs.json",
)

GRUPOS_EXT: Dict[str, Sequence[str]] = {
    "DWG/DXF": [".dwg", ".dxf"],
    "DOC/DOCX": [".doc", ".docx"],
    "XLS/XLSX": [".xls", ".xlsx"],
    "ZIP/RAR": [".zip", ".rar", ".7z"],
    "RVT": [".rvt"],
    "IFC": [".ifc"],
    "NWC": [".nwc"],
    "NWD": [".nwd"],
}

MESES = (
    "janeiro","fevereiro","março","abril","maio","junho",
    "julho","agosto","setembro","outubro","novembro","dezembro"
)

GRD_MASTER_NOME = "GRD_ENTREGAS.xlsx"


def criar_pasta_entrega_ap_pe(
    pasta_entrega_disc: str,
    tipo: str,
    arquivos: list[Tuple[str, str, int, str, str]],
) -> None:
    prefixo = "1.AP - Entrega-" if tipo == "AP" else "2.PE - Entrega-"
    subdir = "AP" if tipo == "AP" else "PE"
    pasta_base = os.path.join(pasta_entrega_disc, subdir)
    os.makedirs(pasta_base, exist_ok=True)

    entregas_ativas = sorted(
        [d for d in os.listdir(pasta_base)
         if d.startswith(prefixo) and not d.endswith("-OBSOLETO")],
        key=lambda n: int(re.search(r"(\d+)$", n).group(1))
    )
    n_prox = (
        int(re.search(r"(\d+)$", entregas_ativas[-1]).group(1)) + 1
        if entregas_ativas else 1
    )

    if entregas_ativas:
        ant_path = os.path.join(pasta_base, entregas_ativas[-1])
        novo_ant = ant_path + "-OBSOLETO"
        seq = 1
        while os.path.exists(novo_ant):
            seq += 1
            novo_ant = f"{ant_path}-OBSOLETO{seq}"
        os.rename(ant_path, novo_ant)

    nova_pasta = os.path.join(pasta_base, f"{prefixo}{n_prox}")
    os.makedirs(nova_pasta, exist_ok=False)

    for (_, nome, _, caminho_full, _) in arquivos:
        try:
            shutil.copy2(caminho_full, os.path.join(nova_pasta, nome))
        except FileNotFoundError:
            continue


def _safe_json_load(fp) -> dict:
    try:
        return json.load(fp)
    except json.JSONDecodeError:
        return {}


def carregar_nomenclatura_json(numero_projeto: str) -> Dict | None:
    if not os.path.exists(NOMENCLATURAS_JSON):
        return None
    with open(NOMENCLATURAS_JSON, "r", encoding="utf-8") as f:
        data = _safe_json_load(f)
    return data.get(numero_projeto)


def salvar_ultimo_diretorio(ultimo_dir: str) -> None:
    try:
        with open(ARQ_ULTIMO_DIR, "w", encoding="utf-8") as f:
            json.dump({"ultimo_diretorio": ultimo_dir}, f, ensure_ascii=False, indent=4)
    except OSError:
        pass


def carregar_ultimo_diretorio() -> str | None:
    if os.path.exists(ARQ_ULTIMO_DIR):
        try:
            with open(ARQ_ULTIMO_DIR, "r", encoding="utf-8") as f:
                data = _safe_json_load(f)
                return data.get("ultimo_diretorio")
        except OSError:
            return None
    return None


def extrair_numero_arquivo(nome_base: str) -> str:
    if len(nome_base) <= 11:
        return ""
    substring: str = nome_base[11:]
    match = re.search(r"(\d{3})", substring)
    return match.group(1) if match else ""


REV_REGEX: re.Pattern[str] = re.compile(r"^(.*?)[-_]R(\d{1,3})$", re.IGNORECASE)


def identificar_nome_com_revisao(nome_arquivo: str) -> Tuple[str, str, str]:
    nome_sem_extensao, extensao = os.path.splitext(nome_arquivo)
    nome_normalizado = nome_sem_extensao.replace("_", "-")
    match = REV_REGEX.match(nome_normalizado)
    if match:
        nome_base = match.group(1)
        revisao = "R" + match.group(2).zfill(2)
        return nome_base, revisao, extensao.lower()
    return nome_sem_extensao, "", extensao.lower()


def _parse_rev(rev: str) -> int:
    if not rev:
        return -1
    digits = re.findall(r"\d+", rev)
    return int(digits[0]) if digits else -1


def comparar_revisoes(r1: str, r2: str) -> int:
    try:
        return _parse_rev(r1) - _parse_rev(r2)
    except ValueError:
        return 0


DEFAULT_SEPARATORS: set[str] = {"-", "."}


def _obter_separadores_do_json(nomenclatura: Dict | None) -> set[str]:
    seps: set[str] = set()
    if nomenclatura:
        for campo in nomenclatura.get("campos", []):
            sep = campo.get("separador")
            if sep and isinstance(sep, str):
                seps.add(sep)
    return seps or DEFAULT_SEPARATORS


def split_including_separators(nome_sem_ext: str, nomenclatura: Dict | None) -> List[str]:
    tokens: List[str] = []
    seps = _obter_separadores_do_json(nomenclatura)
    i = 0
    while i < len(nome_sem_ext):
        c = nome_sem_ext[i]
        if c in seps:
            tokens.append(c)
            i += 1
            continue
        j = i
        while j < len(nome_sem_ext) and nome_sem_ext[j] not in seps:
            j += 1
        tokens.append(nome_sem_ext[i:j])
        i = j
    return tokens


def verificar_tokens(tokens: Sequence[str], nomenclatura: Dict | None) -> List[str]:
    if not nomenclatura:
        return ["mismatch"] * len(tokens)

    campos_cfg = nomenclatura.get("campos", [])
    tokens_esperados: List[Tuple[str, object]] = []
    for idx, cinfo in enumerate(campos_cfg):
        tokens_esperados.append(("campo", cinfo))
        if idx < len(campos_cfg) - 1:
            sep = cinfo.get("separador", "-")
            tokens_esperados.append(("sep", sep))

    result_tags: List[str] = []
    idx_exp = idx_tok = 0
    while idx_tok < len(tokens) and idx_exp < len(tokens_esperados):
        token = tokens[idx_tok]
        tipo_esp, conteudo_esp = tokens_esperados[idx_exp]

        if tipo_esp == "sep":
            result_tags.append("ok" if token == conteudo_esp else "mismatch")
            idx_tok += 1
            idx_exp += 1
            continue

        tipo_campo = conteudo_esp.get("tipo", "Fixo")
        fixos = conteudo_esp.get("valores_fixos", [])
        if tipo_campo == "Fixo" and fixos:
            valores_permitidos = [f.get("value") if isinstance(f, dict) else str(f) for f in fixos]
            result_tags.append("ok" if token in valores_permitidos else "mismatch")
        else:
            result_tags.append("ok")
        idx_tok += 1
        idx_exp += 1

    while idx_tok < len(tokens):
        result_tags.append("mismatch")
        idx_tok += 1
    while idx_exp < len(tokens_esperados): 
        result_tags.append("missing")
        idx_exp += 1
    return result_tags


def identificar_obsoletos_custom(lista_arqs: Sequence[Tuple[str, str, int, str, str]]):
    grouping: Dict[Tuple[str, str], List[Tuple[str, str, int, str, str]]] = {}
    for rv, a, tam, cam, dmod in lista_arqs:
        base, revision, ext = identificar_nome_com_revisao(a)
        key = (base.lower(), ext.lower())
        grouping.setdefault(key, []).append((rv, a, tam, cam, dmod))

    obsoletos: List[Tuple[str, str, int, str, str]] = []
    for arr in grouping.values():
        arr.sort(key=lambda x: _parse_rev(x[0]), reverse=True)
        obsoletos.extend(arr[1:])
    return obsoletos


def carregar_dados_anteriores(diretorio: str) -> Dict:
    caminho = os.path.join(diretorio, "dados_execucao_anterior.json")
    if os.path.exists(caminho):
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                return _safe_json_load(f)
        except OSError:
            pass
    return {}


def salvar_dados(diretorio: str, dados: Dict) -> None:
    caminho = os.path.join(diretorio, "dados_execucao_anterior.json")
    try:
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=4)
    except OSError:
        pass


def obter_info_ultima_entrega(dados_anteriores: Dict) -> str:
    entregas_oficiais = dados_anteriores.get("entregas_oficiais", 0)
    ultima_execucao = dados_anteriores.get("ultima_execucao")
    if ultima_execucao:
        dt = datetime.datetime.strptime(ultima_execucao, "%Y-%m-%d %H:%M:%S")
        return f"Entrega {entregas_oficiais} de dia {dt.day} de {MESES[dt.month-1]} de {dt.year}"
    return f"Entrega {entregas_oficiais}"


def tentar_novamente_operacao(operacao, *args, **kwargs):
    while True:
        try:
            return operacao(*args, **kwargs)
        except PermissionError:
            raise


def gerar_nomes_entrega(num_entrega: int):
    data_atual = datetime.datetime.now().strftime("%d_%m_%Y")
    return (
        f"GRD-ENTREGA.{num_entrega:02d}-{data_atual}.xlsx",
        f"ENTREGA.{num_entrega:02d}-{data_atual}",
        data_atual,
    )


_SIGLAS_STATUS = {"E", "C", "P", "R"}
_SEP_PATTERN = r"[-_.]"


def renomear_para_arquivado(nome_arquivo: str) -> str:
    base, ext = os.path.splitext(nome_arquivo)
    m = re.match(rf"^([ECPR])({_SEP_PATTERN}.+)$", base, re.IGNORECASE)
    if not m:
        return nome_arquivo
    novo_base = "A" + m.group(2)
    return novo_base + ext


def mover_obsoletos_e_grd_anterior(obsoletos, diretorio: str, num_entrega_atual: int):
    n_anterior = num_entrega_atual - 1
    data_atual = datetime.datetime.now().strftime("%d_%m_%Y")
    pasta_pai = os.path.dirname(diretorio)
    pasta_obsoletos = os.path.join(pasta_pai, f"Entrega_{n_anterior:02d}-Obsoletos-{data_atual}")
    os.makedirs(pasta_obsoletos, exist_ok=True)

    nome_arquivo_anterior, _, _ = gerar_nomes_entrega(n_anterior)
    grd_anterior = os.path.join(diretorio, nome_arquivo_anterior)
    if os.path.exists(grd_anterior):
        tentar_novamente_operacao(shutil.move, grd_anterior, os.path.join(pasta_obsoletos, nome_arquivo_anterior))

    with open(os.path.join(pasta_obsoletos, "lista_obsoletos.txt"), "w", encoding="utf-8") as f:
        for rv, arq, *_ in obsoletos:
            f.write(arq + "\n")

    for rv, arq, _, cam, _ in obsoletos:
        try:
            novo_nome = renomear_para_arquivado(arq)
            destino = os.path.join(pasta_obsoletos, novo_nome)

            if os.path.exists(destino):
                base, ext = os.path.splitext(novo_nome)
                seq = 1
                while True:
                    cand = f"{base}_dup{seq}{ext}"
                    destino = os.path.join(pasta_obsoletos, cand)
                    if not os.path.exists(destino):
                        break
                    seq += 1

            tentar_novamente_operacao(shutil.move, cam, destino)
        except FileNotFoundError:
            continue


def criar_arquivo_excel(diretorio: str, num_entrega: int, arquivos):
    nome_arquivo, nome_planilha, _ = gerar_nomes_entrega(num_entrega)
    caminho_excel = os.path.join(diretorio, nome_arquivo)
    wb = Workbook()
    ws = wb.active
    ws.title = nome_planilha
    ws.append(["Nome do arquivo", "Revisão", "Caminho completo"])
    for rv, arq, _, cam, _ in arquivos:
        ws.append([arq, rv or "", cam])
    wb.save(caminho_excel)
    return caminho_excel


def listar_arquivos_no_diretorio(diretorio):
    ignorar = {"dados_execucao_anterior.json", GRD_MASTER_NOME}
    for f in os.listdir(diretorio):
        if f.startswith("GRD-ENTREGA."):
            ignorar.add(f)
    saida = []
    for raiz, _dirs, files in os.walk(diretorio):
        for a in files:
            if a in ignorar:
                continue
            nb, rv, ex = identificar_nome_com_revisao(a)
            if ex in ['.jpg','.jpeg','.dwl','.dwl2','.png','.ini']:
                continue
            cam = os.path.join(raiz, a)
            tam = os.path.getsize(cam)
            dmod_ts = os.path.getmtime(cam)
            dmod = datetime.datetime.fromtimestamp(dmod_ts).strftime("%d/%m/%Y %H:%M")
            saida.append((rv, a, tam, cam, dmod))
    return saida


def analisar_comparando_estado(lista_de_arquivos, dados_anteriores):
    grouping = {}
    for rv, a, tam, cam, dmod in lista_de_arquivos:
        nb, revision, ex = identificar_nome_com_revisao(a)
        key = (nb.lower(), ex.lower())
        grouping.setdefault(key, []).append((rv, a, tam, cam, dmod))
    novos = []
    revisados = []
    alterados = []
    for key, items in grouping.items():
        items.sort(key=lambda x: comparar_revisoes(x[0], 'R99'))
        ant = dados_anteriores.get(f"{key[0]}|{key[1]}", None)
        rev_ant = ant["revisao"] if ant else ""
        tam_ant = ant["tamanho"] if ant else None
        ts_ant = ant.get("timestamp") if ant else None
        if not ant:
            if items:
                novos.append(items[0])
            for it in items[1:]:
                revisados.append(it)
        else:
            maior_rev = items[-1][0]
            comp = comparar_revisoes(maior_rev, rev_ant)
            num_rev_ant = comparar_revisoes(rev_ant, '')
            if comp > 0:
                for (rvx, arqx, tamx, camx, dmodx) in items:
                    nr = comparar_revisoes(rvx, '')
                    if nr > num_rev_ant:
                        revisados.append((rvx, arqx, tamx, camx, dmodx))
                    elif nr == num_rev_ant:
                        ts_now = os.path.getmtime(camx)
                        if tamx != tam_ant or (ts_ant is not None and ts_now != ts_ant):
                            alterados.append((rvx, arqx, tamx, camx, dmodx))
            elif comp == 0:
                for (rvx, arqx, tamx, camx, dmodx) in items:
                    if rvx == rev_ant:
                        ts_now = os.path.getmtime(camx)
                        if tamx != tam_ant or (ts_ant is not None and ts_now != ts_ant):
                            alterados.append((rvx, arqx, tamx, camx, dmodx))
            else:
                for (rvx, arqx, tamx, camx, dmodx) in items:
                    if rvx == rev_ant:
                        ts_now = os.path.getmtime(camx)
                        if tamx != tam_ant or (ts_ant is not None and ts_now != ts_ant):
                            alterados.append((rvx, arqx, tamx, camx, dmodx))
    return (novos, revisados, alterados)


def pos_processamento(
    primeira_entrega,
    diretorio,
    dados_anteriores,
    arquivos_novos,
    arquivos_revisados,
    arquivos_alterados,
    obsoletos,
    tipo_entrega: str | None = None,
):
    """Atualiza registros e planilhas após a análise de entrega."""
    num_entrega_atual = dados_anteriores.get("entregas_oficiais", 0) + 1
    caminho_excel_master = os.path.join(diretorio, GRD_MASTER_NOME)
    if not primeira_entrega:
        if obsoletos or dados_anteriores.get("entregas_oficiais", 0) >= 1:
            mover_obsoletos_e_grd_anterior(obsoletos, diretorio, num_entrega_atual)
    if primeira_entrega:
        union_ = []
        union_.extend(arquivos_novos)
        union_.extend(arquivos_revisados)
        union_.extend(arquivos_alterados)
        if not union_:
            return
    if primeira_entrega:
        lista_para_planilha = (
            arquivos_novos + arquivos_revisados + arquivos_alterados
        )
        if not lista_para_planilha:
            return
    else:
        lista_para_planilha = listar_arquivos_no_diretorio(diretorio)
        if not lista_para_planilha:
            return

    criar_ou_atualizar_planilha(
        caminho_excel=caminho_excel_master,
        tipo_entrega=tipo_entrega or "AP",
        num_entrega=num_entrega_atual,
        diretorio_base=diretorio,
        arquivos=lista_para_planilha,
        estado_anterior=dados_anteriores,
    )
    dados_anteriores["ultima_execucao"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    dados_anteriores["entregas_oficiais"] = num_entrega_atual
    grouping_final = {}
    all_files_now = listar_arquivos_no_diretorio(diretorio)
    for rv, a, tam, cam, dmod in all_files_now:
        nb, rev, ex = identificar_nome_com_revisao(a)
        key = (nb.lower(), ex.lower())
        grouping_final.setdefault(key, []).append((rv, a, tam, cam, dmod))
    for key, arr in grouping_final.items():
        arr.sort(key=lambda x: comparar_revisoes(x[0], 'R99'))
        revf = arr[-1][0]
        tamf = arr[-1][2]
        camf = arr[-1][3]
        tsf = os.path.getmtime(camf)
        dados_anteriores[f"{key[0]}|{key[1]}"] = {
            "revisao": revf if revf else '',
            "tamanho": tamf,
            "timestamp": tsf,
        }
    salvar_dados(diretorio, dados_anteriores)
