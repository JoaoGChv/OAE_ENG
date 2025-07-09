from __future__ import annotations
import datetime
import os
import re
import string
from pathlib import Path
from typing import Dict, List, Sequence, Tuple
import urllib.parse as _up  # apenas para future-proof; não usamos quote agora.
from urllib.parse import quote as _q  # para percent-encoding de caracteres não-ASCII
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import logging

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Estilo e constantes
# ---------------------------------------------------------------------------
COR_REVISADO   = "00A8FF"
COR_NOVO       = "91EF93"
COR_MODIFICADO = "FFA500"
COR_INALTERADO = "FFFFFF"
COR_TITULO     = "5B9BD5"

STATUS_COR = {
    "revisado":   COR_REVISADO,
    "novo":       COR_NOVO,
    "modificado": COR_MODIFICADO,
    "inalterado": COR_INALTERADO,
}

FONT_TITULO  = Font(bold=True, size=14)
ALIGN_CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)

LARG_A_I   = 8.43
LARG_J     = 8.43
LARG_GRUPO = 28.71
LARG_EXT   = 23
LARG_ENT   = 68           # largura de TODAS as colunas de entrega

GRUPOS_EXT: Dict[str, Sequence[str]] = {
    "DWG/DXF": [".dwg", ".dxf"],
    "PDF":     [".pdf"],
    "XLS/XLSX": [".xls", ".xlsx"],
    "DOC/DOCX": [".doc", ".docx"],
}

REV_REGEX = re.compile(r"([-._])R(\d{1,3})$", re.IGNORECASE)

# ---------------------------------------------------------------------------
# Utilidades de coluna/planilha
# ---------------------------------------------------------------------------
def _col_idx(letra: str) -> int:
    return string.ascii_uppercase.index(letra.upper()) + 1


def _merge_ai(ws, row: int, value: str = "", font: Font | None = None):
    ws.merge_cells(start_row=row, start_column=_col_idx("A"),
                   end_row=row,   end_column=_col_idx("I"))
    c = ws.cell(row=row, column=_col_idx("A"), value=value)
    c.alignment = ALIGN_CENTRO
    if font:
        c.font = font


def _set_widths(ws):
    for col in range(_col_idx("A"), _col_idx("I") + 1):
        ws.column_dimensions[get_column_letter(col)].width = LARG_A_I
    ws.column_dimensions["J"].width = LARG_J
    ws.column_dimensions["K"].width = LARG_GRUPO
    ws.column_dimensions["L"].width = LARG_EXT

# ---------------------------------------------------------------------------
# Chave, revisão, grupo
# ---------------------------------------------------------------------------
def _key(nome: str) -> Tuple[str, str]:
    base, ext = os.path.splitext(nome)
    m = REV_REGEX.search(base)
    if m:
        base = base[: m.start()]
    return base.lower(), ext.lower()


def _extrair_rev(nome: str) -> str:
    m = REV_REGEX.search(os.path.splitext(nome)[0])
    return f"R{int(m.group(2)):02d}" if m else ""


def _classificar_extensao(ext: str) -> str:
    for grp, exts in GRUPOS_EXT.items():
        if ext.lower() in exts:
            return grp
    return "Outros"

# ---------------------------------------------------------------------------
# NOVO helper: monta URI file:/// sem percent-encoding
# ---------------------------------------------------------------------------
from urllib.parse import quote as _q

def _to_uri_folder(path: str) -> str:
    path = path.replace("\\", "/").rstrip("/")
    path = re.sub(r"^([A-Za-z]):/{2,}", r"\1:/", path)     # G:// → G:/
    path = re.sub(r"-\s+(\d+)", r"-\1", path)              # Entrega-  2 → Entrega-2
    safe = "/:-_."
    encoded = "".join(
        ch if (ch.isalnum() or ch in safe) else _up.quote(ch, safe="")
        for ch in path
    )
    # barra final + “#” evita que o Excel modifique o hyperlink
    return "file:///" + encoded + "/?open"

# ---------------------------------------------------------------------------
# Helper que hidrata links faltantes em colunas antigas
# ---------------------------------------------------------------------------
def _hidratar_hyperlinks(ws, linha_titulo: int, dir_base: str) -> None:
    """Preenche hyperlinks de colunas antigas apontando para as pastas corretas."""
    col_first = _col_idx("M")  # 13
    mapas: Dict[int, str] = {}

    for col in range(col_first, ws.max_column + 1):
        cab = ws.cell(row=linha_titulo, column=col).value or ""
        m = re.search(r"(1\.AP|2\.PE)\s*-\s*Entrega-\s*(\d+)", str(cab))
        if m:
            numero = m.group(2)
            subdir = "AP" if m.group(1).startswith("1") else "PE"
            prefixo = f"{m.group(1)} - Entrega-"
            pasta = os.path.join(dir_base, subdir, f"{prefixo}{numero}")

            if not os.path.exists(pasta):
                pai = os.path.dirname(pasta)
                padrao = f"{prefixo}{numero}-OBSOLETO"
                candidatos = [p for p in os.listdir(pai) if p.startswith(padrao)] if os.path.isdir(pai) else []
                if candidatos:
                    pasta = os.path.join(pai, candidatos[0])

            mapas[col] = pasta

    for row in range(linha_titulo + 1, ws.max_row + 1):
        for col, pasta in mapas.items():
            cel = ws.cell(row=row, column=col)
            if cel.value:                              # tem nome de arquivo
                link = _to_uri_folder(pasta)
                logger.debug("[hidratar] row %s col %s → %s", row, col, link)
                cel.hyperlink = link 

# ---------------------------------------------------------------------------
# Função principal – lógica original preservada
# ---------------------------------------------------------------------------
def criar_ou_atualizar_planilha(
    caminho_excel: str | Path,
    tipo_entrega: str,
    num_entrega: int,
    diretorio_base: str,
    arquivos: List[Tuple[str, str, int, str, str]],
    estado_anterior: Dict[str, Dict[str, object]] | None = None,
):
    caminho_excel = Path(caminho_excel)
    wb, ws, _ = _abrir_ou_criar_wb(caminho_excel, diretorio_base)

    ws.freeze_panes = "J10"
    _set_widths(ws)

    linha_titulo = 9
    col_ent = _col_idx("M")
    ws.insert_cols(col_ent)
    col_prev = col_ent + 1

    for c in range(col_ent, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = LARG_ENT

    _pintar_cabecalho_titulos(ws, linha_titulo)
    prefixos = {"AP": "1.AP - Entrega-", "PE": "2.PE - Entrega-"}
    cab = ws.cell(row=linha_titulo, column=col_ent,
                  value=f"{prefixos.get(tipo_entrega,'ENT')} {num_entrega}")
    cab.alignment = ALIGN_CENTRO
    cab.font = Font(bold=True)
    cab.fill = PatternFill(start_color=COR_TITULO,
                           end_color=COR_TITULO,
                           fill_type="solid")

    snapshot_ant = _carregar_snapshot(ws, linha_titulo, col_prev, estado_anterior)

    atual_info = {}
    for rev, nome, tam, full_path, _ in arquivos:
        base, ext = _key(nome)
        atual_info[(base, ext)] = {
            "nome": nome,
            "rev": rev or _extrair_rev(nome),
            "tam": tam,
            "ts":  os.path.getmtime(full_path),
            "dir": os.path.dirname(full_path),
            "grupo": _classificar_extensao(ext),
            "ext": ext,
        }

    for snap in snapshot_ant.values():
        ws.cell(row=snap["row"], column=col_ent).fill = PatternFill(
            start_color=COR_INALTERADO, end_color=COR_INALTERADO, fill_type="solid"
        )

    linha_cursor = ws.max_row + 1

    for key_, info in atual_info.items():
        snap = snapshot_ant.get(key_)
        if snap:
            row = snap["row"]
            status = _determinar_status(info, snap) 
        else:
            row = linha_cursor
            linha_cursor += 1
            ws.cell(row=row, column=_col_idx("K"), value=info["grupo"])
            ws.cell(row=row, column=_col_idx("L"), value=info["ext"])
            status = "novo"

        cor = STATUS_COR[status]
        cel = ws.cell(row=row, column=col_ent, value=info["nome"])
        cel.fill = PatternFill(start_color=cor, end_color=cor, fill_type="solid")

        # ---------- hyperlink (para a PASTA, barra final garantida) ----------
        folder = info["dir"]
        link = _to_uri_folder(folder)
        logger.debug("[nova-col] linha %s | %s", row, link)
        cel.hyperlink = link
    # ---- hidrata colunas antigas ------------------------------------------
    _hidratar_hyperlinks(ws, linha_titulo, diretorio_base)
    wb.save(caminho_excel)
    logger.info("Planilha atualizada: %s", caminho_excel)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _determinar_status(atual: dict, snap: dict | None) -> str:
    if snap is None:
        return "novo"
    if atual["rev"] != snap.get("rev", ""):
        return "revisado"
    if snap.get("tam") is not None and atual["tam"] != snap["tam"]:
        return "modificado"
    if snap.get("ts") is not None and atual["ts"] != snap["ts"]:
        return "modificado"
    return "inalterado"


def _carregar_snapshot(ws, linha_titulo: int, col_prev: int,
                       estado_anterior: Dict[str, Dict[str, object]] | None = None):
    snap: Dict[Tuple[str, str], Dict[str, object]] = {}
    for row in range(linha_titulo + 1, ws.max_row + 1):
        nome = ws.cell(row=row, column=col_prev).value
        if not nome:
            continue
        base, ext = _key(nome)
        rev_plan = _extrair_rev(nome)
        dados_json = estado_anterior.get(f"{base}|{ext}", {}) if estado_anterior else {}
        snap[(base, ext)] = {
            "rev": dados_json.get("revisao", rev_plan),
            "tam": dados_json.get("tamanho"),
            "ts":  dados_json.get("timestamp"),
            "row": row,
        }
    return snap

# ---------------------------------------------------------------------------
# Layout inicial / cabeçalhos
# ---------------------------------------------------------------------------
def _abrir_ou_criar_wb(caminho: Path, dir_base: str):
    if caminho.exists():
        wb = load_workbook(caminho)
        return wb, wb.active, False
    wb = Workbook()
    ws = wb.active
    ws.title = "GRD"
    _montar_cabecalho(ws, dir_base)
    return wb, ws, True


def _montar_cabecalho(ws, dir_base: str):
    hoje = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M")
    _merge_ai(ws, 1, "OLIVEIRA ARAÚJO ENGENHARIA", FONT_TITULO)
    _merge_ai(ws, 2, "Lista de arquivos de projetos entregues com controle de revisões")
    _merge_ai(ws, 3, f"Diretório: {dir_base}")
    _merge_ai(ws, 4, f"Data de emissão: {hoje}")

    legendas = [
        ("Arquivo Revisado",   COR_REVISADO),
        ("Arquivo Novo",       COR_NOVO),
        ("Arquivo Modificado s/ Atualizar Revisão", COR_MODIFICADO),
        ("Arquivo Inalterado", COR_INALTERADO),
    ]
    for i, (txt, cor) in enumerate(legendas, start=5):
        _merge_ai(ws, i, txt)
        c = ws.cell(row=i, column=_col_idx("A"))
        c.fill = PatternFill(start_color=cor, end_color=cor, fill_type="solid")
        c.font = Font(bold=True)


def _pintar_cabecalho_titulos(ws, linha: int):
    # pinta K (Grupo) e L (Extensão)
    for col in (_col_idx("K"), _col_idx("L")):
        cel = ws.cell(row=linha, column=col)
        cel.fill = PatternFill(start_color=COR_TITULO,
                               end_color=COR_TITULO,
                               fill_type="solid")
        cel.alignment = ALIGN_CENTRO
        cel.font = Font(bold=True)
