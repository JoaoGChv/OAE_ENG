# utils/planilha_gerador.py
# -*- coding: utf-8 -*-
"""
Gera e atualiza o workbook-mestre (“E principal”).

Regras principais
-----------------
• A coluna **M** (13) é sempre a entrega mais recente; entregas anteriores são
  empurradas para a direita.
• Todas as colunas de entrega têm largura **68 pt**.
• Cabeçalho (linhas 1-9) e legendas permanecem inalterados.
• Cada célula de arquivo contém um **hyperlink** que abre a pasta onde o
  arquivo está salvo (Explorer / Finder).

Alterações de 2025-06-17
------------------------
1.  Todas as colunas de entrega (nova e antigas) recebem width = 68 pt.
2.  O cabeçalho da nova coluna recebe o preenchimento azul `#5B9BD5`.
3.  Hiperlinks adicionados a **todas** as células de arquivo
    (`file:///…/pasta/`) – URL-escaped e com “/” final.
4.  Nenhuma outra lógica foi tocada.
"""

from __future__ import annotations

import datetime
import os
import re
import string
from pathlib import Path
from typing import Dict, List, Sequence, Tuple
from urllib.parse import quote

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

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

# Agrupamento por extensão – ajuste se quiser
GRUPOS_EXT: Dict[str, Sequence[str]] = {
    "DWG/DXF": [".dwg", ".dxf"],
    "PDF": [".pdf"],
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
        base = base[: m.start()]            # descarta sufixo -Rxx
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
# Função principal
# ---------------------------------------------------------------------------
def criar_ou_atualizar_planilha(
    caminho_excel: str | Path,
    tipo_entrega: str,        # "AP" ou "PE"
    num_entrega: int,         # 1, 2, 3…
    diretorio_base: str,
    arquivos: List[Tuple[str, str, int, str, str]],  # (rev, nome, tam, path, ts_str)
    estado_anterior: Dict[str, Dict[str, object]] | None = None,
):
    caminho_excel = Path(caminho_excel)
    wb, ws, _ = _abrir_ou_criar_wb(caminho_excel, diretorio_base)

    ws.freeze_panes = "J10"
    _set_widths(ws)

    # ---------------- inserção coluna nova ----------------
    linha_titulo = 9
    col_ent = _col_idx("M")           # 13
    ws.insert_cols(col_ent)           # empurra anteriores
    col_prev = col_ent + 1

    # aplica largura 68 pt para TODAS as colunas de entrega atuais+antigas
    for c in range(col_ent, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = LARG_ENT

    # ---------------- cabeçalho da nova entrega ------------
    _pintar_cabecalho_titulos(ws, linha_titulo)
    prefixos = {"AP": "1.AP - Entrega-", "PE": "2.PE - Entrega-"}
    cab = ws.cell(row=linha_titulo, column=col_ent,
                  value=f"{prefixos.get(tipo_entrega,'ENT')} {num_entrega}")
    cab.alignment = ALIGN_CENTRO
    cab.font = Font(bold=True)
    cab.fill = PatternFill(start_color=COR_TITULO,
                           end_color=COR_TITULO,
                           fill_type="solid")

    # ---------------- prepara snapshots -------------------
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

    # pré-pinta “inalterado” nas linhas já existentes
    for snap in snapshot_ant.values():
        ws.cell(row=snap["row"], column=col_ent).fill = PatternFill(
            start_color=COR_INALTERADO, end_color=COR_INALTERADO, fill_type="solid"
        )

    linha_cursor = ws.max_row + 1

    # ---------------- insere arquivos ---------------------
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

        # ---------------- hyperlink ------------------------
        folder = info["dir"].replace("\\", "/")
        if not folder.endswith("/"):
            folder += "/"
        cel.hyperlink = "file:///" + quote(folder)

        # actualiza snapshot in-memory
        snapshot_ant[key_] = {
            "rev": info["rev"],
            "tam": info["tam"],
            "ts":  info["ts"],
            "row": row,
        }

    wb.save(caminho_excel)
    print("Planilha atualizada:", caminho_excel)

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
