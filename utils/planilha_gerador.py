# utils/planilha_gerador.py
# -*- coding: utf-8 -*-
"""
Gera e atualiza um único workbook (“E principal”) conforme regras acordadas.

Layout geral (linhas):
 1  OLIVEIRA ARAÚJO ENGENHARIA
 2  Descrição
 3  Diretório
 4  Data de emissão
 5  🔵 Revisado
 6  🟢 Novo
 7  🟠 Modificado s/ Revisão
 8  ⚪ Inalterado
 9  Títulos: Grupo | Extensão | Entrega-1 | Entrega-2 …

Configuração:
  • Coluna J vazia (largura 8,43); dados começam em K.
  • Freeze panes em J10 (fixa linhas 1-9 e colunas A-I).
  • Larguras: K=28,71 ; L=23 ; entregas=47.
  • Cores:
        Revisado   #00A8FF
        Novo       #91EF93
        Modificado #FFA500
        Inalterado #FFFFFF
"""

from __future__ import annotations

import datetime
import os
import re
import string
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

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
LARG_ENT   = 47

# Agrupamento por extensão – ajuste se quiser
GRUPOS_EXT: Dict[str, Sequence[str]] = {
    "DWG/DXF": [".dwg", ".dxf"],
    "PDF": [".pdf"],
    "XLS/XLSX": [".xls", ".xlsx"],
    "DOC/DOCX": [".doc", ".docx"],
}

REV_REGEX = re.compile(r"[._-]R(\d{1,3})$", re.IGNORECASE)  # p/ extrair Rxx

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
    return base.lower(), ext.lower()


def _extrair_rev(nome: str) -> str:
    m = REV_REGEX.search(os.path.splitext(nome)[0])
    return f"R{int(m.group(1)):02d}" if m else ""


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
    wb, ws, primeira_vez = _abrir_ou_criar_wb(caminho_excel, diretorio_base)

    # ajuste constante
    ws.freeze_panes = "J10"
    _set_widths(ws)

    linha_titulo = 9
    _pintar_cabecalho_titulos(ws, linha_titulo, num_entrega)

    # coluna desta entrega
    col_ent = _col_idx("M") + (num_entrega - 1)
    ws.column_dimensions[get_column_letter(col_ent)].width = LARG_ENT
    prefixos = {"AP": "1.AP - Entrega-", "PE": "2.PE - Entrega-"}
    cell = ws.cell(row=linha_titulo, column=col_ent,
               value=f"{prefixos.get(tipo_entrega,'ENT')} {num_entrega}")
    cell.alignment = ALIGN_CENTRO
    cell.font = Font(bold=True)

    # snapshot anterior
    snapshot_ant = _carregar_snapshot(ws, linha_titulo, col_ent - 1, estado_anterior)

    # pré-processa entrega atual
    atual_info = {}
    for rev, nome, tam, _p, ts_str in arquivos:
        base, ext = _key(nome)
        ts_val = os.path.getmtime(_p)
        atual_info[(base, ext)] = {
            "nome": nome,
            "rev": rev or _extrair_rev(nome),
            "tam": tam,
            "ts":  ts_val,
            "grupo": _classificar_extensao(ext),
            "ext": ext,
        }

    # linhas já existentes -> marcar branco por default (inalterado)
    for snap in snapshot_ant.values():
        cell = ws.cell(row=snap["row"], column=col_ent)
        cell.value = ""        # limpar eventual lixo
        cell.fill = PatternFill(start_color=COR_INALTERADO,
                                end_color=COR_INALTERADO,
                                fill_type="solid")

    # cursor para novas linhas (após último registro)
    linha_cursor = ws.max_row + 1

    # processa cada arquivo atual
    for key, info in atual_info.items():
        snap = snapshot_ant.get(key)
        if snap:
            row = snap["row"]
            status = _determinar_status(info, snap)
        else:
            # linha nova
            row = linha_cursor
            ws.cell(row=row, column=_col_idx("K"), value=info["grupo"])
            ws.cell(row=row, column=_col_idx("L"), value=info["ext"])
            linha_cursor += 1
            status = "novo"

        cor = STATUS_COR[status]
        cell = ws.cell(row=row, column=col_ent, value=info["nome"])
        cell.fill = PatternFill(start_color=cor, end_color=cor,
                                fill_type="solid")

        # guardar comentário invisível com tam para futura comparação
        cell.comment = None
        cell.comment = None   # no-op só para garantir objeto comment
        cell._comment = None
        cell.comment = None
        cell.hyperlink = None
        cell._hyperlink = None

        # atualiza snapshot em memória (para futuras execuções durante este run)
        snapshot_ant[key] = {
            "rev": info["rev"],
            "tam": info["tam"],
            "ts": info["ts"],
            "row": row,
        }

    wb.save(caminho_excel)
    print("Planilha atualizada:", caminho_excel)

# ---------------------------------------------------------------------------
# Helpers de status
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

# ---------------------------------------------------------------------------
# Carrega linhas anteriores
# ---------------------------------------------------------------------------
def _carregar_snapshot(ws, linha_titulo: int, col_prev: int,
                       estado_anterior: Dict[str, Dict[str, object]] | None = None):
    """
    Percorre coluna da entrega anterior e devolve
    { (base,ext): {rev,tam,ts,row} }
    """
    snap = {}
    for row in range(linha_titulo + 1, ws.max_row + 1):
        nome = ws.cell(row=row, column=col_prev).value
        if not nome:
            continue
        base, ext = _key(nome)
        rev = _extrair_rev(nome)
        dados = estado_anterior.get(f"{base}|{ext}", {}) if estado_anterior else {}
        snap[(base, ext)] = {
            "rev": dados.get("revisao", rev),
            "tam": dados.get("tamanho"),
            "ts": dados.get("timestamp"),
            "row": row,
        }
    return snap

# ---------------------------------------------------------------------------
# Layout inicial / cabeçalhos
# ---------------------------------------------------------------------------
def _abrir_ou_criar_wb(caminho: Path, diretorio_base: str):
    if caminho.exists():
        wb = load_workbook(caminho)
        return wb, wb.active, False
    wb = Workbook()
    ws = wb.active
    ws.title = "GRD"
    _montar_cabecalho(ws, diretorio_base)
    return wb, ws, True


def _montar_cabecalho(ws, diretorio_base: str):
    hoje = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M")
    _merge_ai(ws, 1, "OLIVEIRA ARAÚJO ENGENHARIA", FONT_TITULO)
    _merge_ai(ws, 2, "Lista de arquivos de projetos entregues com controle de revisões")
    _merge_ai(ws, 3, f"Diretório: {diretorio_base}")
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


def _pintar_cabecalho_titulos(ws, linha: int, num_entrega: int):
    # Grupo | Ext | Entregas
    for col in range(_col_idx("K"), _col_idx("K") + 2 + num_entrega):
        cell = ws.cell(row=linha, column=col)
        cell.fill = PatternFill(start_color=COR_TITULO,
                                end_color=COR_TITULO,
                                fill_type="solid")
        cell.alignment = ALIGN_CENTRO
        cell.font = Font(bold=True)
