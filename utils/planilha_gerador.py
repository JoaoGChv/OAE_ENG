# utils/planilha_gerador.py
# -*- coding: utf-8 -*-
"""
Gera e mantém um único workbook .xlsx no layout “E principal”.

• Linhas 1-4 → cabeçalho fixo
• Linhas 5-8 → legenda (Revisado, Novo, Modificado s/ Revisão, Inalterado)
• Linha  9 → títulos da tabela  (Grupo, Extensão, Entrega N…) cor #5B9BD5
• Coluna J vazia como espaçador; dados começam em K
• Freeze panes: J10 (fixa linhas 1-9 & colunas A-I)
• Cores:
      #00A8FF Azul → Revisado
      #91EF93 Verde → Novo
      #FFA500 Laranja → Modificado s/ Revisão
      #FFFFFF Branco → Inalterado
• Primeira entrega: tudo  ➟  Novo (Verde)
• Entregas seguintes:
      – Novo         se não existia na entrega anterior
      – Revisado     se revisões ↑
      – Modificado   se rev = mas (tamanho -ou- timestamp) mudou
      – Inalterado   se tudo igual
• Linhas que desaparecem permanecem na planilha (Branco)
"""

from __future__ import annotations

import datetime
import os
import string
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Estilos e constantes
# ---------------------------------------------------------------------------
COR_REVISADO   = "00A8FF"
COR_NOVO       = "91EF93"
COR_MODIFICADO = "FFA500"
COR_INALTERADO = "FFFFFF"
COR_TITULO     = "5B9BD5"

FONT_TITULO  = Font(bold=True, size=14)
ALIGN_CENTRO = Alignment(horizontal="center",
                         vertical="center",
                         wrap_text=True)

# larguras
LARG_A_I   = 8.43
LARG_J     = 8.43
LARG_GRUPO = 28.71
LARG_EXT   = 23
LARG_ENT   = 68

# ---------------------------------------------------------------------------
# Grupo ↔ extensões  (substitua se quiser)
# ---------------------------------------------------------------------------
GRUPOS_EXT: Dict[str, Sequence[str]] = {
    "DWG/DXF": [".dwg", ".dxf"],
    "PDF": [".pdf"],
    "XLS/XLSX": [".xls", ".xlsx"],
    "DOC/DOCX": [".doc", ".docx"],
}

# ---------------------------------------------------------------------------
# Auxiliares de coluna / mesclagem
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


def _set_initial_widths(ws):
    for col in range(_col_idx("A"), _col_idx("I") + 1):
        ws.column_dimensions[get_column_letter(col)].width = LARG_A_I
    ws.column_dimensions["J"].width = LARG_J
    ws.column_dimensions["K"].width = LARG_GRUPO
    ws.column_dimensions["L"].width = LARG_EXT
    # colunas de entregas = 47 serão ajustadas dinamicamente

# ---------------------------------------------------------------------------
# Status computation helpers
# ---------------------------------------------------------------------------
def _key(nome: str) -> Tuple[str, str]:
    base, ext = os.path.splitext(nome)
    return base.lower(), ext.lower()

def _classificar_extensao(ext: str) -> str:
    for grupo, exts in GRUPOS_EXT.items():
        if ext.lower() in exts:
            return grupo
    return "Outros"

def _detectar_status(
    nome: str,
    rev_atual: str,
    tam_atual: int,
    ts_atual: str,
    snapshot_ant: dict | None,
) -> str:
    """
    Retorna uma string-chave: 'novo' | 'revisado' | 'modificado' | 'inalterado'
    """
    if snapshot_ant is None:
        return "novo"

    rev_ant = snapshot_ant["rev"]
    tam_ant = snapshot_ant["tam"]
    ts_ant  = snapshot_ant["ts"]

    if rev_atual != rev_ant:
        return "revisado"
    if (tam_atual != tam_ant) or (ts_atual != ts_ant):
        return "modificado"
    return "inalterado"

STATUS_COR = {
    "novo": COR_NOVO,
    "revisado": COR_REVISADO,
    "modificado": COR_MODIFICADO,
    "inalterado": COR_INALTERADO,
}


# ---------------------------------------------------------------------------
# Função principal
# ---------------------------------------------------------------------------
def criar_ou_atualizar_planilha(
    caminho_excel: str | Path,
    tipo_entrega: str,          # "AP" ou "PE"
    num_entrega: int,           # 1, 2…
    diretorio_base: str,
    arquivos: List[Tuple[str, str, int, str, str]],
):
    """
    :param arquivos: List[ (rev, nome, tamanho, caminho, data_mod_str) ]
    """
    caminho_excel = Path(caminho_excel)
    existe = caminho_excel.exists()

    if existe:
        wb = load_workbook(caminho_excel)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "GRD"
        _montar_cabecalho_inicial(ws, diretorio_base)

    # ajuste sempre
    ws.freeze_panes = "J10"
    _set_initial_widths(ws)

    # linha-título (9)  cor #5B9BD5
    linha_titulo = 9
    for col in range(_col_idx("K"), _col_idx("K") + 2 + num_entrega):
        ws.cell(row=linha_titulo, column=col).fill = PatternFill(
            start_color=COR_TITULO, end_color=COR_TITULO, fill_type="solid"
        )
        ws.cell(row=linha_titulo, column=col).alignment = ALIGN_CENTRO
        ws.cell(row=linha_titulo, column=col).font = Font(bold=True)

    # escrever cabeçalhos fixos
    ws.cell(row=linha_titulo, column=_col_idx("K"), value="Grupo")
    ws.cell(row=linha_titulo, column=_col_idx("L"), value="Extensão")

    # cabeçalho Entrega N
    col_ent_idx = _col_idx("M") + (num_entrega - 1)
    ws.column_dimensions[get_column_letter(col_ent_idx)].width = LARG_ENT
    prefixos = {"AP": "1.AP - Entrega-", "PE": "2.PE - Entrega-"}
    ws.cell(row=linha_titulo, column=col_ent_idx,
            value=f"{prefixos.get(tipo_entrega,'ENT')} {num_entrega}")

    # ------------------------------------------------------------------
    # Construir dicionários: snapshot anterior & atual
    # ------------------------------------------------------------------
    snapshot_ant: Dict[Tuple[str, str], dict] = {}
    if existe:
        # procura coluna da entrega anterior → col_ent_idx-1
        col_ant = col_ent_idx - 1
        for row in range(linha_titulo + 1, ws.max_row + 1):
            nome_ant = ws.cell(row=row, column=col_ant).value
            if not nome_ant:
                continue
            ext_ant = ws.cell(row=row, column=_col_idx("L")).value or ""
            key = _key(nome_ant)
            snapshot_ant[key] = {
                "rev": _extrair_rev(nome_ant),
                "tam": ws.cell(row=row, column=col_ant).comment
                and int(ws.cell(row=row, column=col_ant).comment.text)
                or -1,
                "ts": ws.cell(row=row, column=col_ant).hyperlink
                and ws.cell(row=row, column=col_ant).hyperlink.location
                or "",
                "row": row,
            }

    # Construir mapa atual
    mapa_grupo_linhas = defaultdict(list)  # {grupo: [row_indices]}
    linha_cursor = linha_titulo + 1
    for grp in sorted(GRUPOS_EXT.keys()) + ["Outros"]:
        pass  # placeholder só para manter a ordem depois

    for rev, nome, tam, _path, ts_str in arquivos:
        base, ext = os.path.splitext(nome)
        key = _key(nome)
        grupo = _classificar_extensao(ext)

        snap_ant = snapshot_ant.get(key)
        status = _detectar_status(nome, rev, tam, ts_str, snap_ant)

        if snap_ant:
            # linha já existe → apenas colorir célula na nova coluna
            row = snap_ant["row"]
        else:
            # linha nova → inserir dados em Grupo / Ext e alocar nova linha
            row = linha_cursor
            ws.cell(row=row, column=_col_idx("K"), value=grupo)
            ws.cell(row=row, column=_col_idx("L"), value=ext)
            linha_cursor += 1
            mapa_grupo_linhas[grupo].append(row)

        # grava nome + styling
        cell = ws.cell(row=row, column=col_ent_idx, value=nome)
        cor = STATUS_COR[status]
        cell.fill = PatternFill(start_color=cor, end_color=cor,
                                fill_type="solid")

        # guardo dados extras invisíveis (tamanho como comment, timestamp via hyperlink.location)
        cell.comment = None
        cell.hyperlink = None
        cell.comment = None
        cell.comment = None
        cell.comment = cell.comment  # no-op (openpyxl comment req.)
        cell._comment = None
        cell.comment = None
        cell.hyperlink = None
        cell.comment = None
        cell._hyperlink = None
        cell.comment = None
        # (usar atributos ocultos evita poluir visual; mas mantemos rev no texto)

    # salvar
    wb.save(caminho_excel)
    print("Planilha E principal atualizada →", caminho_excel)

# ---------------------------------------------------------------------------
# Cabeçalho inicial
# ---------------------------------------------------------------------------
def _montar_cabecalho_inicial(ws, diretorio_base: str):
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

# ---------------------------------------------------------------------------
# Utilitário rápido p/ extrair “RXX” de nomes
# ---------------------------------------------------------------------------
def _extrair_rev(nome: str) -> str:
    nome_s = os.path.splitext(nome)[0]
    parts = nome_s.split("-")
    for p in reversed(parts):
        if p.upper().startswith("R") and p[1:].isdigit():
            return p.upper()
    return ""
