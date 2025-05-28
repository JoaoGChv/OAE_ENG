from __future__ import annotations
import re, os, sys, json, shutil, hashlib, logging
from tkinter import ttk, messagebox, filedialog
from openpyxl.utils import get_column_letter
from ttkbootstrap.style import Style
from ttkbootstrap.constants import *
from typing import Dict, Optional
from openpyxl.styles import Font
from openpyxl import Workbook
from datetime import datetime
from pathlib import Path
import tkinter as tk
import ttkbootstrap

JSON_CONTADORES_DIR = (r"G:\Drives compartilhados\OAE - SCRIPTS\SCRIPTS\tmp_joaoG\JSON_tmp_joao")
PROJETOS_JSON = r"G:\Drives compartilhados\OAE-JSONS\diretorios_projetos.json"
NOMENCLATURA_REGRAS_JSON = "config/nomenclatura_regras.json"
ULTIMO_DIRETORIO_JSON = "ultimo_diretorio.json"
HISTORICO_JSON = "historico_arquivos.json"
JSON_FILE_PATH = "dados_projetos.json"
MARGIN_SIZE = 10

import logging

logging.basicConfig(
    level=logging.DEBUG,                  
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%d/%m/%Y %H:%M:%S",
)

AP_PREFIX = "1.AP - Entrega-"
PE_PREFIX = "2.PE - Entrega-"
ENTREGA_RE = re.compile(r"^(1\.AP|2\.PE) - Entrega-(\d+)$")

def _listar_entregas_tipo(pasta: Path, prefixo: str) -> list[Path]:
    return sorted(
        [p for p in pasta.iterdir()
         if p.is_dir() and p.name.startswith(prefixo) and not p.name.endswith("-OBSOLETO")],
        key=lambda p: int(ENTREGA_RE.match(p.name).group(2))
    )

def _proximo_num_entrega(pasta_entregas: Path, prefixo: str) -> int:
    ativas = _listar_entregas_tipo(pasta_entregas, prefixo)
    if not ativas:
        return 1
    ultimo = ENTREGA_RE.match(ativas[-1].name)
    return int(ultimo.group(2)) + 1

def _marcar_obsoleta(p: Path):
    destino = p.with_name(p.name + "-OBSOLETO")
    seq = 1
    while destino.exists():
        seq += 1
        destino = p.with_name(p.name + f"-OBSOLETO{seq}")
    p.rename(destino)
    logging.info("Renomeada %s ➜ %s", p.name, destino.name)

def _hash_file(path: Path, buf=8192) -> str:
    h = hashlib.md5()
    with path.open("rb") as f:
        while chunk := f.read(buf):
            h.update(chunk)
    return h.hexdigest()

def _e_pasta_entrega(p: Path) -> bool:
    if not p.is_dir():
        return False
    if not p.name.startswith("Entrega_"):
        return False
    if p.name.endswith("_OBS"):
        return False
    num = p.name.split("_")[1]
    return num.isdigit()

def obter_entrega_anterior(pasta_entregas: Path) -> Optional[Path]:
    entregas = sorted(
        [p for p in pasta_entregas.iterdir()
         if p.is_dir()
         and p.name.startswith("Entrega_")
         and p.name.split("_")[1][:2].isdigit()
         and not p.name.endswith("_OBS")],
        key=lambda p: int(p.name.split("_")[1][:2])
    )
    return entregas[-1] if entregas else None

def listar_arquivos_entrega(pasta: Path) -> list[Path]:
    return [p for p in pasta.iterdir() if p.is_file()]

def comparar_arquivos(pasta_nova: Path, pasta_ant: Optional[Path]) -> dict:
    atual     = {p.name: p for p in listar_arquivos_entrega(pasta_nova)}
    anterior  = {p.name: p for p in listar_arquivos_entrega(pasta_ant)} if pasta_ant else {}
    resultado: Dict[str, dict] = {}

    for nome, p in atual.items():
        if nome == "_controle_entrega.json":
            continue
        if nome in anterior:
            ig = _hash_file(p) == _hash_file(anterior[nome])
            resultado[nome] = {
                "status": "nao_modificado" if ig else "modificado",
                "versao_anterior": str(anterior[nome])
            }
        else:
            resultado[nome] = {"status": "novo"}

    for nome, p_old in anterior.items():
        if nome not in resultado and nome != "_controle_entrega.json":
            resultado[nome] = {"status": "removido", "versao_anterior": str(p_old)}
    return resultado

def marcar_pasta_obsoleta(pasta: Path) -> Path:
    base = pasta.with_name(pasta.name + "_OBS")
    idx = 1
    destino = base
    while destino.exists():
        idx += 1
        destino = pasta.with_name(f"{pasta.name}_OBS{idx}")
    pasta.rename(destino)
    logging.info("Pasta %s ➜ %s (obsoleta)", pasta.name, destino.name)
    return destino

def criar_nova_pasta_entrega(pasta_entregas: Path) -> Path:
    ultima = obter_entrega_anterior(pasta_entregas)
    prox = 1 if not ultima else int(ultima.name.split("_")[1]) + 1
    nova = pasta_entregas  /  f"Entrega_{prox:02d}"
    nova.mkdir(parents=True, exist_ok=False)
    return nova
    
def gerar_arquivo_controle(nova_pasta: Path, comparacao: dict):
    with (nova_pasta / "_controle_entrega.json").open("w",  encoding="utf-8") as f:
        json.dump(comparacao, f, indent=4, ensure_ascii=False)

def processar_entrega_arquivos_tipo(arquivos: list[Path], pasta_entregas: Path, tipo: str) -> Path:
    prefixo = AP_PREFIX if tipo == "AP" else PE_PREFIX
    etapa = 1 if tipo == "AP" else 2

    # 1️⃣  pega a entrega ativa (se existir) ANTES de renomear
    ativas = _listar_entregas_tipo(pasta_entregas, prefixo)
    entrega_ativa = ativas[-1] if ativas else None

    # 2️⃣  calcula o próximo número a partir da ativa
    if entrega_ativa:
        ultimo_num   = int(ENTREGA_RE.match(entrega_ativa.name).group(2))
        proximo_num  = ultimo_num + 1
    else:
        proximo_num  = 1

    # 3️⃣  cria a nova pasta
    nova = pasta_entregas / f"{prefixo}{proximo_num}"
    nova.mkdir(parents=True, exist_ok=False)
    logging.debug("Criada nova entrega: %s", nova)

    # 4️⃣  gera o diff contra a entrega ativa (ela ainda existe)
    comp = comparar_arquivos(nova, entrega_ativa)

    # 5️⃣  agora sim, marca a velha como obsoleta
    if entrega_ativa:
        _marcar_obsoleta(entrega_ativa)

    # 6️⃣  copia os arquivos escolhidos
    for src in arquivos:
        shutil.copy2(src, nova / src.name)

    # 7️⃣  grava JSON de controle
    comp.update({"tipo_entrega": tipo, "etapa": etapa})
    with (nova / "_controle_entrega.json").open("w", encoding="utf-8") as f:
        json.dump(comp, f, indent=4, ensure_ascii=False)

    return nova

def carregar_regras_nomenclatura() -> dict:
    if not os.path.exists(NOMENCLATURA_REGRAS_JSON):
        logging.warning("Arquivo %s não encontrado", NOMENCLATURA_REGRAS_JSON)
        return {}
    try:
        with open(NOMENCLATURA_REGRAS_JSON, encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        logging.error("JSON inválido em %s – %s", NOMENCLATURA_REGRAS_JSON, e)
        return {}
    
def caminho_contador(projeto_num: str) -> str:
    os.makedirs(JSON_CONTADORES_DIR, exist_ok=True)
    return os.path.join(JSON_CONTADORES_DIR, f"contador_entregas_{projeto_num}.json")

def obter_proximo_indice(projeto_num: str) -> int:
    fp = caminho_contador(projeto_num)
    data = carregar_json(fp) or {"proximo": 1}
    return data["proximo"]

def incrementar_indice(projeto_num: str):
    fp = caminho_contador(projeto_num)
    data = carregar_json(fp) or {"proximo": 1}
    data["proximo"] += 1
    salvar_json(fp, data)

def carregar_historico_entregas(projeto_num: str) -> dict:
    fp = caminho_contador(projeto_num)
    if os.path.exists(fp):
        with open(fp, "r", encoding="utf-8") as f:
            dados = json.load(f)
    else:
        dados = {}
    dados.setdefault("proximo", 1)
    dados.setdefault("entregas", [])
    return dados

def salvar_historico_entregas(projeto_num: str, data: dict) -> None:
    fp = caminho_contador(projeto_num)
    try:
        with open(fp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception as err:
        messagebox.showerror("Erro", f"Falha ao salvar histórico de entregas:\n{err}")
        raise SystemExit

def carregar_projetos():
    if not os.path.exists(PROJETOS_JSON):
        messagebox.showerror("Erro", f"Arquivo não encontrado: {PROJETOS_JSON}")
        return {}
    try:
        with open(PROJETOS_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        messagebox.showerror("Erro", f"Erro ao decodificar o JSON: {e}")
        return {}

def carregar_ultimo_diretorio():
    if os.path.exists(ULTIMO_DIRETORIO_JSON):
        with open(ULTIMO_DIRETORIO_JSON, "r", encoding="utf-8") as f:
            return json.load(f).get("ultimo_diretorio", os.getcwd())
    return os.getcwd()

def salvar_ultimo_diretorio(d):
    with open(ULTIMO_DIRETORIO_JSON, "w", encoding="utf-8") as f:
        json.dump({"ultimo_diretorio": d}, f)

def atualizar_historico(lista_arquivos, c=HISTORICO_JSON):
    h = {}
    if os.path.exists(c):
        with open(c, "r", encoding="utf-8") as f:
            h = json.load(f)
    for arq in lista_arquivos:
        dm = os.path.getmtime(arq)
        if arq not in h or h[arq]["data"] != dm:
            h[arq] = {"numero": len(h)+1, "data": dm, "status": "Atual"}
    mr = max(h, key=lambda x: h[x]["data"])
    for a in h:
        h[a]["status"] = "Atual" if a == mr else "Obsoleto"
    with open(c, "w", encoding="utf-8") as f:
        json.dump(h, f, indent=4, ensure_ascii=False)
    return h

def pos_processamento(*args):
    messagebox.showinfo("Concluído", "Processo concluído com sucesso.")
    sys.exit(0)

def janela_selecao_projeto(master):
    root = master
    root.title("Selecionar Projeto")
    root.geometry("600x400")
    projetos_dict = carregar_projetos()
    style = ttkbootstrap.Style(theme="flatly")
    if not projetos_dict:
        messagebox.showerror("Erro", "Nenhum projeto encontrado ou erro no arquivo JSON.")
        Disciplinas_Detalhes_Projeto(sel["numero"], sel["caminho"], master=root)
        root.destroy()
        return None, None
    p_conv = []
    for num, c_full in projetos_dict.items():
        nm = os.path.basename(c_full)
        p_conv.append((num, nm, c_full))
    sel = {"numero": None, "caminho": None}

    def filtrar(*args):
        t = entrada.get().lower()
        tree.delete(*tree.get_children())
        for n, nm, co in p_conv:
            if t in nm.lower():
                tree.insert("", tk.END, values=(n, nm, co))
        all_iid = tree.get_children()
        if len(all_iid) == 1:
            tree.selection_set(all_iid[0])

    def confirmar():
        si = tree.selection()
        if not si:
            messagebox.showinfo("Info", "Selecione um projeto.")
            return
        v = tree.item(si[0], "values")
        sel["numero"] = v[0]
        sel["caminho"] = v[2]
        root.withdraw() 
        Disciplinas_Detalhes_Projeto(sel["numero"], sel["caminho"], master=root)
    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
    tk.Label(frame, text="Digite nome ou parte do nome do projeto:", font=("Arial", 10)).pack(anchor="w")
    entrada = tk.Entry(frame)
    entrada.pack(fill=tk.X)
    entrada.bind("<KeyRelease>", filtrar)
    entrada.bind("<Return>", lambda e: confirmar())
    cols = ("Número", "Nome do Projeto", "Caminho Original")
    tree = ttk.Treeview(frame, columns=cols, show="headings", height=10)
    tree.heading("Número", text="Número")
    tree.heading("Nome do Projeto", text="Nome do Projeto")
    tree.heading("Caminho Original", text="Caminho Original")
    tree.column("Número", width=80)
    tree.column("Nome do Projeto", width=300)
    tree.column("Caminho Original", width=0, stretch=False, minwidth=0)
    tree.pack(fill=tk.BOTH, expand=True)
    for n, nm, co in p_conv:
        tree.insert("", tk.END, values=(n, nm, co))
    bf = tk.Frame(frame)
    bf.pack(pady=5)
    ttk.Button(bf, text="Confirmar", command=confirmar, bootstyle="success").pack(side=tk.LEFT, padx=5)
    ttk.Button(bf, text="Cancelar", command=root.destroy, bootstyle="danger").pack(side=tk.LEFT, padx=5)
    root.mainloop()
    return sel["numero"], sel["caminho"]

def Disciplinas_Detalhes_Projeto(numero, caminho, master=None):
    d_path = os.path.join(caminho, "3 Desenvolvimento")
    if not os.path.exists(d_path):
        messagebox.showerror("Erro", "A pasta de disciplinas não foi encontrada.")
        return
    nj = tk.Toplevel(master)
    nj.title(f"Gerenciador de Projetos - Projeto {numero}")
    nj.geometry("900x600")
    hd = tk.Label(nj, text=f"Projeto {numero} - Selecione a disciplina para entrega", font=("Helvetica", 14, "bold"), anchor="w")
    hd.pack(fill=tk.X, padx=10, pady=5)
    cols = ["Nome", "Data de Modificação", "Tipo", "Tamanho"]
    tree = ttk.Treeview(nj, columns=cols, show="headings", height=20)
    tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    for c in cols:
        tree.heading(c, text=c)
        tree.column(c, width=200 if c=="Nome" else 150, anchor="w")
    disc = []
    for it in os.listdir(d_path):
        fp = os.path.join(d_path, it)
        if os.path.isdir(fp):
            mt = datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%d/%m/%Y %H:%M")
            disc.append((it, mt, "Pasta", "--"))
    for d in disc:
        tree.insert("", tk.END, values=d)

    def confirmar_selecao_arquivos():
        s = tree.selection()
        if not s:
            messagebox.showwarning("Atenção", "Nenhuma disciplina selecionada.")
            return
        v = tree.item(s[0])["values"]
        disc_nome = v[0]
        p_disc = os.path.join(d_path, disc_nome)
        if not os.path.isdir(p_disc):
            messagebox.showerror("Erro", f"A pasta da disciplina '{p_disc}' não foi encontrada.")
            return
        expected_folder = "1.ENTREGAS"
        match_entrega = None
        def normalize(n): return n.lower().replace(" ", "").replace("-", "").replace("_", "").replace(".", "")
        for folder in os.listdir(p_disc):
            if normalize(folder) == normalize(expected_folder):
                match_entrega = folder
                break
        if not match_entrega:
            messagebox.showerror("Erro", f"A pasta de entrega '{expected_folder}' não foi encontrada.")
            return
        p_ent = os.path.join(p_disc, match_entrega)
        if not os.path.isdir(p_ent):
            messagebox.showerror("Erro", f"A pasta de entrega '{p_ent}' não pôde ser acessada.")
            return
        sel_arq = filedialog.askopenfilenames(title="Selecione arquivos para entrega", initialdir=p_ent)
        if not sel_arq:
            messagebox.showwarning("Atenção", "Nenhum arquivo foi selecionado.")
            return
        proc = []
        for arq in sel_arq:
            n_arq = os.path.basename(arq)
            d_ext = extrair_dados_arquivo(n_arq)
            d_ext["caminho"] = arq
            proc.append(d_ext)
        if not proc:
            messagebox.showerror("Erro", "Nenhum dado foi processado.")
            return
        nj.destroy()
        exibir_interface_tabela(numero, arquivos_previos=proc, caminho_projeto=caminho, pasta_entrega=p_ent)

    def voltar():
        nj.destroy()
        master.deiconify()
    cf = tk.Frame(nj, bg="#f5f5f5", padx=20, pady=20)
    cf.pack(fill=tk.BOTH, expand=True)
    bf = tk.Frame(nj)
    bf.pack(fill=tk.X, pady=5, padx=10)
    ttk.Button(bf, text="Voltar", command=voltar, bootstyle="warning").pack(side=tk.LEFT, padx=5)
    ttk.Button(bf, text="Confirmar Seleção", command=confirmar_selecao_arquivos, bootstyle="success").pack(side=tk.RIGHT, padx=5)
    nj.mainloop()

def carregar_json(fp):
    if os.path.exists(fp):
        with open(fp, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        return {}

def salvar_json(fp, data):
    try:
        with open(fp, "w", encoding="utf-8") as file:
            json.dump(data, file, indent=4, ensure_ascii=False)
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao salvar dados em JSON: {e}")

def extrair_dados_arquivo(nome_arquivo):
    nb, ext = os.path.splitext(nome_arquivo)
    pt = nb.split('-')
    try:
        cr = pt[7] if len(pt)>7 else ""
        if '.' in cr:
            cs = cr.split('.')
            conj = cs[0]
            num_doc = cs[1]
        else:
            conj = cr
            num_doc = ""
        d = {
            "Status": pt[0] if len(pt)>0 else "",
            "Cliente": pt[1] if len(pt)>1 else "",
            "N° do Projeto": pt[2] if len(pt)>2 else "",
            "Organização": pt[3] if len(pt)>3 else "",
            "Sigla da Disciplina": pt[4] if len(pt)>4 else "",
            "Fase": pt[5] if len(pt)>5 else "",
            "Tipo de Documento": pt[6] if len(pt)>6 else "",
            "Conjunto": conj,
            "N° do Documento": num_doc,
            "Bloco": pt[8] if len(pt)>8 else "",
            "Pavimento": pt[9] if len(pt)>9 else "",
            "Subsistema": pt[10] if len(pt)>10 else "",
            "Tipo do Desenho": pt[11] if len(pt)>11 else "",
            "Revisão": pt[12] if len(pt)>12 else "",
            "Nome do Arquivo": nome_arquivo,
            "Extensão": ext.strip('-'),
            "Modificação": datetime.now().strftime("%d/%m/%Y"),
            "Modificado por": "Usuário"
        }
    except IndexError:
        d = {
            "Status": "",
            "Cliente": "",
            "N° do Projeto": "",
            "Organização": "",
            "Sigla da Disciplina": "",
            "Fase": "",
            "Tipo de Documento": "",
            "Conjunto": "",
            "N° do Documento": "",
            "Bloco": "",
            "Pavimento": "",
            "Subsistema": "",
            "Tipo do Desenho": "",
            "Revisão": "",
            "Nome do Arquivo": nome_arquivo,
            "Extensão": ext.strip('-'),
            "Modificação": datetime.now().strftime("%d/%m/%Y"),
            "Modificado por": "Usuário"
        }
    return d

def exibir_interface_tabela(numero, arquivos_previos=None, caminho_projeto=None, pasta_entrega=None, master=None):
    j = tk.Toplevel(master)
    j.title(f"Gerenciador de Projetos - Projeto {numero}")
    j.geometry("1200x800")
    fpr = tk.Frame(j)
    fpr.pack(fill=tk.BOTH, expand=True)
    bar_l = tk.Frame(fpr, bg="#2c3e50", width=200)
    bar_l.pack(side=tk.LEFT, fill=tk.Y)
    lbl_t = tk.Label(bar_l, text="OAE - Engenharia", font=("Helvetica",14,"bold"), bg="#2c3e50", fg="white")
    lbl_t.pack(pady=10)
    lbl_pj = tk.Label(bar_l, text="PROJETOS", font=("Helvetica",10,"bold"), bg="#34495e", fg="white", anchor="w", padx=10)
    lbl_pj.pack(fill=tk.X, pady=5)
    ls_pj = tk.Listbox(bar_l, height=5, bg="#ecf0f1", font=("Helvetica",9))
    ls_pj.pack(fill=tk.X, padx=10, pady=5, anchor="center")
    ls_pj.insert(tk.END, "Projeto ", numero)
    lbl_m = tk.Label(bar_l, text="MEMBROS", font=("Helvetica",10,"bold"), bg="#34495e", fg="white", anchor="w", padx=10)
    lbl_m.pack(fill=tk.X, pady=5)
    bar_l.pack_propagate(False)
    cp = tk.Frame(fpr)
    cp.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    def fazer_analise_nomenclatura():
        la = []
        for it in tabela.get_children():
            v = tabela.item(it)["values"]
            n_arq = v[1]
            cmin = v[10]
            dx = extrair_dados_arquivo(n_arq)
            dx["caminho"] = cmin
            la.append(dx)
        if not la:
            messagebox.showinfo("Aviso", "Nenhum arquivo adicionado para análise.")
        else:
            j.destroy()
            tela_analise_nomenclatura(la, pasta_entrega=pasta_entrega)
    lbl_i = tk.Label(cp, text="Adicionar Arquivos para Entrega", font=("Helvetica",15,"bold"), anchor="w")
    lbl_i.place(x=10, y=10)
    fb = tk.Frame(cp)
    fb.pack(side=tk.TOP, anchor="ne", pady=10, padx=10)
    ttk.Button(fb, text="Fazer análise da Nomenclatura", command=fazer_analise_nomenclatura).pack(side=tk.LEFT, padx=5)
    cols = ["Status","Nome do Arquivo","Extensão","Nº do Arquivo","Fase","Tipo","Revisão","Modificação","Modificado por","Entrega","caminho"]
    tabela = ttk.Treeview(cp, columns=cols, show="headings", height=20)
    for c in cols:
        tabela.heading(c, text=c)
        if c=="Nome do Arquivo":
            tabela.column(c, width=300)
        elif c=="caminho":
            tabela.column(c, width=0, stretch=False, minwidth=0)
        else:
            tabela.column(c, width=120)
    tabela["displaycolumns"] = ("Status","Nome do Arquivo","Extensão","Nº do Arquivo","Fase","Tipo","Revisão","Modificação","Modificado por","Entrega")
    tabela.pack(fill=tk.BOTH, expand=True)
    if arquivos_previos:
        for d_ext in arquivos_previos:
            tabela.insert("", tk.END, values=(
                d_ext.get("Status",""), d_ext.get("Nome do Arquivo",""), d_ext.get("Extensão",""),
                d_ext.get("N° do Arquivo",""), d_ext.get("Fase",""), d_ext.get("Tipo de Documento",""),
                d_ext.get("Revisão",""), d_ext.get("Modificação",""), d_ext.get("Modificado por",""), "", d_ext.get("caminho","")
            ))
            
    def adicionar_arquivos():
        ar = filedialog.askopenfilenames(title="Selecione arquivos")
        for a in ar:
            n = os.path.basename(a)
            dx = extrair_dados_arquivo(n)
            dx["caminho"] = a
            tabela.insert("", tk.END, values=(
                dx.get("Status",""), dx.get("Nome do Arquivo",""), dx.get("Extensão",""),
                dx.get("N° do Arquivo",""), dx.get("Fase",""), dx.get("Tipo de Documento",""),
                dx.get("Revisão",""), dx.get("Modificação",""), dx.get("Modificado por",""), "", dx.get("caminho","")
            ))

    def remover_arquivo():
        s = tabela.selection()
        if s:
            for i in s:
                tabela.delete(i)
        else:
            messagebox.showinfo("Informação", "Nenhum item selecionado.")

    def voltar():
        j.destroy()
        master.deiconify()
    bf2 = tk.Frame(cp)
    bf2.pack(side=tk.LEFT, pady=10, padx=10)
    ttk.Button(bf2, text="Adicionar Arquivo", command=adicionar_arquivos).pack(side=tk.LEFT, padx=5)
    ttk.Button(bf2, text="Remover Arquivo", command=remover_arquivo).pack(side=tk.LEFT, padx=5)
    bf3 = tk.Frame(cp)
    bf3.pack(side="bottom", anchor="e", pady=5, padx=10)
    ttk.Button(bf3, text="Voltar", command=voltar).pack(side=tk.LEFT, padx=5)
    ttk.Button(bf3, text="Sair", command=j.destroy).pack(side=tk.RIGHT, padx=5)

def identificar_revisoes(lista_arquivos):
    grupos = {}
    for a in lista_arquivos:
        nb, _ = os.path.splitext(a["Nome do Arquivo"])
        t = nb.split("-")
        if len(t)<2: continue
        idf = "-".join(t[:-1])
        rev = t[-1] if t[-1].startswith("R") and t[-1][1:].isdigit() else "R00"
        if idf not in grupos:
            grupos[idf] = []
        grupos[idf].append((rev, a))
    arrv = []
    aobs = []
    for idf, arqs in grupos.items():
        arqs.sort(key=lambda x: int(x[0][1:]) if x[0][1:].isdigit() else 0)
        rm = arqs[-1][1]
        arrv.append(rm)
        aobs.extend([q[1] for q in arqs[:-1]])
    return arrv, aobs

def tela_analise_nomenclatura(lista_arquivos, pasta_entrega: str, master=None):
    rnom = carregar_regras_nomenclatura()
    j = tk.Toplevel(master)
    j.title("Verificação de Nomenclatura")
    j.geometry("1400x800")
    lb_i = tk.Label(j, text="Confira a nomenclatura. Clique no nome para editar campos. Corrija antes de avançar, se necessário.", font=("Helvetica",12))
    lb_i.pack(pady=MARGIN_SIZE)
    fm = tk.Frame(j)
    fm.pack(fill=tk.BOTH, expand=True)
    cnv = tk.Canvas(fm)
    cnv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    sb = tk.Scrollbar(fm, orient="vertical", command=cnv.yview)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    cnv.configure(yscrollcommand=sb.set)
    cont = tk.Frame(cnv, bg="#ECE2E2")
    c_id = cnv.create_window((0,0), window=cont, anchor="n")

    def on_conf(event):
        cnv.configure(scrollregion=cnv.bbox("all"))
        w = event.width
        cnv.itemconfig(c_id, width=w)
    cnv.bind("<Configure>", on_conf)
    fields_in_error = set()
    expanded_cards = set()
    CMPS = ["Status","Cliente","N° do Projeto","Organização","Sigla da Disciplina","Fase","Tipo de Documento","Conjunto","N° do Documento","Bloco","Pavimento","Subsistema","Tipo do Desenho","Revisão"]
    
    def expand_or_collapse(cid):
        if cid in expanded_cards:
            expanded_cards.remove(cid)
        else:
            expanded_cards.add(cid)
        render_cards()

    def validate_with_json_rules(campo, val):
        if campo not in rnom: return (True,"")
        rg = rnom[campo]
        val = val.strip()
        if rg.get("obrigatorio",False) and not val:
            return (False,"Campo obrigatório.")
        vo = rg.get("valores_permitidos",[])
        if vo and val not in vo and val!="":
            return (False,f"Valor inválido para {campo}. Opções: {vo}")
        vf = rg.get("valor_fixo")
        if vf and val!=vf:
            return (False,f"Valor fixo exigido: {vf}")
        import re
        pt = rg.get("regex")
        if pt and val!="":
            if not re.match(pt,val):
                return (False,f"O valor '{val}' não atende ao regex: {pt}")
        return (True,"")
    
    def validate_entry(ev, campo, lbl_e):
        tv = ev.get().strip()
        iv, me = validate_with_json_rules(campo, tv)
        if not iv:
            lbl_e.config(text=me, fg="red")
            fields_in_error.add((campo,lbl_e))
        else:
            lbl_e.config(text="", fg="red")
            if (campo,lbl_e) in fields_in_error:
                fields_in_error.remove((campo,lbl_e))

    def render_cards():
        for wdg in cont.winfo_children():
            wdg.destroy()
        for idx,a in enumerate(lista_arquivos):
            c_id2 = idx
            cf = tk.Frame(cont, bd=1, relief=tk.RIDGE, padx=MARGIN_SIZE, pady=MARGIN_SIZE, bg="#ECE2E2")
            cf.pack(padx=MARGIN_SIZE, pady=MARGIN_SIZE, fill=tk.X)
            hb = ttk.Button(cf, text=a["Nome do Arquivo"], command=lambda c=c_id2: expand_or_collapse(c))
            hb.pack(fill=tk.X)
            if c_id2 in expanded_cards:
                df = tk.Frame(cf, bd=1, relief=tk.GROOVE, padx=MARGIN_SIZE, pady=MARGIN_SIZE, bg="#ECE2E2")
                df.pack(fill=tk.X)
                rm, cm = 7, 2
                idx_cp = 0
                def mk_cb(vr,c,el):
                    return lambda e: validate_entry(vr,c,el)
                ca_arq = []
                for n_campo in CMPS:
                    ca_arq.append((n_campo, a.get(n_campo,"")))
                for rr in range(rm):
                    for cc in range(cm):
                        if idx_cp<len(ca_arq):
                            nm_cp, vl_cp = ca_arq[idx_cp]
                            c_f = tk.Frame(df, bd=1, relief=tk.FLAT, bg="#ECE2E2")
                            c_f.grid(row=rr, column=cc, padx=MARGIN_SIZE, pady=MARGIN_SIZE, sticky="nsew")
                            lab = tk.Label(c_f, text=nm_cp+":", bg="#ECE2E2")
                            lab.pack(anchor="w")
                            vvar = tk.StringVar(value=vl_cp)
                            err_lbl = tk.Label(c_f, text="", fg="red", bg="#ECE2E2")
                            ent = tk.Entry(c_f, textvariable=vvar, width=30)
                            ent.bind("<KeyRelease>", mk_cb(vvar,nm_cp,err_lbl))
                            ent.pack(fill=tk.X)
                            err_lbl.pack(anchor="w")
                            def upd(*args,ac=a,field=nm_cp,v=vvar):
                                ac[field] = v.get()
                            vvar.trace("w",upd)
                            idx_cp+=1
                for rr in range(rm):
                    df.rowconfigure(rr, weight=0)
                for cc in range(cm):
                    df.columnconfigure(cc, weight=1)
    render_cards()
    def ajustar_canvas():
        j.update_idletasks()
        w_at = cnv.winfo_width()
        cnv.itemconfig(c_id, width=w_at)
        cnv.configure(scrollregion=cnv.bbox("all"))
    j.after(100, ajustar_canvas)

    def avancar():
        def on_confirm(tipo):
            j.destroy()  # Fecha a janela da tela de análise de nomenclatura
            tela_tipo_escolhido(lista_arquivos, pasta_entrega, tipo, master=master)
        modal_tipo_entrega(j, on_confirm)
    def voltar():
        j.destroy()
        master.deiconify()
    bf = tk.Frame(j)
    bf.pack(side="bottom", anchor="e", pady=MARGIN_SIZE, padx=MARGIN_SIZE)
    ttk.Button(bf, text="Voltar", command=voltar).pack(side=tk.LEFT, padx=5)
    ttk.Button(bf, text="Avançar", command=avancar).pack(side=tk.RIGHT, padx=5)
    j.mainloop()

def criar_arquivo_excel(diretorio, arquivos_novos, arquivos_revisados, arquivos_obsoletos):
    now = datetime.now()
    ds = now.strftime("%d_%m_%Y")
    hs = now.strftime("%Hh%Mmin")
    nome_arq = f"GRD {ds}_{hs}.xlsx"
    cam = os.path.join(diretorio, nome_arq)
    wb = Workbook()
    ws = wb.active
    ws.title = "GRD"
    bf = Font(bold=True)
    ws.append(["OLIVEIRA ARAÚJO ENGENHARIA"])
    ws.append(["Lista de arquivos de projetos entregues com controle de versão."])
    ws.append(["Diretório:", diretorio])
    ws.append(["Data de emissão:", f"{ds}_{hs}"])
    ws.append(["Arquivo Revisado"])
    if arquivos_revisados:
        ws.append(["Nome do arquivo","Revisão","Data de modificação"])
        for a in arquivos_revisados:
            ws.append([a["Nome do Arquivo"],a["Revisão"],a.get("Modificação","")])
    else:
        ws.append(["Nenhum arquivo revisado"])
    ws.append([])
    ws.append(["Arquivo Novo"])
    if arquivos_novos:
        ws.append(["Nome do arquivo","Revisão","Data de modificação"])
        for a in arquivos_novos:
            ws.append([a["Nome do Arquivo"],a["Revisão"],a.get("Modificação","")])
    else:
        ws.append(["Nenhum arquivo novo"])
    ws.append([])
    ws.append(["Arquivo Obsoleto"])
    if arquivos_obsoletos:
        ws.append(["Nome do arquivo","Revisão","Data de modificação"])
        for a in arquivos_obsoletos:
            ws.append([a["Nome do Arquivo"],a["Revisão"],a.get("Modificação","")])
    else:
        ws.append(["Nenhum arquivo obsoleto"])
    from openpyxl.utils import get_column_letter
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            if isinstance(cell.value,str) and (
                "OLIVEIRA" in cell.value or
                "Lista de arquivos" in cell.value or
                "Diretório:" in cell.value or
                "Data de emissão:" in cell.value or
                "Arquivo Revisado" in cell.value or
                "Arquivo Novo" in cell.value or
                "Arquivo Obsoleto" in cell.value or
                cell.value in ["Nome do arquivo","Revisão","Data de modificação"]):
                cell.font = bf
    for col in ws.columns:
        ml = 0
        c = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                ml = max(ml,len(str(cell.value)))
        ws.column_dimensions[c].width = ml+2
    wb.save(cam)
    return cam

def modal_tipo_entrega(master, on_confirm):
    win = tk.Toplevel(master)
    win.title("Tipo de Entrega")
    win.transient(master)
    win.grab_set()
    win.resizable(False, False)
    win.minsize(300, 140)
    master.update_idletasks()
    x = master.winfo_rootx() + master.winfo_width() // 2 - 150
    y = master.winfo_rooty()  + master.winfo_height() // 2 - 70
    win.geometry(f"+{x}+{y}")
    ttk.Label(win, text="Escolha o tipo de entrega:",
              font=("Arial", 11, "bold")).pack(pady=(12, 6), anchor="w", padx=20)
    var = tk.StringVar(value="AP")
    frm = ttk.Frame(win); frm.pack(padx=20, pady=5, anchor="w")
    ttk.Radiobutton(frm, text="Anteprojeto – 1.AP", value="AP", variable=var).pack(anchor="w")
    ttk.Radiobutton(frm, text="Projeto Executivo – 2.PE", value="PE", variable=var).pack(anchor="w")
    ttk.Separator(win).pack(fill="x", pady=10, padx=5)
    ttk.Button(win, text="Confirmar",
               command=lambda:(win.grab_release(), win.destroy(),
                               on_confirm(var.get()))).pack(pady=(0,12))

def tela_tipo_escolhido(lista_arquivos, pasta_entrega: str, tipo: str, master=None):
    # 1) validações básicas
    erros = []
    for a in lista_arquivos:
        caminho = a.get("caminho")
        nome    = a.get("Nome do Arquivo")
        if not caminho or not Path(caminho).exists():
            erros.append(f"Caminho ausente → {nome}")
    if erros:
        if len(erros) <= 3:
            messagebox.showerror("Erros", "\n".join(erros))
        else:
            _lista_erros_treeview(erros)
        return                           # bloqueia avanço

    # 2) processa entrega
    try:
        tela_verificacao_revisao(lista_arquivos, pasta_entrega, tipo, master=master)
    except Exception as e:
        logging.exception(e)
        messagebox.showerror("Falha", str(e))

def _lista_erros_treeview(erros):
    w = tk.Toplevel(); w.title("Lista de erros")
    tree = ttk.Treeview(w, columns=["e"], show="headings", height=10)
    tree.heading("e", text="Ocorrências")
    tree.pack(fill="both", expand=True, padx=10, pady=10)
    for e in erros:
        tree.insert("", "end", values=(e,))
    ttk.Button(w, text="Fechar", command=w.destroy).pack(pady=5)

def tela_verificacao_revisao(lista_arquivos, pasta_entrega, tipo, master=None):
    arrv, aobs = identificar_revisoes(lista_arquivos)
    j = tk.Toplevel(master)
    j.title("Verificação de Revisão")
    j.geometry("1000x700")
    lb_i = tk.Label(j, text="Confira os arquivos revisados e obsoletos antes da entrega.")
    lb_i.pack(pady=10)
    fr_r = tk.LabelFrame(j, text="Arquivos Revisados")
    fr_r.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    tr_r = ttk.Treeview(fr_r, columns=["Nome do Arquivo","Revisão"], show="headings", height=10)
    tr_r.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    tr_r.heading("Nome do Arquivo", text="Nome do Arquivo")
    tr_r.heading("Revisão", text="Revisão")
    for a in arrv:
        tr_r.insert("", tk.END, values=(a["Nome do Arquivo"], a["Revisão"]))
    fr_o = tk.LabelFrame(j, text="Arquivos Obsoletos")
    fr_o.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    tr_o = ttk.Treeview(fr_o, columns=["Nome do Arquivo","Revisão"], show="headings", height=10)
    tr_o.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    tr_o.heading("Nome do Arquivo", text="Nome do Arquivo")
    tr_o.heading("Revisão", text="Revisão")
    for a in aobs:
        tr_o.insert("", tk.END, values=(a["Nome do Arquivo"], a["Revisão"]))

    def voltar():
        j.destroy()
        master.deiconify()

    def confirmar():
        try:
            caminhos = [Path(a["caminho"]) for a in (arrv + aobs)]
            pasta_raiz_entregas = Path(pasta_entrega)       
            nova = processar_entrega_arquivos_tipo(caminhos, pasta_raiz_entregas, tipo)
            messagebox.showinfo(
                "Sucesso",
                f"Nova entrega criada:\n{nova}\n"
                "_controle_entrega.json gerado com o status dos arquivos.")
            
            j.destroy()
            if master is not None:
                master.destroy()
            sys.exit(0)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar entrega:\n{e}")

    bf = tk.Frame(j)
    bf.pack(side="bottom", anchor="e", pady=5, padx=10)
    ttk.Button(bf, text="Voltar", command=voltar).pack(side=tk.LEFT, padx=5)
    ttk.Button(bf, text="Confirmar", command=confirmar).pack(side=tk.RIGHT, padx=5)
    ttk.Button(j, text="Fechar", command=j.destroy).pack(pady=10)
    j.mainloop()

if __name__ == "__main__":
    root = tk.Tk()           # única raiz real
    root.withdraw()          # se não quiser que apareça imediatamente
    janela_selecao_projeto(root)   # primeira tela
    root.mainloop()
