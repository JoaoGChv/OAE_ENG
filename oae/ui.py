"""Tkinter user interface for delivery management."""
from __future__ import annotations
import os
import re
import sys
import json
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Sequence, Tuple

from .file_ops import (
    PROJETOS_JSON,
    NOMENCLATURAS_JSON,
    GRUPOS_EXT,
    criar_pasta_entrega_ap_pe,
    carregar_nomenclatura_json,
    carregar_ultimo_diretorio,
    salvar_ultimo_diretorio,
    extrair_numero_arquivo,
    split_including_separators,
    verificar_tokens,
    identificar_obsoletos_custom,
    identificar_nome_com_revisao,
    carregar_dados_anteriores,
    obter_info_ultima_entrega,
    listar_arquivos_no_diretorio,
    analisar_comparando_estado,
    pos_processamento,
    _safe_json_load,
)

try:
    from send2trash import send2trash  # type: ignore
except Exception:
    send2trash = None  # pragma: no cover - optional dependency

PASTA_ENTREGA_GLOBAL: str | None = None
NOMENCLATURA_GLOBAL: Dict | None = None
NUM_PROJETO_GLOBAL: str | None = None
TIPO_ENTREGA_GLOBAL: str | None = None

def _center(win: tk.Toplevel | tk.Tk, parent: tk.Toplevel | tk.Tk | None = None) -> None:
    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    if parent:
        parent.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
    else:
        x = (win.winfo_screenwidth() - w) // 2
        y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

def escolher_tipo_entrega(master: tk.Toplevel | tk.Tk, size: tuple[int, int] = (480, 280)) -> str | None:
    escolha = {"val": None}
    selecionado = tk.StringVar(value="AP")
    win = tk.Toplevel(master)
    win.withdraw()
    win.title("Tipo de Entrega")
    win.transient(master)
    win.grab_set()
    win.resizable(False, False)
    win.geometry(f"{size[0]}x{size[1]}")
    ttk.Label(win, text="Escolha o tipo de entrega:", font=("Arial", 11, "bold")).pack(pady=(12, 12))
    area = tk.Frame(win); area.pack(expand=True)
    CARD_W = (size[0] // 2) - 40
    CARD_H = size[1] - 120
    FONT_BIG = ("Arial", 28, "bold")
    def _build_card(parent, texto, valor):
        frame = tk.Frame(parent, width=CARD_W, height=CARD_H, borderwidth=2, relief="ridge", bg="#f0f0f0")
        frame.pack_propagate(False)
        lbl_icon = tk.Label(frame, text=valor, font=FONT_BIG, bg="#f0f0f0")
        lbl_icon.pack(expand=True)
        lbl_txt = tk.Label(frame, text=texto, font=("Arial", 10), bg="#f0f0f0")
        lbl_txt.pack(pady=(0, 6))
        def _on_click(*_):
            selecionado.set(valor)
            _update_highlight()
        frame.bind("<Button-1>", _on_click)
        lbl_icon.bind("<Button-1>", _on_click)
        lbl_txt.bind("<Button-1>", _on_click)
        return frame
    def _update_highlight():
        for card, val in ((card_ap, "AP"), (card_pe, "PE")):
            if selecionado.get() == val:
                card.config(bg="#dbe9ff", highlightbackground="#4e9af1")
            else:
                card.config(bg="#f0f0f0", highlightbackground="#d0d0d0")
    card_ap = _build_card(area, "Anteprojeto – 1.AP", "AP")
    card_pe = _build_card(area, "Projeto Executivo – 2.PE", "PE")
    card_ap.pack(side="left", padx=10, pady=5, expand=True, fill="both")
    card_pe.pack(side="left", padx=10, pady=5, expand=True, fill="both")
    _update_highlight()
    btn_box = ttk.Frame(win); btn_box.pack(pady=(6, 12))
    ttk.Button(btn_box, text="OK", command=lambda: (escolha.update(val=selecionado.get()), win.destroy())).pack(side="left", padx=6)
    ttk.Button(btn_box, text="Cancelar", command=win.destroy).pack(side="left", padx=6)
    _center(win, master)
    win.deiconify()
    master.wait_window(win)
    return escolha["val"]
def _carregar_projetos() -> List[Tuple[str, str]]:
    if not os.path.exists(PROJETOS_JSON):
        messagebox.showerror("Erro", f"Arquivo de projetos não encontrado:\n{PROJETOS_JSON}")
        sys.exit(1)
    with open(PROJETOS_JSON, "r", encoding="utf-8") as f:
        projetos_dict = _safe_json_load(f)
    return list(projetos_dict.items())


def janela_selecao_projeto():
    """
    Abre uma janela com a lista de projetos.

    Colunas exibidas:
        • Número      (ex.: 231)
        • Nome do Projeto  (ex.: OAE-231 - LAEPE)

    O valor interno ‘caminho’ continua sendo o path completo,
    mas na tabela o usuário vê apenas o nome da pasta.
    """
    root = tk.Tk()

    root.geometry("600x400")   # largura x altura em pixels
    root.minsize(900, 500)      # impede que fique muito pequeno

    root.title("Selecionar Projeto")
    root.resizable(True, True)

    if not os.path.exists(PROJETOS_JSON):
        messagebox.showerror("Erro",
                             f"Arquivo de projetos não encontrado:\n{PROJETOS_JSON}")
        sys.exit(1)

    # ---- lê o JSON e cria lista (numero, caminho, nome_exibicao) ----
    with open(PROJETOS_JSON, "r", encoding="utf-8") as f:
        temp = json.load(f).items()
    projetos = [(n, c, os.path.basename(c)) for n, c in temp]

    sel = {"num": None, "path": None}

    # --------------- callbacks internos ---------------
    def confirmar():
        sel_i = tree.selection()
        if not sel_i:
            return
        iid = sel_i[0]
        sel["num"]  = tree.set(iid, "Número")
        índice      = int(tree.item(iid, "text"))          # usamos o índice salvo no item
        sel["path"] = projetos[índice][1]                  # caminho completo
        root.destroy()

    def filtrar(*_):
        termo = entrada.get().lower()
        tree.delete(*tree.get_children())
        for idx, (n, _, nome_disp) in enumerate(projetos):
            if termo in nome_disp.lower():
                tree.insert("", tk.END, text=str(idx),
                            values=(n, nome_disp))

    # --------------- construção da UI ---------------
    tk.Label(root, text="Digite nome ou parte do nome do projeto:"
             ).pack(anchor="w", padx=10, pady=5)

    entrada = tk.Entry(root)
    entrada.pack(fill=tk.X, padx=10)
    entrada.bind("<KeyRelease>", filtrar)
    entrada.bind("<Return>", lambda e: confirmar())

    cols = ("Número", "Nome do Projeto")
    tree = ttk.Treeview(root, columns=cols, show="headings", height=15)
    for c in cols:
        tree.heading(c, text=c)
    tree.column("Número", width=60, anchor="w")
    tree.column("Nome do Projeto", anchor="w")
    tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    for idx, (n, _, nome_disp) in enumerate(projetos):
        tree.insert("", tk.END, text=str(idx), values=(n, nome_disp))

    tk.Button(root, text="Confirmar", command=confirmar).pack(pady=10)
    root.mainloop()
    return sel["num"], sel["path"]


def selecionar_pasta_entrega(diretorio_inicial: str):
    root = tk.Tk(); root.withdraw()
    pasta = filedialog.askdirectory(title="Selecione a pasta de entrega", initialdir=diretorio_inicial)
    root.destroy()
    return pasta


def janela_erro_revisao(arquivos_alterados):
    janela = tk.Toplevel()
    janela.title("Possível Erro de Revisão")
    janela.configure(bg="#FFA07A")
    msg = (
        "Foi identificado que o tamanho dos arquivos abaixo mudou em relação à entrega anterior.\n"
        "Confirma que isso está correto?\n\n"
    )
    for rv, arq, *_ in arquivos_alterados:
        msg += f"{arq} - Revisão: {rv or 'Sem Revisão'}\n"
    tk.Label(janela, text=msg, bg="#FFA07A", font=("Arial", 12)).pack(padx=10, pady=10)

    def _encerra():
        janela.destroy()
        sys.exit(0)

    tk.Button(janela, text="Confirmar e sair", command=_encerra).pack(pady=5)
    tk.Button(janela, text="Ignorar", command=janela.destroy).pack(pady=5)
    janela.grab_set()
    janela.mainloop()


def janela_selecao_disciplina(numero_proj: str, caminho_proj: str) -> tuple[str | None, bool]:
    """Exibe a lista de disciplinas em cartões para seleção."""
    pasta_desenvol = os.path.join(caminho_proj, "3 Desenvolvimento")
    if not os.path.isdir(pasta_desenvol):
        messagebox.showerror(
            "Erro",
            f"A pasta '3 Desenvolvimento' não foi encontrada em:\n{caminho_proj}"
        )
        return None

    ignorar = ("3.0 compatibilização", "3.1 projetos externos")
    disciplinas = []
    for nome in sorted(os.listdir(pasta_desenvol)):
        if any(nome.strip().lower().startswith(term) for term in ignorar):
            continue
        p = os.path.join(pasta_desenvol, nome)
        if os.path.isdir(p):
            disciplinas.append(p)

    if not disciplinas:
        messagebox.showerror("Erro", "Nenhuma disciplina encontrada.")
        return None

    root = tk.Tk()
    root.title(f"Projeto {numero_proj} – Selecionar Disciplina")
    root.geometry("700x500")
    root.minsize(700, 400)

    tk.Label(root, text="Filtrar disciplina:").pack(anchor="w", padx=10, pady=5)
    var_filtro = tk.StringVar()
    entrada = tk.Entry(root, textvariable=var_filtro)
    entrada.pack(fill=tk.X, padx=10)

    main_content_frame = tk.Frame(root)
    main_content_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    canvas = tk.Canvas(main_content_frame, borderwidth=0, highlightthickness=0)
    vsb = tk.Scrollbar(main_content_frame, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)
    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    frame = tk.Frame(canvas, bg="#f4f4f4")
    canvas.create_window((0, 0), window=frame, anchor="nw")

    def _on_mousewheel(event):
        if event.num == 4 or event.delta > 0:
            canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            canvas.yview_scroll(1, "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    canvas.bind_all("<Button-4>", _on_mousewheel)
    canvas.bind_all("<Button-5>", _on_mousewheel)

    def _unbind_wheel():
        canvas.unbind_all("<MouseWheel>")
        canvas.unbind_all("<Button-4>")
        canvas.unbind_all("<Button-5>")

    CARD_W, CARD_H = 120, 120
    sel = {"widget": None, "path": None}
    voltar_flag = {"val": False}


    def _render():
        for w in frame.winfo_children():
            w.destroy()

        termo = var_filtro.get().lower()
        paths = [p for p in disciplinas if termo in os.path.basename(p).lower()]
        cols = max(1, (canvas.winfo_width() // (CARD_W + 20)))

        for idx, path in enumerate(paths):
            r, c = divmod(idx, cols)
            card = tk.Frame(
                frame,
                width=CARD_W,
                height=CARD_H,
                bg="white",
                highlightthickness=1,
                highlightbackground="#c0c0c0",
                relief="flat",
            )
            card.grid(row=r, column=c, padx=10, pady=10)
            card.grid_propagate(False)

            lbl = tk.Label(
                card,
                text=os.path.basename(path),
                wraplength=CARD_W - 10,
                justify="center",
                font=("TkDefaultFont", 10),
            )
            lbl.place(relx=0.5, rely=0.5, anchor="center")

            def _select(widget=card, caminho=path):
                if sel["widget"]:
                    sel["widget"].config(bg="white", highlightbackground="#c0c0c0")
                widget.config(bg="#dbe9ff", highlightbackground="#4e9af1")
                sel["widget"], sel["path"] = widget, caminho

            card.bind("<Button-1>", lambda e, w=card, p=path: _select(w, p))
            card.bind("<Double-Button-1>", lambda e, p=path: _confirmar(p))
            lbl.bind("<Button-1>", lambda e, w=card, p=path: _select(w, p))
            lbl.bind("<Double-Button-1>", lambda e, p=path: _confirmar(p))

    def _confirmar(caminho_sel):
        if not caminho_sel:
            return
        pasta_entregas = None
        for folder in os.listdir(caminho_sel):
            if folder.strip().lower().replace(" ", "") == "1.entregas":
                pasta_entregas = os.path.join(caminho_sel, folder)
                break
        if not pasta_entregas or not os.path.isdir(pasta_entregas):
            messagebox.showerror(
                "Erro",
                "A subpasta '1.ENTREGAS' não foi encontrada dentro da disciplina."
            )
            return
        nonlocal_result[0] = pasta_entregas
        _unbind_wheel()
        root.destroy()

    footer = tk.Frame(root)
    footer.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
    button_frame = tk.Frame(footer)
    button_frame.pack(expand=True)

    btn_voltar = tk.Button(
        button_frame,
        text="Voltar",
        width=15,
        command=lambda: (voltar_flag.update(val=True), root.destroy()),
    )
    btn_confirmar = tk.Button(
        button_frame,
        text="Confirmar",
        width=15,
        command=lambda: _confirmar(sel["path"]),
    )
    btn_voltar.pack(side=tk.LEFT, padx=10)
    btn_confirmar.pack(side=tk.LEFT, padx=10)

    var_filtro.trace_add("write", lambda *_: _render())
    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind(
        "<Configure>",
        lambda e: (_render(), canvas.itemconfig("all", width=canvas.winfo_width())),
    )

    nonlocal_result = [None]
    _render()
    root.mainloop()
    return nonlocal_result[0], voltar_flag["val"]

class TelaVisualizacaoEntregaAnterior(tk.Tk):
    """Mostra arquivos da entrega AP/PE mais recente e permite renomear/excluir."""

    def __init__(self, pasta_entregas: str, projeto_num: str, disciplina: str, lista_inicial=None, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Entrega Anterior – Visualização")
        self.geometry("1000x600")
        self.resizable(True, True)
        self.configure(bg="#f5f5f5")

        self.pasta_entregas = pasta_entregas  # caminho para 1.ENTREGAS
        self.projeto_num = projeto_num
        self.disciplina = disciplina
        self.lista_inicial = lista_inicial
        self.reabrir_disciplina = False

        # ------ estado ------
        self.tipo_var = tk.StringVar(value="AP")  # AP ou PE selecionado para visualização
        self.lista_arquivos: list[tuple[str, str, int, str, str]] = []  # (rev, nome, tam, cam, dt)

        # UI ------------------------------------------------------------------
        header = tk.Frame(self, bg="#2c3e50")
        header.pack(fill=tk.X)
        tk.Label(header, text=f"Projeto {projeto_num}  •  {disciplina}", fg="white", bg="#2c3e50",
                 font=("Helvetica", 14, "bold")).pack(padx=10, pady=6, anchor="w")

        ctrl = tk.Frame(self)
        ctrl.pack(fill=tk.X, padx=10, pady=(10, 5))
        tk.Label(ctrl, text="Visualizar entregas de:").pack(side=tk.LEFT)
        cmb_tipo = ttk.Combobox(ctrl, values=["AP", "PE"], textvariable=self.tipo_var,
                               width=4, state="readonly")
        cmb_tipo.pack(side=tk.LEFT, padx=5)
        cmb_tipo.bind("<<ComboboxSelected>>", lambda e: self._carregar_entrega())

        ttk.Button(ctrl, text="Excluir selecionados", command=self._excluir_selecionados).pack(side=tk.RIGHT, padx=5)
        ttk.Button(ctrl, text="Avançar", command=self._avancar).pack(side=tk.RIGHT, padx=5)
        ttk.Button(ctrl, text="Voltar", command=self._voltar).pack(side=tk.RIGHT, padx=5)

        # tabela --------------------------------------------------------------
        tbl_frame = tk.Frame(self)
        tbl_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.sb_y = tk.Scrollbar(tbl_frame, orient="vertical")
        self.sb_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree = ttk.Treeview(
            tbl_frame,
            columns=("sel", "nome", "dt", "tam"),
            show="headings",
            yscrollcommand=self.sb_y.set,
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sb_y.config(command=self.tree.yview)
        self.tree.heading("sel", text="Sel")
        self.tree.heading("nome", text="Nome do arquivo", anchor="w")
        self.tree.heading("dt", text="Data mod.")
        self.tree.heading("tam", text="Tamanho (KB)")
        self.tree.column("sel", width=40, anchor="center")
        self.tree.column("nome", width=400, anchor="w")
        self.tree.column("dt", width=130, anchor="center")
        self.tree.column("tam", width=100, anchor="e")

        self.tree.bind("<Double-1>", self._iniciar_edicao_nome)
        self.tree.bind("<Button-1>", self._toggle_checkbox)

        self.checked = {}

        if self.lista_inicial:
            self._carregar_lista_inicial()
        else:
            self._carregar_entrega()

    # ------------------------------------------------------------------
    # lógica de descoberta de entrega mais recente
    # ------------------------------------------------------------------
    def _folder_mais_recente(self, tipo: str) -> str | None:
        pasta_tipo = os.path.join(self.pasta_entregas, tipo)
        if not os.path.isdir(pasta_tipo):
            return None
        candidatas = []
        for d in os.listdir(pasta_tipo):
            if d.endswith("-OBSOLETO"):
                continue
            m = re.search(r"Entrega-(\d+)$", d)
            if m:
                candidatas.append((int(m.group(1)), d))
        if not candidatas:
            return None
        _, pasta_nome = max(candidatas, key=lambda t: t[0])
        return os.path.join(pasta_tipo, pasta_nome)

    def _carregar_entrega(self):
        if "cam_full" not in self.tree["columns"]:
            self.tree["columns"] = self.tree["columns"] + ("cam_full",)
            self.tree.column("cam_full", width=0, stretch=False)
        self.tree.delete(*self.tree.get_children())
        self.lista_arquivos.clear()
        self.checked.clear()

        todos_root = listar_arquivos_no_diretorio(self.pasta_entregas)
        tipo = self.tipo_var.get()

        todos = []
        for rv, nome, tam, cam, dt_mod in todos_root:
            caminho_norm = cam.replace("\\", "/").lower()
            if f"/{tipo.lower()}/" in caminho_norm:
                todos.append((rv, nome, tam, cam, dt_mod))
            else:
                pai = os.path.basename(os.path.dirname(cam)).lower()
                if tipo == "AP" and pai.startswith("1.ap"):
                    todos.append((rv, nome, tam, cam, dt_mod))
                elif tipo == "PE" and pai.startswith("2.pe"):
                    todos.append((rv, nome, tam, cam, dt_mod))

        if not todos:
            messagebox.showinfo(
                "Info",
                "Nenhum arquivo válido foi encontrado nas pastas de entrega."
            )
            return

        obsoletos = set(identificar_obsoletos_custom(todos))
        validos = [t for t in todos if t not in obsoletos]

        for rv, nome, tam, cam, dt_mod in sorted(validos, key=lambda x: x[1].lower()):
            tam_kb = tam // 1024
            self.lista_arquivos.append((rv, nome, tam, cam, dt_mod))
            iid = self.tree.insert("", tk.END, values=("\u2610", nome, dt_mod, tam_kb))
            self.checked[iid] = False
            # marca para uso posterior (cam path)
            self.tree.set(iid, "cam_full", cam)

    def _carregar_lista_inicial(self):
        """Preenche a Treeview usando self.lista_inicial recebida do botão Voltar."""
        if not self.lista_inicial:
            return
        # garante coluna oculta
        if "cam_full" not in self.tree["columns"]:
            self.tree["columns"] = self.tree["columns"] + ("cam_full",)
            self.tree.column("cam_full", width=0, stretch=False)

        self.tree.delete(*self.tree.get_children())
        self.checked.clear()
        for rv, nome, tam, cam, dt in self.lista_inicial:
            iid = self.tree.insert("", tk.END,
                                values=("\u2610", nome, dt, tam // 1024))
            self.tree.set(iid, "cam_full", cam)
            self.checked[iid] = False
        # mantém lista interna sincronizada
        self.lista_arquivos = self.lista_inicial.copy()

    def _toggle_checkbox(self, event):
        iid = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if col != "#1" or not iid:
            return
        new_state = not self.checked.get(iid, False)
        self.checked[iid] = new_state
        vals = list(self.tree.item(iid, "values"))
        vals[0] = "\u2611" if new_state else "\u2610"
        self.tree.item(iid, values=vals)
        return "break"


    # util simples para pegar revisao do nome (usa regex já presente no código original)
    @staticmethod
    def _identificar_rev(nome):
        nb, rev, ex = identificar_nome_com_revisao(nome)
        return rev, ex

    # --------------------------- edição ------------------------------
    def _iniciar_edicao_nome(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        col = self.tree.identify_column(event.x)
        if col != "#2":  # só permite editar coluna 'nome'
            return
        x, y, w, h = self.tree.bbox(iid, col)
        valor_antigo = self.tree.set(iid, "nome")

        entry = tk.Entry(self.tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, valor_antigo)
        entry.focus_set()

        def _salvar(e):
            novo_nome = entry.get().strip()
            entry.destroy()
            if not novo_nome or novo_nome == valor_antigo:
                return
            # renomeia no sistema de arquivos
            idx = self.tree.index(iid)
            cam_antigo = self.lista_arquivos[idx][3]
            novo_cam = os.path.join(os.path.dirname(cam_antigo), novo_nome)
            try:
                os.rename(cam_antigo, novo_cam)
            except OSError as err:
                messagebox.showerror("Erro", f"Falha ao renomear arquivo:\n{err}")
                return
            # atualiza estruturas
            tam, dtmod = self.lista_arquivos[idx][2], self.lista_arquivos[idx][4]
            self.lista_arquivos[idx] = (self.lista_arquivos[idx][0], novo_nome, tam, novo_cam, dtmod)
            self.tree.set(iid, "nome", novo_nome)
        entry.bind("<Return>", _salvar)
        entry.bind("<FocusOut>", _salvar)

    # --------------------------- exclusão ----------------------------
    def _excluir_selecionados(self):
        marked = [iid for iid, val in self.checked.items() if val]
        if not marked:
            return
        if not messagebox.askyesno(
            "Confirmação",
            "Tem certeza de que deseja excluir os arquivos selecionados?\nEles serão enviados à lixeira.",
        ):
            return
        erros = []
        for iid in sorted(marked, key=self.tree.index, reverse=True):
            idx = self.tree.index(iid)
            _, nome, _, cam_full, _ = self.lista_arquivos[idx]
            cam_full = os.path.normpath(cam_full)
            try:
                if send2trash:
                    try:
                        send2trash(cam_full)
                    except OSError:
                        os.remove(cam_full)
                else:
                    os.remove(cam_full)
            except OSError as err:
                erros.append(str(err))
                continue
            self.tree.delete(iid)
            self.lista_arquivos.pop(idx)
            self.checked.pop(iid, None)
        if erros:
            messagebox.showwarning(
                "Aviso",
                "Alguns arquivos não puderam ser excluídos:\n" + "\n".join(erros),
            )

    # --------------------------- navegação ---------------------------
    def _avancar(self):
        lista_init = []
        for iid in self.tree.get_children():
            if self.checked.get(iid):
                idx = self.tree.index(iid)
                rv, nome, tam, cam, dt = self.lista_arquivos[idx]
                lista_init.append((rv, nome, tam, cam, dt))
        self.destroy()
        TelaAdicaoArquivos(lista_inicial=lista_init, pasta_entrega=self.pasta_entregas,
                           numero_projeto=self.projeto_num).mainloop()

    def _voltar(self):
        # caminho do projeto = três níveis acima de 1.ENTREGAS
        caminho_proj = os.path.dirname(
            os.path.dirname(          #  Disciplina
                os.path.dirname(      #  3 Desenvolvimento
                    self.pasta_entregas)))  # 1.ENTREGAS

        # esconde a janela atual temporariamente
        self.withdraw()
        nova_pasta = janela_selecao_disciplina(self.projeto_num, caminho_proj)

        if not nova_pasta:  # usuário cancelou
            # reexibe a janela atual
            self.deiconify()
            return

        # seleção válida: fecha a atual e abre novamente com a nova pasta
        self.destroy()
        TelaVisualizacaoEntregaAnterior(
            pasta_entregas=nova_pasta,
            projeto_num=self.projeto_num,
            disciplina=os.path.basename(os.path.dirname(nova_pasta))
        ).mainloop()

# -----------------------------------------------------
# Janela 1: TelaAdicaoArquivos
# -----------------------------------------------------
class TelaAdicaoArquivos(tk.Tk):    
    """
    Janela inicial: o usuário seleciona os arquivos para entrega.
    """
    def __init__(self, lista_inicial=None, pasta_entrega: str | None = None, numero_projeto: str | None = None, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.pasta_entrega   = pasta_entrega
        self.numero_projeto  = numero_projeto
        self.disciplina      = (os.path.basename(os.path.dirname(pasta_entrega))if pasta_entrega else "")
        root_frame = tk.Frame(self)          # capa principal
        root_frame.pack(fill=tk.BOTH, expand=True)

        bar_l = tk.Frame(root_frame, bg="#2c3e50", width=200)
        bar_l.pack(side=tk.LEFT, fill=tk.Y)
        bar_l.pack_propagate(False)

        tk.Label(bar_l, text="OAE - Engenharia",
                font=("Helvetica", 14, "bold"),
                bg="#2c3e50", fg="white").pack(pady=10)

        tk.Label(bar_l, text="PROJETOS",
                font=("Helvetica", 10, "bold"),
                bg="#34495e", fg="white", anchor="w", padx=10
                ).pack(fill=tk.X, pady=(0, 5))

        lst_proj = tk.Listbox(bar_l, height=5,
                            bg="#ecf0f1", font=("Helvetica", 9))
        lst_proj.pack(fill=tk.X, padx=10, pady=5)
        if numero_projeto:
            lst_proj.insert(tk.END, f"Projeto {numero_projeto}")

        tk.Label(bar_l, text="MEMBROS",
                font=("Helvetica", 10, "bold"),
                bg="#34495e", fg="white", anchor="w", padx=10
                ).pack(fill=tk.X, pady=5)

        # Painel onde ficará TODO o conteúdo antigo
        content = tk.Frame(root_frame, bg="#f5f5f5")
        content.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        lbl_title = tk.Label(content, text="Adicionar Arquivos para Entrega",
                        font=("Helvetica", 15, "bold"), bg="#f5f5f5",
                        anchor="w")
        lbl_title.pack(fill=tk.X, pady=(10, 5), padx=10)

        self.resizable(True, True)
        self.title("Adicionar Arquivos para Entrega")

        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        new_height = int(screen_h * 0.85)
        new_width = 1100
        self.geometry(f"{new_width}x{new_height}")

        self.view_mode = "grouped"

        container = tk.Frame(content)
        container.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        lbl = tk.Label(container, text="Selecione arquivos para entrega. A listagem aparecerá conforme forem adicionados.")
        lbl.pack(pady=5)

        btn_frame = tk.Frame(container)
        btn_frame.pack(fill=tk.X, pady=5)

        tk.Button(btn_frame, text="Adicionar arquivos", command=self.adicionar_arquivos).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Remover arquivos selecionados", command=self.remover_selecionados).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="Visualizar tudo", command=self.ativar_visualizar_tudo).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Agrupar por tipo", command=self.ativar_agrupado).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="Analisar", width=18, command=self.proxima_janela_nomenclatura).pack(side=tk.RIGHT, padx=5)

        self.canvas_global = tk.Canvas(content)
        self.canvas_global.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar_global = tk.Scrollbar(container, orient="vertical", command=self.canvas_global.yview, width=20)
        self.scrollbar_global.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas_global.configure(yscrollcommand=self.scrollbar_global.set)

        self.inner_frame = tk.Frame(self.canvas_global)
        self.inner_frame_id = self.canvas_global.create_window((0,0), window=self.inner_frame, anchor="nw")

        self.canvas_global.bind("<Configure>", self.on_canvas_configure)
        self.inner_frame.bind("<Configure>", self.on_frame_configure)

        bottom = tk.Frame(content)
        # ancorado no rodapé:
        bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 10))

        tk.Button(bottom, text="Cancelar",
                  command=self._cancelar).pack(side=tk.LEFT, padx=5)

        tk.Button(bottom, text="Voltar",
                  command=self._voltar).pack(side=tk.RIGHT, padx=5)

        self.arquivos_por_grupo = {}

        self.paned = None
        self.table_all = None

        self.colunas = ("num_arq", "nome_arq", "revisao", "dt_mod", "ext")
        self.coluna_titulos = {
            "num_arq": "Número do arquivo",
            "nome_arq": "Nome do arquivo",
            "revisao": "Revisão",
            "dt_mod": "Data de modificação",
            "ext": "Extensão",
        }
        self.tables = {}

        if lista_inicial:
            self.recarregar_lista(lista_inicial)

        self.render_view()

        if pasta_entrega and (not lista_inicial):
            self._abrir_filedialog_inicial(pasta_entrega)

    def _abrir_filedialog_inicial(self, pasta_entrega):
        init_dir = carregar_ultimo_diretorio()
        if not init_dir or not os.path.exists(init_dir):
            init_dir = pasta_entrega if os.path.isdir(pasta_entrega) else os.path.expanduser("~")
        paths = filedialog.askopenfilenames(
            title="Selecione arquivos para entrega",
            initialdir=init_dir
        )
        if not paths:
            return
        dir_of_first = os.path.dirname(paths[0])
        salvar_ultimo_diretorio(dir_of_first)

        for p in paths:
            if not os.path.isfile(p):
                continue
            base, rev, ext = identificar_nome_com_revisao(os.path.basename(p))
            data_ts = os.path.getmtime(p)
            data_mod = datetime.datetime.fromtimestamp(data_ts).strftime("%d/%m/%Y %H:%M")
            no_arq = extrair_numero_arquivo(base)
            grupo = self.classificar_extensao(ext)
            self.arquivos_por_grupo.setdefault(grupo, [])
            if p not in [x[0] for x in self.arquivos_por_grupo[grupo]]:
                self.arquivos_por_grupo[grupo].append((p, no_arq, base, rev, data_mod, ext))

        self.render_view()

    def recarregar_lista(self, lista_inicial):
        for (rv, arq, tam, path, dmod) in lista_inicial:
            base, revisao, ext = identificar_nome_com_revisao(arq)
            no_arq = extrair_numero_arquivo(base)
            grupo = self.classificar_extensao(ext)
            if grupo not in self.arquivos_por_grupo:
                self.arquivos_por_grupo[grupo] = []
            self.arquivos_por_grupo[grupo].append((path, no_arq, base, revisao, dmod, ext))
        self.render_view()

    def on_canvas_configure(self, event):
        self.canvas_global.itemconfig(self.inner_frame_id, width=event.width)

    def on_frame_configure(self, event):
        self.canvas_global.configure(scrollregion=self.canvas_global.bbox("all"))

    def _cancelar(self):
        self.destroy()
        sys.exit(0)

    def _voltar(self):
        lista_atual = []
        for grupo, lst in self.arquivos_por_grupo.items():
            for (path, no_arq, base, rev, data_mod, ext) in lst:
                tam = os.path.getsize(path)
                arq = os.path.basename(path)
                lista_atual.append((rev, arq, tam, path, data_mod))

        self.destroy()

        TelaVisualizacaoEntregaAnterior(
            pasta_entregas=self.pasta_entrega,
            projeto_num=self.numero_projeto,
            disciplina=self.disciplina,
            lista_inicial=lista_atual
        ).mainloop()

    def adicionar_arquivos(self):
        init_dir = carregar_ultimo_diretorio()
        if not init_dir or not os.path.exists(init_dir):
            init_dir = os.path.expanduser("~")

        paths = filedialog.askopenfilenames(title="Selecione arquivos para entrega", initialdir=init_dir)
        if not paths:
            return
        dir_of_first = os.path.dirname(paths[0])
        salvar_ultimo_diretorio(dir_of_first)

        for p in paths:
            if not os.path.isfile(p):
                continue
            base, rev, ext = identificar_nome_com_revisao(os.path.basename(p))
            data_ts = os.path.getmtime(p)
            data_mod = datetime.datetime.fromtimestamp(data_ts).strftime("%d/%m/%Y %H:%M")
            no_arq = extrair_numero_arquivo(base)
            grupo = self.classificar_extensao(ext)

            if grupo not in self.arquivos_por_grupo:
                self.arquivos_por_grupo[grupo] = []
            if p not in [x[0] for x in self.arquivos_por_grupo[grupo]]:  
                self.arquivos_por_grupo[grupo].append((p, no_arq, base, rev, data_mod, ext))

        self.render_view()

    def classificar_extensao(self, ext):
        if ext == ".pdf":
            return "PDF"
        for grupo, exts in GRUPOS_EXT.items():
            if ext in exts:
                return grupo
        return "Outros"

    def ativar_visualizar_tudo(self):
        self.view_mode = "all"
        self.render_view()

    def ativar_agrupado(self):
        self.view_mode = "grouped"
        self.render_view()

    def render_view(self):
        for child in self.inner_frame.winfo_children():
            child.destroy()
        self.tables.clear()
        self.table_all = None
        self.paned = None

        if self.view_mode == "grouped":
            self.render_tables_grouped()
        else:
            self.render_table_all()

        self.inner_frame.update_idletasks()
        self.canvas_global.config(scrollregion=self.canvas_global.bbox("all"))

    def render_tables_grouped(self):
        """
        Cria uma Treeview por grupo de extensão.
        • Se houver só 1 grupo → coloca o quadro direto no inner_frame
          (evita colapso do Panedwindow).
        • Se houver 2+ grupos → mantém Panedwindow como antes.
        """
        grupos = [(g, lst) for g, lst in self.arquivos_por_grupo.items() if lst]

        # ---------- CASO 1 GRUPO ----------------------------------
        if len(grupos) == 1:
            grupo, lista = grupos[0]
            self._criar_quadro_grupo(self.inner_frame, grupo, lista)
            return  # pronto

        # ---------- CASO 2+ GRUPOS (comportamento original) -------
        self.paned = ttk.Panedwindow(self.inner_frame, orient="vertical")
        self.paned.pack(fill=tk.BOTH, expand=True)

        ttk.Style().configure("Treeview", rowheight=24)

        for grupo, lista in grupos:
            quadro = tk.Frame(self.paned)
            self._criar_quadro_grupo(quadro, grupo, lista)
            # minsize evita altura 0 mesmo se lista esvaziar depois
            self.paned.add(quadro, weight=1, minsize=120)

    # --------------------------------------------------------------
    # Função auxiliar reaproveitada pelos dois caminhos
    # --------------------------------------------------------------
    def _criar_quadro_grupo(self, parent, grupo, lista_arqs):
        """Monta label + Treeview para um grupo de arquivos."""
        tk.Label(parent, text=grupo, font=("Arial", 11, "bold")
                 ).pack(anchor="w", pady=5)

        sub = tk.Frame(parent); sub.pack(fill=tk.BOTH, expand=True)
        sub.config(height=120); sub.pack_propagate(False)

        sb = tk.Scrollbar(sub, orient="vertical", width=20)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        tree = ttk.Treeview(
            sub, columns=self.colunas, show="headings",
            yscrollcommand=sb.set
        )
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.config(command=tree.yview)

        for col in self.colunas:
            tree.heading(col, text=self.coluna_titulos[col],
                         command=lambda c=col, t=tree:
                             self.sort_column(t, c, False))
            tree.column(col, width=120, anchor='w', stretch=True)
        tree.column("nome_arq", width=220)

        for full, num, base, rev, dmod, ext in lista_arqs:
            tree.insert("", tk.END, values=(num, base, rev, dmod, ext))

        # Guardamos referência para futura remoção/ordenamento
        self.tables[grupo] = tree

    def render_table_all(self):
        all_files = []
        for grupo, lista_arqs in self.arquivos_por_grupo.items():
            all_files.extend(lista_arqs)

        frame_all = tk.Frame(self.inner_frame)
        frame_all.pack(fill=tk.BOTH, expand=True)

        scrollbar_all = tk.Scrollbar(frame_all, orient="vertical", width=20)
        scrollbar_all.pack(side=tk.RIGHT, fill=tk.Y)

        self.table_all = ttk.Treeview(
            frame_all,
            columns=self.colunas,
            show="headings",
            yscrollcommand=scrollbar_all.set
        )
        self.table_all.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_all.config(command=self.table_all.yview)

        style = ttk.Style()
        style.configure("Treeview", rowheight=24)

        for col in self.colunas:
            self.table_all.heading(
                col,
                text=self.coluna_titulos[col],
                command=lambda c=col: self.sort_column(self.table_all, c, False)
            )
            self.table_all.column(col, width=120, anchor='w', stretch=True)

        self.table_all.column("nome_arq", width=220)

        for (fullpath, no_arq, base, rev, data_mod, ext) in all_files:
            self.table_all.insert("", tk.END, values=(no_arq, base, rev, data_mod, ext))

    def sort_column(self, tree, col, reverse):
        def try_convert(v):
            try:
                return (0, float(v))
            except:
                return (1, str(v))
        l = []
        for k in tree.get_children(""):
            val = tree.set(k, col)
            l.append((try_convert(val), k))
        l.sort(key=lambda x: x[0], reverse=reverse)

        for index, (_, iid) in enumerate(l):
            tree.move(iid, '', index)
        tree.heading(col, command=lambda: self.sort_column(tree, col, not reverse))

    def remover_selecionados(self):
        if self.view_mode == "grouped":
            for grupo, tree in self.tables.items():
                sel = tree.selection()
                if not sel:
                    continue
                for iid in reversed(sel):
                    vals = tree.item(iid,"values")
                    no_arq, base, rev, data_mod, ext = vals
                    idx_rm = None
                    for i, tup in enumerate(self.arquivos_por_grupo[grupo]):
                        if (
                            tup[1] == no_arq and
                            tup[2] == base and
                            tup[3] == rev and
                            tup[4] == data_mod and
                            tup[5] == ext
                        ):
                            idx_rm = i
                            break
                    if idx_rm is not None:
                        self.arquivos_por_grupo[grupo].pop(idx_rm)
                    tree.delete(iid)
        else:
            if self.table_all:
                sel = self.table_all.selection()
                if not sel:
                    return
                for iid in reversed(sel):
                    vals = self.table_all.item(iid,"values")
                    no_arq, base, rev, data_mod, ext = vals
                    self.table_all.delete(iid)
                    for gkey in list(self.arquivos_por_grupo.keys()):
                        lst = self.arquivos_por_grupo[gkey]
                        idx_rm = None
                        for i, tup in enumerate(lst):
                            if (
                                tup[1] == no_arq and
                                tup[2] == base and
                                tup[3] == rev and
                                tup[4] == data_mod and
                                tup[5] == ext
                            ):
                                idx_rm = i
                                break
                        if idx_rm is not None:
                            self.arquivos_por_grupo[gkey].pop(idx_rm)
                            break

    def proxima_janela_nomenclatura(self):
        final_list = []
        for grupo, lista_arqs in self.arquivos_por_grupo.items():
            for (path, no_arq, base, rev, data_mod, ext) in lista_arqs:
                tam = os.path.getsize(path)
                arq = os.path.basename(path)
                final_list.append((rev, arq, tam, path, data_mod))
        self.destroy()
        tela = TelaVerificacaoNomenclatura(final_list)
        tela.mainloop()


# -----------------------------------------------------
# Janela 2: TelaVerificacaoNomenclatura
# -----------------------------------------------------
class TelaVerificacaoNomenclatura(tk.Tk):
    def __init__(self, lista_arquivos, *args, **kwargs):
        super().__init__(*args, **kwargs)
        global NOMENCLATURA_GLOBAL, NUM_PROJETO_GLOBAL
        if NUM_PROJETO_GLOBAL:
            NOMENCLATURA_GLOBAL = carregar_nomenclatura_json(NUM_PROJETO_GLOBAL)
        self.lista_arquivos = lista_arquivos.copy()  # ← evitar referência residual

        self._token_map = {}
        self._tags_map  = {}
        for rv, arq, tam, path, dmod in self.lista_arquivos:
            nome_sem_ext, _ = os.path.splitext(arq)
            tokens = split_including_separators(nome_sem_ext, NOMENCLATURA_GLOBAL)
            tags   = verificar_tokens(tokens, NOMENCLATURA_GLOBAL)
            self._token_map[arq] = tokens
            self._tags_map[arq]  = tags

        self.title("Verificação de Nomenclatura")
        self.resizable(True, True)
        self.lista_arquivos = lista_arquivos
        self.geometry("1200x700")

        container = tk.Frame(self)
        container.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        lbl = tk.Label(container, text="Confira a nomenclatura (campos e separadores). Caso haja erros, corrija antes de avançar.")
        lbl.pack(anchor="w", pady=5)

        frm_botoes = tk.Frame(container)
        frm_botoes.pack(fill=tk.X, pady=5)

        tk.Button(frm_botoes, text="Mostrar Padrão", command=self.mostrar_nomenclatura_padrao).pack(side=tk.LEFT, padx=5)
        tk.Button(frm_botoes, text="Voltar", command=self.voltar).pack(side=tk.LEFT, padx=5)
        tk.Button(frm_botoes, text="Avançar", command=self.avancar).pack(side=tk.RIGHT, padx=5)

        # Cria Treeview
        frame_tv = tk.Frame(container)
        frame_tv.pack(fill=tk.BOTH, expand=True)

        self.tv_scroll_y = tk.Scrollbar(frame_tv, orient="vertical")
        self.tv_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(frame_tv, show="headings", yscrollcommand=self.tv_scroll_y.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tv_scroll_y.config(command=self.tree.yview)

        # Definimos a quantidade de colunas = "máximo de tokens" que possamos ter
        # em qualquer arquivo. A cada arquivo, iremos tokenizar. Se um arquivo tiver 20 tokens,
        # e outro 24, tomamos o max.
        self.max_tokens = 0
        self.tabela_por_arquivo = []  # guardamos tokens de cada arquivo

        # Calcula o max_tokens
        for (rv, arq, tam, path, dmod) in self.lista_arquivos:
            nome_sem_ext, _ = os.path.splitext(arq)
            tokens = split_including_separators(nome_sem_ext, NOMENCLATURA_GLOBAL)
            if len(tokens) > self.max_tokens:
                self.max_tokens = len(tokens)

        # Criamos colunas: col0 => "Arquivo", col1.. => tokens
        # Mas removemos a "Arquivo"? O enunciado pediu sem a primeira col "Nome do arquivo".
        # Entretanto, precisamos exibir de algum modo. Então, assumo que "Arquivo" pode ser col0
        # Se realmente quiser remover, basta não ter a col do nome do arquivo. 
        # Mas a solicitação diz: "Não há necessidade da primeira coluna 'Nome do arquivo'." 
        # Então, iremos exibir apenas as colunas de tokens.
        
        col_names = [f"T{i}" for i in range(1, self.max_tokens+1)]
        self.tree["columns"] = col_names
        for cn in col_names:
            self.tree.heading(cn, text="")
            self.tree.column(cn, width=60, anchor="center")

        # Configuramos tags
        self.tree.tag_configure("mismatch", background="#FF9999")  # vermelho clarinho
        self.tree.tag_configure("missing", background="#FFFF99")   # amarelo clarinho
        self.tree.tag_configure("ok", background="white")

        self.preencher_arvore()

    def preencher_arvore(self):
        self.tree.delete(*self.tree.get_children())

        for (rv, arq, tam, path, dmod) in self.lista_arquivos:
            nome_sem_ext, _ = os.path.splitext(arq)
            tokens = split_including_separators(nome_sem_ext, NOMENCLATURA_GLOBAL)
            tags_result = verificar_tokens(tokens, NOMENCLATURA_GLOBAL)
            # usa sempre o que foi pré-computado
            tokens      = self._token_map[arq]
            tags_result = self._tags_map[arq]
            
            # Montamos a row com length = self.max_tokens
            row_vals = []
            row_tags = []
            for i in range(self.max_tokens):
                if i < len(tokens):
                    row_vals.append(tokens[i])
                else:
                    row_vals.append("")  # faltando
                
                # Tag da célula
                if i < len(tags_result):
                    if tags_result[i] == 'mismatch':
                        row_tags.append('mismatch')
                    elif tags_result[i] == 'missing':
                        row_tags.append('missing')
                    else:
                        row_tags.append('ok')
                else:
                    # se sobrou
                    row_tags.append('missing')

            # insere item
            item_id = self.tree.insert("", tk.END, values=row_vals)
            # para cada coluna, se row_tags[i] != 'ok', definimos a tag naquela célula
            for i, tg in enumerate(row_tags):
                if tg in ['mismatch','missing']:
                    self.tree.set(item_id, self.tree["columns"][i], row_vals[i])
                    self.tree.item(item_id, tags=(tg,))  # Marca a linha com a tag
                # Observação: com a Treeview default, a tag é para a linha inteira,
                # mas podemos simular "célula" colorida com 'tag_cell = (item_id, column)' 
                # e um workaround. Para simplificar, a doc. oficial do tkinter não
                # coloriza células individualmente sem bibliotecas extras.
                # Então a approach do treeview do tkinter colore a "linha" 
                # se encontrar mismatch em qualquer token.
                if tg != 'ok':
                    # se preferir colorir a linha toda se tiver 1 mismatch
                    # mas o enunciado diz "não toda a linha, apenas a célula".
                    # Precisaríamos de um workaround custom. 
                    pass

    def mostrar_nomenclatura_padrao(self):
        """Exibe uma janela com a nomenclatura padrão, cada campo + separador."""
        if not NOMENCLATURA_GLOBAL:
            messagebox.showinfo("Info", "Nomenclatura não definida para este projeto.")
            return

        win = tk.Toplevel(self)
        win.title("Nomenclatura Padrão")
        win.geometry("1000x300")

        campos = NOMENCLATURA_GLOBAL.get("campos", [])

        # Montar colunas => "campo, sep, campo, sep, campo..."
        # Exemplo: 4 campos => 7 colunas (campo,sep,campo,sep,campo,sep,campo)
        # mas o último separador é opcional, dependendo do JSON. Ajuste a gosto.
        col_count = 2*len(campos) - 1
        col_ids = [f"C{i}" for i in range(col_count)]
        
        frm_tree = tk.Frame(win)
        frm_tree.pack(fill=tk.BOTH, expand=True)

        sbx = tk.Scrollbar(frm_tree, orient="horizontal")
        sby = tk.Scrollbar(frm_tree, orient="vertical")
        sby.pack(side=tk.RIGHT, fill=tk.Y)
        sbx.pack(side=tk.BOTTOM, fill=tk.X)

        tv = ttk.Treeview(frm_tree, columns=col_ids, show="headings",
                          xscrollcommand=sbx.set, yscrollcommand=sby.set)
        tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sbx.config(command=tv.xview)
        sby.config(command=tv.yview)

        for cid in col_ids:
            tv.heading(cid, text="")
            tv.column(cid, width=80, anchor="center")

        # Montar 1 row de exemplo
        row_vals = []
        for i, cinfo in enumerate(campos):
            # Exibe os fixos (primeiro) ou (livre)
            fixos = cinfo.get("valores_fixos",[])
            if fixos:
                if isinstance(fixos[0], dict):
                    row_vals.append(fixos[0].get("value",""))
                else:
                    row_vals.append(str(fixos[0]))
            else:
                row_vals.append("(livre)")

            if i < len(campos)-1:
                # exibe o separador. p.ex cinfo.get("separador","-")
                sep_ = cinfo.get("separador","-")
                row_vals.append(sep_)

        tv.insert("", tk.END, values=row_vals)

    def voltar(self):
        self.destroy()
        # Reabrir a janela 1, repassando self.lista_arquivos
        TelaAdicaoArquivos(
            lista_inicial=self.lista_arquivos,
            pasta_entrega=PASTA_ENTREGA_GLOBAL,
            numero_projeto=NUM_PROJETO_GLOBAL
        ).mainloop()

    def avancar(self):
        global TIPO_ENTREGA_GLOBAL
        for iid in self.tree.get_children():
            tags_ = self.tree.item(iid, "tags")
            if "mismatch" in tags_ or "missing" in tags_:
                messagebox.showwarning(
                    "Atenção",
                    "Há campos com erro (vermelho) ou faltando (amarelo). "
                    "Corrija antes de avançar."
                )
                return

        # --- Escolha AP / PE ---
        escolha = escolher_tipo_entrega(self)
        if escolha is None:        # usuário cancelou
            return
        TIPO_ENTREGA_GLOBAL = escolha
        self.destroy()
        TelaVerificacaoRevisao(self.lista_arquivos).mainloop()

# -----------------------------------------------------
# Janela 3: TelaVerificacaoRevisao
# -----------------------------------------------------
class TelaVerificacaoRevisao(tk.Tk):
    def __init__(self, lista_arquivos, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Verificação de Revisão")
        self.resizable(True, True)
        self.lista_arquivos = lista_arquivos
        self.geometry("800x600")
        self.diretorio = PASTA_ENTREGA_GLOBAL
        self.dados_anteriores = carregar_dados_anteriores(self.diretorio)
        self.primeira_entrega = (len(self.dados_anteriores) == 0)
        (self.arquivos_novos,
         self.arquivos_revisados,
         self.arquivos_alterados) = analisar_comparando_estado(self.lista_arquivos, self.dados_anteriores)
        todos_diretorio = listar_arquivos_no_diretorio(self.diretorio)
        self.obsoletos = identificar_obsoletos_custom(todos_diretorio)
        container = tk.Frame(self)
        container.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        if self.primeira_entrega:
            info_text = "Essa é a primeira análise. Esses foram os arquivos:"
        else:
            info = obter_info_ultima_entrega(self.dados_anteriores)
            info_text = f"Esses foram os arquivos analisados a partir da última entrega, {info}."
        lbl = tk.Label(container, text=info_text, font=("Arial", 12, "bold"))
        lbl.pack(pady=5)
        self.criar_tabela(container, "Arquivos novos", self.arquivos_novos)
        self.criar_tabela(container, "Arquivos revisados", self.arquivos_revisados)
        btn_frame = tk.Frame(container)
        btn_frame.pack(fill=tk.X, pady=5)
        tk.Button(btn_frame, text="Voltar", command=self.voltar).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Confirmar", command=self.confirmar).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text="Cancelar", command=lambda: (self.destroy(), sys.exit(0))).pack(side=tk.RIGHT, padx=5)

    def criar_tabela(self, parent, titulo, arr):
        lf = tk.LabelFrame(parent, text=titulo, font=("Arial", 11, "bold"))
        lf.pack(fill=tk.BOTH, expand=True, pady=5)
        scrollbar_local = tk.Scrollbar(lf, orient="vertical", width=20)
        scrollbar_local.pack(side=tk.RIGHT, fill=tk.Y)
        cols = ("Nome do arquivo","Revisão","Data de modificação")
        tree = ttk.Treeview(lf, columns=cols, show="headings", height=5, yscrollcommand=scrollbar_local.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_local.config(command=tree.yview)
        style = ttk.Style()
        style.configure("Treeview", rowheight=24)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=200, anchor='w', stretch=True)
        if not arr:
            tree.insert("", tk.END, values=("Nenhum","",""))
        else:
            for (rv,a,tam,cam,dmod) in arr:
                tree.insert("", tk.END, values=(a, rv if rv else "Sem Revisão", dmod))
        return tree

    def voltar(self):
        self.destroy()
        TelaVerificacaoNomenclatura(self.lista_arquivos).mainloop()

    def confirmar(self):
        global TIPO_ENTREGA_GLOBAL
        if self.arquivos_alterados:
            self.withdraw()
            janela_erro_revisao(self.arquivos_alterados)
            self.deiconify()
        if not messagebox.askyesno("Confirmação Final", "Confirma que estes arquivos estão corretos?"):
            self.destroy()
            sys.exit(0)
        if not messagebox.askyesno("Entrega Oficial", "Essa é uma entrega oficial?"):
            self.destroy()
            sys.exit(0)
        try:
            if TIPO_ENTREGA_GLOBAL:
                criar_pasta_entrega_ap_pe(
                    self.diretorio,
                    TIPO_ENTREGA_GLOBAL,
                    self.arquivos_novos + self.arquivos_revisados + self.arquivos_alterados
                )
            subdir = "AP" if TIPO_ENTREGA_GLOBAL == "AP" else "PE"
            pasta_base = os.path.join(self.diretorio, subdir)
            entregas_ativas = [
                    d for d in os.listdir(pasta_base)
                if d.startswith(("1.AP - Entrega-", "2.PE - Entrega-"))
                and not d.endswith("-OBSOLETO")
            ]
            if entregas_ativas:                       
                pasta_destino = os.path.join(
                    pasta_base,
                    max(
                        entregas_ativas,
                        key=lambda n: int(re.search(r"(\d+)$", n).group(1)),
                    ),
                )

                def _redir(lista):
                    for i, tup in enumerate(list(lista)):
                        nome = os.path.basename(tup[3])
                        lista[i] = tup[:3] + (
                            os.path.join(pasta_destino, nome),
                        ) + tup[4:]
                _redir(self.arquivos_novos)
                _redir(self.arquivos_revisados)
                _redir(self.arquivos_alterados)

        except Exception as e:
            messagebox.showerror(
                "Erro",
                f"Falha ao criar/copiar pasta de entrega AP/PE:\n{e}"
            )
        pos_processamento(
            self.primeira_entrega,
            self.diretorio,
            self.dados_anteriores,
            self.arquivos_novos,
            self.arquivos_revisados,
            self.arquivos_alterados,
            self.obsoletos,
            TIPO_ENTREGA_GLOBAL,
        )
        sys.exit(0)

def main():
    while True:  # seleção de projeto
        num_proj, caminho_proj = janela_selecao_projeto()
        if not num_proj or not caminho_proj:
            return
        while True:  # seleção de disciplina
            pasta_entrega, voltar_proj = janela_selecao_disciplina(num_proj, caminho_proj)
            if voltar_proj:
                break  # volta para seleção de projeto
            if not pasta_entrega:
                return  # usuário cancelou

            global NOMENCLATURA_GLOBAL, PASTA_ENTREGA_GLOBAL, NUM_PROJETO_GLOBAL
            NOMENCLATURA_GLOBAL = carregar_nomenclatura_json(num_proj)
            PASTA_ENTREGA_GLOBAL = pasta_entrega
            NUM_PROJETO_GLOBAL = num_proj

            tela = TelaVisualizacaoEntregaAnterior(
                pasta_entregas=pasta_entrega,
                projeto_num=num_proj,
                disciplina=os.path.basename(os.path.dirname(pasta_entrega))
            )
            tela.mainloop()
            if getattr(tela, "reabrir_disciplina", False):
                continue  # reabrir seleção de disciplina
            return