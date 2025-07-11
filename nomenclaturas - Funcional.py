#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ---------------------------------------------------------------------------
# Gerenciador de Nomenclaturas – GUI para criar/editar regras de nomenclatura
# ---------------------------------------------------------------------------

import os      # Acesso a arquivos/diretórios
import json    # Serialização/desserialização de JSON
import copy    # Cópia profunda (evitar mutações indesejadas)
import tkinter as tk                 # Toolkit gráfico base
from tkinter import ttk, messagebox  # Widgets estilizados + diálogos

# ---------------------------------------------------------------------------
# ToolTip – janela flutuante com dica de texto 
# ---------------------------------------------------------------------------
class ToolTip:  # Tooltip simples
    def __init__(self, widget, text):
        self.widget = widget          # Widget onde a dica aparece
        self.text = text              # Texto da dica
        self.tipwindow = None         # Janela da dica

        self.widget.bind("<Enter>", self.show_tip, add="+")   # Hover entra
        self.widget.bind("<Leave>", self.hide_tip, add="+")   # Hover sai

    def show_tip(self, *_):  # Exibe tooltip
        if self.tipwindow or not self.text.strip():
            return
        x = self.widget.winfo_rootx() + 20   # Posição X da janela
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5  # Posição Y
        self.tipwindow = tw = tk.Toplevel(self.widget)  # Cria toplevel
        tw.overrideredirect(True)   # Remove borda
        tw.geometry(f"+{x}+{y}")    # Move janela
        tw.attributes("-topmost", True)  # Mantém acima
        lbl = tk.Label(              # Texto da dica
            tw, text=self.text.rstrip(), justify=tk.LEFT,
            background="#ffffe0", relief=tk.SOLID, borderwidth=1,
            font=("tahoma", 9, "normal")
        )
        lbl.pack(ipadx=1)

    def hide_tip(self, *_):  # Fecha tooltip
        if self.tipwindow:
            self.tipwindow.destroy()
        self.tipwindow = None

# ---------------------------------------------------------------------------
# Constantes globais
# ---------------------------------------------------------------------------
CAMINHO_NOMENCLATURAS_JSON = r"G:\Drives compartilhados\OAE-JSONS\nomenclaturas.json"  # Regras salvas

# REGRA_PADRAO – template de campos padrão
REGRA_PADRAO = [
    # STATUS
    {"nome": "STATUS", "tipo": "Fixo",
     "valores_fixos": [{"value": "E", "desc": ""},
                       {"value": "C", "desc": ""},
                       {"value": "P", "desc": ""},
                       {"value": "A", "desc": ""}],
     "valor_padrao": "P", "separador": "-"},
    # CLIENTE
    {"nome": "CLIENTE", "tipo": "Fixo",
     "valores_fixos": [{"value": "SS ESJ", "desc": ""},
                       {"value": "OutroCliente", "desc": ""}],
     "valor_padrao": "SS ESJ", "separador": "-"},
    # N° DO PROJETO
    {"nome": "N° DO PROJETO", "tipo": "Fixo",
     "valores_fixos": [{"value": "429", "desc": ""}],
     "valor_padrao": "429", "separador": "-"},
    # ORGANIZAÇÃO
    {"nome": "ORGANIZAÇÃO", "tipo": "Fixo",
     "valores_fixos": [{"value": "OAE", "desc": ""},
                       {"value": "ABC", "desc": ""},
                       {"value": "XYZ", "desc": ""}],
     "valor_padrao": "OAE", "separador": "-"},
    # SIGLA DISCIPLINA
    {"nome": "SIGLA DISCIPLINA", "tipo": "Fixo",
     "valores_fixos": [{"value": "ELE", "desc": ""},
                       {"value": "MEC", "desc": ""},
                       {"value": "ARQ", "desc": ""}],
     "valor_padrao": "ELE", "separador": "-"},
    # FASE PROJETO
    {"nome": "FASE PROJETO", "tipo": "Fixo",
     "valores_fixos": [{"value": "PE", "desc": ""},
                       {"value": "EX", "desc": ""},
                       {"value": "OP", "desc": ""}],
     "valor_padrao": "PE", "separador": "-"},
    # TIPO DOCUMENTO
    {"nome": "TIPO DOCUMENTO", "tipo": "Fixo",
     "valores_fixos": [{"value": "DTE", "desc": ""},
                       {"value": "GT",  "desc": ""},
                       {"value": "PRJ", "desc": ""}],
     "valor_padrao": "DTE", "separador": "-"},
    # CONJUNTO
    {"nome": "CONJUNTO", "tipo": "Fixo",
     "valores_fixos": [{"value": "X", "desc": ""},
                       {"value": "Y", "desc": ""},
                       {"value": "Z", "desc": ""}],
     "valor_padrao": "X", "separador": "."},
    # N° DOCUMENTO
    {"nome": "N° DOCUMENTO", "tipo": "Fixo",
     "valores_fixos": [{"value": "001", "desc": ""},
                       {"value": "002", "desc": ""},
                       {"value": "999", "desc": ""}],
     "valor_padrao": "001", "separador": "-"},
    # BLOCO
    {"nome": "BLOCO", "tipo": "Fixo",
     "valores_fixos": [{"value": "BLH", "desc": ""},
                       {"value": "BLQ", "desc": ""},
                       {"value": "BLA", "desc": ""}],
     "valor_padrao": "BLH", "separador": "-"},
    # PAVIMENTO
    {"nome": "PAVIMENTO", "tipo": "Fixo",
     "valores_fixos": [{"value": "1PAV", "desc": ""},
                       {"value": "2PAV", "desc": ""},
                       {"value": "3PAV", "desc": ""}],
     "valor_padrao": "1PAV", "separador": "-"},
    # SUBSISTEMA
    {"nome": "SUBSISTEMA", "tipo": "Fixo",
     "valores_fixos": [{"value": "BSW", "desc": ""},
                       {"value": "FW",  "desc": ""},
                       {"value": "XX",  "desc": ""}],
     "valor_padrao": "BSW", "separador": "-"},
    # TIPO DO DESENHO
    {"nome": "TIPO DO DESENHO", "tipo": "Fixo",
     "valores_fixos": [{"value": "PTB", "desc": ""},
                       {"value": "COR", "desc": ""},
                       {"value": "DET", "desc": ""}],
     "valor_padrao": "PTB", "separador": "-"},
    # REVISÃO especial (sempre ao final)
    {"nome": "REVISÃO_ESPECIAL",
     "revisao_opcao": "Numérico", "revisao_prefixo": "R",
     "revisao_ndigitos": 2, "revisao_separador": "-"}
]

# ---------------------------------------------------------------------------
# NomenclaturaApp – janela principal
# ---------------------------------------------------------------------------
class NomenclaturaApp:
    """GUI principal da aplicação."""

    def __init__(self, master: tk.Tk):
        self.master = master                      # Referência raiz
        master.geometry("1200x600")               # Tamanho inicial
        master.title("Gerenciador de Nomenclaturas")

        # ---- Dados ----
        self.regras_dict = self.carregar_regras() # Todas as regras carregadas
        self.numero_projeto = None                # Projeto ativo
        self.campos = []                          # Campos sem revisão
        self.revisao_config = {}                  # Config da revisão

        # ---- Widgets/frames ----
        self.canvas_campos = None    # Canvas onde ficam as linhas
        self.items_campos = {}       # Mapa id->info (frame, índice)
        self.lbl_preview = None      # Label preview

        self.frame_inicial = None    # Frame da tela 1
        self.frame_revisao = None    # Frame config revisão

        self.criar_janela_inicial()  # Monta tela inicial

    # -----------------------------------------------------------------------
    # Tela inicial
    # -----------------------------------------------------------------------
    def criar_janela_inicial(self):
        """Tela 1 – digitar nº do projeto."""
        self.frame_inicial = ttk.Frame(self.master, padding=10)
        self.frame_inicial.pack(fill=tk.BOTH, expand=True)

        ttk.Label(self.frame_inicial,
                  text="Digite o número do projeto:").pack(pady=5)

        self.entry_num_projeto = ttk.Entry(self.frame_inicial, width=20)  # Input nº
        self.entry_num_projeto.pack(pady=5)

        ttk.Button(self.frame_inicial, text="Pesquisar",
                   command=self.pesquisar_regra).pack(pady=5)  # Btn pesquisar

    def pesquisar_regra(self):
        """Procura regra existente ou cria padrão."""
        numero = self.entry_num_projeto.get().strip()
        if not numero:  # Campo vazio
            messagebox.showwarning("Atenção", "Informe o número do projeto.")
            return

        self.numero_projeto = numero.upper()  # Normaliza
        self.frame_inicial.destroy()          # Fecha tela 1

        # Recupera lista de campos
        dados = (self.regras_dict.get(self.numero_projeto, {"campos": []})["campos"]
                 or copy.deepcopy(REGRA_PADRAO))  # Usa padrão se não existir

        # Separa revisão e campos normais
        self.campos = []
        self.revisao_config = {"revisao_opcao": "Numérico",
                               "revisao_prefixo": "R",
                               "revisao_ndigitos": 2,
                               "revisao_separador": "-"}
        for c in dados:
            if c.get("nome") == "REVISÃO_ESPECIAL":  # Config revisão
                self.revisao_config.update(
                    revisao_opcao=c.get("revisao_opcao", "Numérico"),
                    revisao_prefixo=c.get("revisao_prefixo", "R"),
                    revisao_ndigitos=c.get("revisao_ndigitos", 2),
                    revisao_separador=c.get("revisao_separador", "-"))
            else:
                self.campos.append(c)

        # Obriga "N° DOCUMENTO" a ser Flexível
        for campo in self.campos:
            if campo.get("nome") == "N° DOCUMENTO":
                campo["tipo"] = "Flexível"

        self.criar_janela_principal()  # Abre editor

    # -----------------------------------------------------------------------
    # Tela principal (editor)
    # -----------------------------------------------------------------------
    def criar_janela_principal(self):
        """Tela 2 – editor de regra."""
        self.master.title(f"REGRAS DE NOMECLATURA – PROJETO: {self.numero_projeto}")

        frm = ttk.Frame(self.master, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm,
                  text=f"EDIÇÃO DA REGRA – PROJETO: {self.numero_projeto}",
                  font=("Arial", 12, "bold")).pack(anchor=tk.W, pady=(0, 10))

        # Cabeçalho da tabela
        cab = ttk.Frame(frm)
        cab.pack(fill=tk.X)
        titulos = ["", "ADICIONAR", "CAMPO:", "SEPARADOR:", "TIPO:",
                   "VALOR EXEMPLO:", ""]
        for col, txt in enumerate(titulos):
            ttk.Label(cab, text=txt, font=("Arial", 10, "bold")).grid(
                row=0, column=col, padx=15 if col else 1)

        # Canvas p/ linhas de campos
        frame_canvas = ttk.Frame(frm)
        frame_canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas_campos = tk.Canvas(frame_canvas, bg="white",
                                       highlightthickness=1, highlightbackground="#ccc")
        self.canvas_campos.pack(fill=tk.BOTH, expand=True)

        # Cria linhas existentes
        for i, campo in enumerate(self.campos):
            self.criar_item_no_canvas(i, campo)

        # Configuração de revisão
        self.frame_revisao = ttk.LabelFrame(frm, text="Configuração de Revisão",
                                            padding=5)
        self.frame_revisao.pack(fill=tk.X, pady=10)
        self.criar_linha_revisao()  # Widgets de revisão

        # Rodapé
        bot = ttk.Frame(frm)
        bot.pack(fill=tk.X, side=tk.BOTTOM)

        prev = ttk.Frame(bot)
        prev.pack(fill=tk.X, pady=5)
        ttk.Label(prev, text="Pré-visualização da nomenclatura:"
                  ).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(prev, text="Atualizar",
                   command=self.atualizar_preview).pack(side=tk.LEFT)

        self.lbl_preview = ttk.Label(prev, text="", foreground="blue",
                                     font=("Arial", 10, "bold"))
        self.lbl_preview.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        ttk.Button(bot, text="Salvar Regra",
                   command=self.on_salvar_clicked).pack(anchor=tk.E, pady=(0, 5))

        # Primeira pré-visualização
        self.atualizar_preview()
        self.master.update_idletasks()
        self.reposicionar_itens()

    # -----------------------------------------------------------------------
    # Callbacks coluna
    # -----------------------------------------------------------------------
    def on_nome_changed(self, index, var_nome):
        """Mudou o nome do campo."""
        novo = var_nome.get()
        if novo == "NOVO_CAMPO":            # Pede criação
            self.janela_cadastrar_novo_campo(index)
        else:
            self.campos[index]["nome"] = novo
        self.atualizar_preview()

    def on_sep_changed(self, index, var_sep):
        """Mudou separador."""
        self.campos[index]["separador"] = var_sep.get()
        self.atualizar_preview()

    def on_tipo_changed(self, index, var_tipo):
        """Mudou tipo Fixo/Flexível."""
        self.campos[index]["tipo"] = var_tipo.get()
        self.reload_canvas()  # Precisa redesenhar linha

    def on_val_changed(self, index, var_val):
        """Mudou valor padrão."""
        self.campos[index]["valor_padrao"] = var_val.get()
        self.atualizar_preview()

    # -----------------------------------------------------------------------
    # Linha no canvas
    # -----------------------------------------------------------------------
    def criar_item_no_canvas(self, index, campo_data):
        """Insere widgets para um campo na posição index."""
        # Larguras
        w_arrow, w_trash, w_plus = 3, 5, 5
        w_cb_nome, w_sep, w_tipo, w_val, w_edit = 30, 8, 15, 24, 6

        f = ttk.Frame(self.canvas_campos)  # Container da linha

        # Botões mover/excluir/inserir
        ttk.Button(f, text="▲", width=w_arrow,
                   command=lambda i=index: self.mover_campo_cima(i)
                   ).grid(row=0, column=0)
        ttk.Button(f, text="▼", width=w_arrow,
                   command=lambda i=index: self.mover_campo_baixo(i)
                   ).grid(row=0, column=1)
        ttk.Button(f, text="🗑", width=w_trash,
                   command=lambda i=index: self.excluir_campo(i)
                   ).grid(row=0, column=2)
        ttk.Button(f, text="+", width=w_plus,
                   command=lambda i=index + 1: self.inserir_novo_campo(i)
                   ).grid(row=0, column=3)

        # Combobox nome
        var_nome = tk.StringVar(value=campo_data.get("nome", ""))
        cb_nomes = ttk.Combobox(
            f, textvariable=var_nome, width=w_cb_nome,
            values=["STATUS", "CLIENTE", "N° DO PROJETO", "ORGANIZAÇÃO",
                    "SIGLA DISCIPLINA", "FASE PROJETO", "TIPO DOCUMENTO",
                    "CONJUNTO", "N° DOCUMENTO", "BLOCO", "PAVIMENTO",
                    "SUBSISTEMA", "TIPO DO DESENHO", "NOVO_CAMPO"])
        cb_nomes.grid(row=0, column=4, padx=1)
        cb_nomes.bind("<<ComboboxSelected>>",
                      lambda e, i=index, v=var_nome: self.on_nome_changed(i, v))
        var_nome.trace_add("write",
                           lambda *_a, i=index, v=var_nome: self.on_nome_changed(i, v))

        # Separador
        var_sep = tk.StringVar(value=campo_data.get("separador", ""))
        ttk.Entry(f, textvariable=var_sep, width=w_sep
                  ).grid(row=0, column=5, padx=1)
        var_sep.trace_add("write",
                          lambda *_a, i=index, v=var_sep: self.on_sep_changed(i, v))

        # Tipo Fixo/Flexível
        var_tipo = tk.StringVar(value=campo_data.get("tipo", "Fixo"))
        cb_tipo = ttk.Combobox(f, textvariable=var_tipo, width=w_tipo,
                               values=["Fixo", "Flexível"], state="readonly")
        cb_tipo.grid(row=0, column=6, padx=1)
        cb_tipo.bind("<<ComboboxSelected>>",
                     lambda e, i=index, v=var_tipo: self.on_tipo_changed(i, v))

        # Valor (combobox ou entry)
        var_val = tk.StringVar(value=campo_data.get("valor_padrao", ""))
        lista_vals = campo_data.get("valores_fixos", [])
        cb_valores = ttk.Combobox(
            f, textvariable=var_val,
            values=[v["value"] for v in lista_vals],
            width=w_val, state="readonly")
        ent_val = ttk.Entry(f, textvariable=var_val, width=w_val)
        var_val.trace_add("write",
                          lambda *_a, i=index, v=var_val: self.on_val_changed(i, v))

        btn_editar = ttk.Button(f, text="Editar", width=w_edit,
                                command=lambda i=index: self.editar_valores_fixo(i))

        # Tooltip valores
        tt_text = "\n".join(f'{d["value"]} -> {d.get("desc", "")}'.rstrip(" -> ")
                            for d in lista_vals) or "Nenhuma descrição cadastrada."
        lbl_lupa = ttk.Label(f, text="🔍", width=2, anchor=tk.CENTER,
                             foreground="blue", cursor="hand2")
        ToolTip(lbl_lupa, tt_text)

        # Layout depende do tipo
        if var_tipo.get() == "Fixo":
            cb_valores.grid(row=0, column=7, padx=1)
            btn_editar.grid(row=0, column=8, padx=1)
            lbl_lupa.grid(row=0, column=9, padx=1)
        else:
            ent_val.grid(row=0, column=7, padx=1)

        # Adiciona linha no canvas
        item_id = self.canvas_campos.create_window(0, 0, window=f, anchor="nw")
        self.items_campos[item_id] = {"frame": f, "index": index}

    # -----------------------------------------------------------------------
    # Operações CRUD linha
    # -----------------------------------------------------------------------
    def inserir_novo_campo(self, pos):
        """Insere campo padrão na posição 'pos'."""
        self.campos.insert(pos, {"nome": "NOVO_CAMPO", "tipo": "Fixo",
                                 "valores_fixos": [{"value": "EX1", "desc": ""},
                                                   {"value": "EX2", "desc": ""}],
                                 "valor_padrao": "EX1", "separador": "-"})
        self.reload_canvas()

    def excluir_campo(self, index):
        """Remove campo do índice."""
        if 0 <= index < len(self.campos):
            self.campos.pop(index)
            self.reload_canvas()

    def mover_campo_cima(self, index):
        """Move campo para cima."""
        if index > 0:
            self.campos[index - 1], self.campos[index] = self.campos[index], self.campos[index - 1]
            self.reload_canvas()

    def mover_campo_baixo(self, index):
        """Move campo para baixo."""
        if index < len(self.campos) - 1:
            self.campos[index], self.campos[index + 1] = self.campos[index + 1], self.campos[index]
            self.reload_canvas()

    # -----------------------------------------------------------------------
    # Canvas helpers
    # -----------------------------------------------------------------------
    def reload_canvas(self):
        """Redesenha todas as linhas."""
        self.canvas_campos.delete("all")
        self.items_campos.clear()
        for i, campo in enumerate(self.campos):
            self.criar_item_no_canvas(i, campo)
        self.atualizar_preview()
        self.master.update_idletasks()
        self.reposicionar_itens()

    def reposicionar_itens(self):
        """Posiciona frames manualmente (sem scrollbar)."""
        y = 10
        for item_id, info in sorted(self.items_campos.items(),
                                    key=lambda kv: kv[1]["index"]):
            h = info["frame"].winfo_reqheight()
            self.canvas_campos.coords(item_id, 10, y)
            y += h + 5
        self.canvas_campos.config(scrollregion=self.canvas_campos.bbox("all"))

    # -----------------------------------------------------------------------
    # Revisão
    # -----------------------------------------------------------------------
    def criar_linha_revisao(self):
        """Widgets de configuração da revisão."""
        frm = self.frame_revisao
        ttk.Label(frm, text="Tipo Revisão:").grid(row=0, column=0, pady=2, sticky=tk.W)

        self.var_rev_opcao = tk.StringVar(value=self.revisao_config["revisao_opcao"])
        cb_rev = ttk.Combobox(frm, textvariable=self.var_rev_opcao,
                              values=["Numérico", "Alfanumérico"], width=12,
                              state="readonly")
        cb_rev.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        cb_rev.bind("<<ComboboxSelected>>", lambda *_: self.atualizar_preview())

        ttk.Label(frm, text="Prefixo:").grid(row=0, column=2, padx=5, sticky=tk.W)
        self.var_rev_prefixo = tk.StringVar(value=self.revisao_config["revisao_prefixo"])
        ttk.Entry(frm, textvariable=self.var_rev_prefixo, width=8
                  ).grid(row=0, column=3, padx=5, sticky=tk.W)
        self.var_rev_prefixo.trace_add("write",
                                       lambda *_: self.atualizar_preview())

        ttk.Label(frm, text="N° de dígitos:").grid(row=0, column=4, padx=5, sticky=tk.W)
        self.var_rev_ndig = tk.IntVar(value=self.revisao_config["revisao_ndigitos"])
        ttk.Spinbox(frm, from_=1, to=99, textvariable=self.var_rev_ndig,
                    width=5, command=self.atualizar_preview
                    ).grid(row=0, column=5, padx=5, sticky=tk.W)

        ttk.Label(frm, text="Separador:").grid(row=0, column=6, padx=5, sticky=tk.W)
        self.var_rev_sep = tk.StringVar(value=self.revisao_config["revisao_separador"])
        ttk.Entry(frm, textvariable=self.var_rev_sep, width=8
                  ).grid(row=0, column=7, padx=5, sticky=tk.W)
        self.var_rev_sep.trace_add("write", lambda *_: self.atualizar_preview())

    # -----------------------------------------------------------------------
    # Preview + salvar
    # -----------------------------------------------------------------------
    def atualizar_preview(self):
        """Monta string de exemplo na label."""
        partes = []
        for i, campo in enumerate(self.campos):
            partes.append(campo.get("valor_padrao", "").strip())
            if i < len(self.campos) - 1:
                partes.append(campo.get("separador", ""))
        ndig = max(1, self.var_rev_ndig.get())  # Garantir >=1
        corpo = "1".rjust(ndig, "0") if self.var_rev_opcao.get() == "Numérico" else "A" * ndig
        partes.append(f'{self.var_rev_sep.get()}{self.var_rev_prefixo.get()}{corpo}')
        if self.lbl_preview:
            self.lbl_preview.config(text="".join(partes))

    def on_salvar_clicked(self):
        """Dispara salvamento."""
        self.atualizar_preview()
        self.salvar_regra()

    def salvar_regra(self):
        """Serializa regra no JSON."""
        campos_final = []
        for ordem, campo in enumerate(self.campos, start=1):
            campos_final.append({
                "ordem": str(ordem), "nome": campo.get("nome", ""),
                "tipo": campo.get("tipo", "Fixo"),
                "valores_fixos": campo.get("valores_fixos", []),
                "valor_padrao": campo.get("valor_padrao", ""),
                "separador": campo.get("separador", "")
            })
        campos_final.append({  # Bloco de revisão
            "nome": "REVISÃO_ESPECIAL",
            "revisao_opcao": self.var_rev_opcao.get(),
            "revisao_prefixo": self.var_rev_prefixo.get(),
            "revisao_ndigitos": self.var_rev_ndig.get(),
            "revisao_separador": self.var_rev_sep.get()
        })
        self.regras_dict[self.numero_projeto] = {"campos": campos_final}

        try:  # Grava arquivo
            with open(CAMINHO_NOMENCLATURAS_JSON, "w", encoding="utf-8") as f:
                json.dump(self.regras_dict, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("Sucesso",
                                f"Regra do projeto {self.numero_projeto} salva com sucesso!")
            self.master.destroy()
        except Exception as exc:
            messagebox.showerror("Erro", f"Erro ao salvar: {exc}")

    # -----------------------------------------------------------------------
    # Arquivo JSON
    # -----------------------------------------------------------------------
    @staticmethod
    def carregar_regras():
        """Carrega arquivo JSON de regras; retorna {} se inexistente/erro."""
        if os.path.exists(CAMINHO_NOMENCLATURAS_JSON):
            try:
                with open(CAMINHO_NOMENCLATURAS_JSON, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    # -----------------------------------------------------------------------
    # Modal – novo campo
    # -----------------------------------------------------------------------
    def janela_cadastrar_novo_campo(self, index):
        """Modal para criar campo fixo novo."""
        top = tk.Toplevel(self.master)          # Janela modal
        top.title("Cadastrar Novo Campo")
        top.geometry("1200x900")
        top.transient(self.master)
        top.grab_set()

        frm_top = ttk.Frame(top, padding=10)
        frm_top.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm_top, text="Nome do novo campo:").pack(anchor=tk.W)
        var_nomecampo = tk.StringVar()  # Nome digitado
        ttk.Entry(frm_top, textvariable=var_nomecampo, width=60
                  ).pack(pady=5, fill=tk.X)

        valores_local = []  # Valores fixos temporários

        frm_list = tk.Frame(frm_top)  # Lista valores
        frm_list.pack(fill=tk.BOTH, expand=True)

        def redesenhar_lista():
            """Atualiza lista na UI."""
            for w in frm_list.winfo_children():
                w.destroy()
            for i, item in enumerate(valores_local):
                rowf = tk.Frame(frm_list); rowf.pack(fill=tk.X, pady=2)

                def mk_up(idx=i):   # Sobe valor
                    def do_up():
                        if idx > 0:
                            valores_local[idx], valores_local[idx - 1] = \
                                valores_local[idx - 1], valores_local[idx]
                            redesenhar_lista()
                    return do_up

                def mk_down(idx=i):  # Desce valor
                    def do_down():
                        if idx < len(valores_local) - 1:
                            valores_local[idx], valores_local[idx + 1] = \
                                valores_local[idx + 1], valores_local[idx]
                            redesenhar_lista()
                    return do_down

                def mk_del(idx=i):   # Remove valor
                    def do_del():
                        valores_local.pop(idx)
                        redesenhar_lista()
                    return do_del

                tk.Button(rowf, text="▲", command=mk_up()).pack(side=tk.LEFT)
                tk.Button(rowf, text="▼", command=mk_down()).pack(side=tk.LEFT)
                tk.Button(rowf, text="🗑", command=mk_del()).pack(side=tk.LEFT, padx=5)

                var_val = tk.StringVar(value=item.get("value", ""))  # Valor
                tk.Entry(rowf, textvariable=var_val, width=20
                         ).pack(side=tk.LEFT, padx=5)
                var_val.trace_add("write", lambda *_a, ref=item, v=var_val:
                                  ref.update(value=v.get()))

                var_desc = tk.StringVar(value=item.get("desc", ""))  # Descrição
                tk.Entry(rowf, textvariable=var_desc, width=50
                         ).pack(side=tk.LEFT, padx=5)
                var_desc.trace_add("write", lambda *_a, ref=item, v=var_desc:
                                   ref.update(desc=v.get()))
        redesenhar_lista()

        # Linha adicionar
        frm_add = tk.Frame(frm_top)
        frm_add.pack(fill=tk.X, pady=5)
        var_val = tk.StringVar(); tk.Entry(frm_add, textvariable=var_val,
                                           width=25).pack(side=tk.LEFT, padx=5)
        ttk.Label(frm_add, text="(valor)").pack(side=tk.LEFT)
        var_desc = tk.StringVar(); tk.Entry(frm_add, textvariable=var_desc,
                                            width=50).pack(side=tk.LEFT, padx=5)
        ttk.Label(frm_add, text="(descrição)").pack(side=tk.LEFT)

        def add_valor():  # Adiciona valor à lista
            v = var_val.get().strip(); d = var_desc.get().strip()
            if v:
                valores_local.append({"value": v, "desc": d})
                var_val.set(""); var_desc.set("")
                redesenhar_lista()
        ttk.Button(frm_add, text="Adicionar valor",
                   command=add_valor).pack(side=tk.LEFT, padx=10)

        # Botões OK / Cancelar
        frm_btns = ttk.Frame(frm_top); frm_btns.pack(side=tk.BOTTOM, pady=5)

        def on_ok():  # Valida e grava
            nm = var_nomecampo.get().strip()
            if not nm:
                messagebox.showwarning("Atenção", "Informe o nome do novo campo."); return
            if not valores_local:
                messagebox.showwarning("Atenção", "Adicione ao menos um valor."); return
            self.campos[index]["nome"] = nm
            self.campos[index]["tipo"] = "Fixo"
            self.campos[index]["valores_fixos"] = valores_local
            self.campos[index]["valor_padrao"] = valores_local[0]["value"]
            top.destroy()
            self.reload_canvas()
        ttk.Button(frm_btns, text="OK", command=on_ok
                   ).pack(side=tk.LEFT, padx=10)
        ttk.Button(frm_btns, text="Cancelar", command=top.destroy
                   ).pack(side=tk.LEFT, padx=10)

    # -----------------------------------------------------------------------
    # Modal – editar valores fixos
    # -----------------------------------------------------------------------
    def editar_valores_fixo(self, index):
        """Modal para editar valores fixos de campo existente."""
        if not (0 <= index < len(self.campos)): return
        campo = self.campos[index]
        if campo.get("tipo") != "Fixo": return

        top = tk.Toplevel(self.master); top.title(
            f"Editar valores fixos – {campo.get('nome', '')}")  # Título
        top.geometry("1200x900"); top.transient(self.master); top.grab_set()

        valores_list = campo.setdefault("valores_fixos", [])

        frm_main = ttk.Frame(top, padding=10); frm_main.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm_main, text="Edite valores (valor + descrição)."
                  ).pack(anchor=tk.W, pady=5)

        frm_list = ttk.Frame(frm_main); frm_list.pack(fill=tk.BOTH, expand=True)

        def redesenhar():
            """Atualiza listagem."""
            for w in frm_list.winfo_children(): w.destroy()
            for i, item in enumerate(valores_list):
                rowf = tk.Frame(frm_list); rowf.pack(fill=tk.X, pady=3)

                def mk_up(idx=i):
                    def do_up():
                        if idx > 0:
                            valores_list[idx], valores_list[idx - 1] = \
                                valores_list[idx - 1], valores_list[idx]; redesenhar()
                    return do_up
                def mk_down(idx=i):
                    def do_down():
                        if idx < len(valores_list) - 1:
                            valores_list[idx], valores_list[idx + 1] = \
                                valores_list[idx + 1], valores_list[idx]; redesenhar()
                    return do_down
                def mk_del(idx=i):
                    def do_del():
                        valores_list.pop(idx); redesenhar()
                    return do_del

                tk.Button(rowf, text="▲", command=mk_up()).pack(side=tk.LEFT)
                tk.Button(rowf, text="▼", command=mk_down()).pack(side=tk.LEFT)
                tk.Button(rowf, text="🗑", command=mk_del()).pack(side=tk.LEFT, padx=5)

                var_val = tk.StringVar(value=item.get("value", ""))
                tk.Entry(rowf, textvariable=var_val, width=20
                         ).pack(side=tk.LEFT, padx=5)
                var_val.trace_add("write", lambda *_a, ref=item, v=var_val:
                                  ref.update(value=v.get()))

                var_desc = tk.StringVar(value=item.get("desc", ""))
                tk.Entry(rowf, textvariable=var_desc, width=50
                         ).pack(side=tk.LEFT, padx=5)
                var_desc.trace_add("write", lambda *_a, ref=item, v=var_desc:
                                   ref.update(desc=v.get()))
        redesenhar()

        # Linha adicionar novo
        frm_add = tk.Frame(frm_main); frm_add.pack(fill=tk.X, pady=5)
        var_new_val = tk.StringVar(); tk.Entry(frm_add, textvariable=var_new_val,
                                               width=20).pack(side=tk.LEFT, padx=5)
        ttk.Label(frm_add, text="(valor)").pack(side=tk.LEFT)
        var_new_desc = tk.StringVar(); tk.Entry(frm_add, textvariable=var_new_desc,
                                                width=50).pack(side=tk.LEFT, padx=5)
        ttk.Label(frm_add, text="(descrição)").pack(side=tk.LEFT)

        def add_new():
            v = var_new_val.get().strip(); d = var_new_desc.get().strip()
            if v:
                valores_list.append({"value": v, "desc": d})
                var_new_val.set(""); var_new_desc.set(""); redesenhar()
        ttk.Button(frm_add, text="Adicionar valor", command=add_new
                   ).pack(side=tk.LEFT, padx=10)

        # OK / Cancelar
        frm_btns = ttk.Frame(frm_main); frm_btns.pack(side=tk.BOTTOM, pady=5)
        def on_ok():
            vp = campo.get("valor_padrao", "")
            list_values = [i["value"] for i in valores_list]
            if vp not in list_values and list_values:
                campo["valor_padrao"] = list_values[0]  # Ajusta valor padrão
            top.destroy(); self.reload_canvas()
        ttk.Button(frm_btns, text="OK", command=on_ok
                   ).pack(side=tk.LEFT, padx=5)
        ttk.Button(frm_btns, text="Cancelar", command=top.destroy
                   ).pack(side=tk.LEFT, padx=5)

# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------
def main():  # Função main
    root = tk.Tk()           # Cria raiz
    NomenclaturaApp(root)    # Instancia app
    root.mainloop()          # Loop da GUI

if __name__ == "__main__":  # Se executado diretamente
    main()
