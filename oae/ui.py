"""Tkinter user interface for delivery management."""
# Docstring: Descreve o propósito do arquivo como a interface de usuário Tkinter para o gerenciamento de entregas.
from __future__ import annotations # Permite o uso de anotações de tipo de futuro (ex: list[str])
import os # Módulo para interagir com o sistema operacional (caminhos, diretórios)
import re # Módulo para expressões regulares
import sys # Módulo para funções do sistema (ex: sys.exit)
import json # Módulo para trabalhar com dados JSON
import datetime # Módulo para trabalhar com datas e horas
import tkinter as tk # Importa o módulo Tkinter para a criação da GUI
from tkinter import filedialog, messagebox, ttk # Importa componentes específicos do Tkinter:
                                                # filedialog: para diálogos de seleção de arquivo/diretório
                                                # messagebox: para exibir caixas de mensagem
                                                # ttk: Themed Tkinter, para widgets com aparência nativa do OS
from typing import Dict, List, Sequence, Tuple # Tipos para anotações de tipo
import logging # Módulo para registro de eventos (logs)

# Importa funções e constantes específicas do módulo interno .file_ops
from .file_ops import (
    PROJETOS_JSON, # Caminho para o arquivo JSON que lista os projetos
    NOMENCLATURAS_JSON, # Caminho para o arquivo JSON de regras de nomenclatura
    GRUPOS_EXT, # Dicionário com grupos de extensões de arquivo
    criar_pasta_entrega_ap_pe, # Cria a estrutura de pastas para uma nova entrega
    carregar_nomenclatura_json, # Carrega regras de nomenclatura de um JSON
    carregar_ultimo_diretorio, # Carrega o último diretório usado (para diálogos de arquivo)
    salvar_ultimo_diretorio, # Salva o último diretório usado
    extrair_numero_arquivo, # Extrai um número de um nome de arquivo
    split_including_separators, # Divide uma string, mantendo os separadores
    verificar_tokens, # Verifica tokens de nomenclatura contra regras
    identificar_obsoletos_custom, # Identifica arquivos obsoletos
    identificar_nome_com_revisao, # Identifica o nome de um arquivo com revisão
    carregar_dados_anteriores, # Carrega dados de entregas anteriores
    obter_info_ultima_entrega, # Obtém informações da última entrega
    listar_arquivos_no_diretorio, # Lista arquivos em um dado diretório
    analisar_comparando_estado, # Analisa arquivos comparando com um estado anterior
    pos_processamento, # Função para pós-processamento de dados
    folder_mais_recente, # Encontra a pasta de entrega mais recente
    _safe_json_load, # Função para carregar JSON de forma segura
)

try:
    from send2trash import send2trash  # type: ignore # Tenta importar 'send2trash' para mover arquivos para a lixeira
except Exception:
    send2trash = None  # pragma: no cover - optional dependency # Se falhar (não instalado), define como None. 'pragma: no cover' é uma diretiva para ferramentas de cobertura de código.

# Variáveis globais para armazenar o estado atual da aplicação
PASTA_ENTREGA_GLOBAL: str | None = None # Armazena o caminho da pasta de entrega selecionada
NOMENCLATURA_GLOBAL: Dict | None = None # Armazena as regras de nomenclatura carregadas
NUM_PROJETO_GLOBAL: str | None = None # Armazena o número do projeto selecionado
TIPO_ENTREGA_GLOBAL: str | None = None # Armazena o tipo de entrega selecionado (AP/PE)

logger = logging.getLogger(__name__) # Configura um logger para este módulo

def _center(win: tk.Toplevel | tk.Tk, parent: tk.Toplevel | tk.Tk | None = None) -> None:
    # Centraliza uma janela (Toplevel ou Tk) na tela ou em relação a uma janela pai.
    win.update_idletasks() # Garante que os cálculos de dimensão da janela estejam atualizados
    w, h = win.winfo_width(), win.winfo_height() # Obtém largura e altura da janela
    if parent: # Se um pai for especificado, centraliza em relação a ele
        parent.update_idletasks() # Atualiza as tarefas ociosas do pai
        x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2 # Calcula a posição X
        y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2 # Calcula a posição Y
    else: # Caso contrário, centraliza na tela
        x = (win.winfo_screenwidth() - w) // 2 # Calcula a posição X na tela
        y = (win.winfo_screenheight() - h) // 2 # Calcula a posição Y na tela
    win.geometry(f"{w}x{h}+{x}+{y}") # Define a geometria da janela (largura x altura + posição X + posição Y)

def escolher_tipo_entrega(master: tk.Toplevel | tk.Tk, size: tuple[int, int] = (480, 280)) -> str | None:
    # Abre uma janela modal para o usuário escolher o tipo de entrega (AP ou PE).
    escolha = {"val": None} # Dicionário para armazenar a escolha do usuário
    selecionado = tk.StringVar(value="AP") # Variável Tkinter para rastrear a opção selecionada, padrão "AP"
    win = tk.Toplevel(master) # Cria uma nova janela de nível superior (filha da janela 'master')
    win.withdraw() # Esconde a janela inicialmente para evitar flashes visuais
    win.title("Tipo de Entrega") # Define o título da janela
    win.transient(master) # Faz a janela se comportar como uma janela transitória (sempre acima da master)
    win.grab_set() # Captura todos os eventos de entrada, tornando a janela modal
    win.resizable(False, False) # Impede que a janela seja redimensionada
    win.geometry(f"{size[0]}x{size[1]}") # Define o tamanho da janela

    # Rótulo de instrução
    ttk.Label(win, text="Escolha o tipo de entrega:", font=("Arial", 11, "bold")).pack(pady=(12, 12))
    area = tk.Frame(win) # Cria um frame para os cartões de seleção
    area.pack(expand=True) # Empacota o frame, permitindo que ele se expanda

    CARD_W = (size[0] // 2) - 40 # Largura calculada para os cartões
    CARD_H = size[1] - 120 # Altura calculada para os cartões
    FONT_BIG = ("Arial", 28, "bold") # Fonte grande para os ícones dos cartões

    def _build_card(parent, texto, valor):
        # Constrói um "cartão" de seleção visualmente atraente.
        frame = tk.Frame(parent, width=CARD_W, height=CARD_H, borderwidth=2, relief="ridge", bg="#f0f0f0")
        frame.pack_propagate(False) # Impede que o frame redimensione para caber seus filhos
        lbl_icon = tk.Label(frame, text=valor, font=FONT_BIG, bg="#f0f0f0") # Rótulo para o ícone (AP/PE)
        lbl_icon.pack(expand=True) # Empacota o ícone, expandindo para preencher o espaço
        lbl_txt = tk.Label(frame, text=texto, font=("Arial", 10), bg="#f0f0f0") # Rótulo para o texto descritivo
        lbl_txt.pack(pady=(0, 6)) # Empacota o texto, com um pequeno preenchimento inferior

        def _on_click(*_):
            # Função chamada ao clicar no cartão.
            selecionado.set(valor) # Atualiza a variável 'selecionado' com o valor do cartão
            _update_highlight() # Atualiza o destaque visual dos cartões

        # Associa o evento de clique a todos os elementos do cartão
        frame.bind("<Button-1>", _on_click)
        lbl_icon.bind("<Button-1>", _on_click)
        lbl_txt.bind("<Button-1>", _on_click)
        return frame # Retorna o frame do cartão

    def _update_highlight():
        # Atualiza o destaque visual dos cartões com base na seleção.
        for card, val in ((card_ap, "AP"), (card_pe, "PE")): # Itera sobre os cartões AP e PE
            if selecionado.get() == val: # Se o valor do cartão for o selecionado
                card.config(bg="#dbe9ff", highlightbackground="#4e9af1") # Destaca com cor de fundo azul claro e borda azul
            else:
                card.config(bg="#f0f0f0", highlightbackground="#d0d0d0") # Retorna ao estilo padrão (cinza claro)

    # Cria os cartões AP e PE
    card_ap = _build_card(area, "Anteprojeto – 1.AP", "AP")
    card_pe = _build_card(area, "Projeto Executivo – 2.PE", "PE")

    # Empacota os cartões lado a lado
    card_ap.pack(side="left", padx=10, pady=5, expand=True, fill="both")
    card_pe.pack(side="left", padx=10, pady=5, expand=True, fill="both")

    _update_highlight() # Garante que o cartão inicial ("AP") esteja destacado
    btn_box = ttk.Frame(win) # Cria um frame para os botões OK/Cancelar
    btn_box.pack(pady=(6, 12)) # Empacota com preenchimento

    def _on_ok_click():
        # Função chamada ao clicar no botão OK.
        logger.debug("Escolher tipo OK clicked") # Loga o clique
        try:
            escolha.update(val=selecionado.get()) # Armazena a escolha final
            win.destroy() # Fecha a janela
            logger.debug("Escolher tipo OK executed successfully") # Loga o sucesso
        except Exception:
            logger.exception("Erro ao confirmar tipo de entrega") # Loga qualquer exceção

    def _on_cancel_click():
        # Função chamada ao clicar no botão Cancelar.
        logger.debug("Escolher tipo Cancelar clicked") # Loga o clique
        win.destroy() # Fecha a janela
        logger.debug("Escolher tipo Cancelar executed successfully") # Loga o sucesso

    # Cria e empacota os botões OK e Cancelar
    ttk.Button(btn_box, text="OK", command=_on_ok_click).pack(side="left", padx=6)
    ttk.Button(btn_box, text="Cancelar", command=_on_cancel_click).pack(side="left", padx=6)

    _center(win, master) # Centraliza a janela modal
    win.deiconify() # Mostra a janela (se estava escondida)
    master.wait_window(win) # Faz a janela mestre esperar o fechamento desta janela modal
    return escolha["val"] # Retorna o valor escolhido (AP ou PE) ou None se cancelado

def _carregar_projetos() -> List[Tuple[str, str]]:
    # Função interna para carregar a lista de projetos do arquivo JSON.
    if not os.path.exists(PROJETOS_JSON): # Verifica se o arquivo de projetos existe
        messagebox.showerror("Erro", f"Arquivo de projetos não encontrado:\n{PROJETOS_JSON}") # Exibe erro
        sys.exit(1) # Sai do programa
    with open(PROJETOS_JSON, "r", encoding="utf-8") as f: # Abre o arquivo JSON
        projetos_dict = _safe_json_load(f) # Carrega o JSON de forma segura
    return list(projetos_dict.items()) # Retorna uma lista de tuplas (número_projeto, caminho_projeto)


def janela_selecao_projeto():
    """
    Abre uma janela com a lista de projetos.

    Colunas exibidas:
        • Número        (ex.: 231)
        • Nome do Projeto (ex.: OAE-231 - LAEPE)

    O valor interno ‘caminho’ continua sendo o path completo,
    mas na tabela o usuário vê apenas o nome da pasta.
    """
    root = tk.Tk() # Cria a janela principal do Tkinter

    root.geometry("600x400")   # Define o tamanho inicial da janela (largura x altura)
    root.minsize(900, 500)     # Impede que a janela seja redimensionada para um tamanho muito pequeno

    root.title("Selecionar Projeto") # Define o título da janela
    root.resizable(True, True) # Permite redimensionar a janela (largura e altura)

    if not os.path.exists(PROJETOS_JSON): # Verifica novamente se o arquivo de projetos existe
        messagebox.showerror("Erro",
                             f"Arquivo de projetos não encontrado:\n{PROJETOS_JSON}") # Exibe erro
        sys.exit(1) # Sai do programa

    # ---- lê o JSON e cria lista (numero, caminho, nome_exibicao) ----
    with open(PROJETOS_JSON, "r", encoding="utf-8") as f: # Abre o arquivo JSON
        temp = json.load(f).items() # Carrega os itens do JSON
    # Cria uma lista de tuplas: (número do projeto, caminho completo, nome de exibição da pasta)
    projetos = [(n, c, os.path.basename(c)) for n, c in temp]

    sel = {"num": None, "path": None} # Dicionário para armazenar a seleção final do usuário

    # --------------- callbacks internos ---------------
    def confirmar():
        # Função chamada ao confirmar a seleção do projeto.
        logger.debug("Selecionar projeto confirmar clicked") # Loga o clique
        try:
            sel_i = tree.selection() # Obtém os itens selecionados na treeview
            if not sel_i: # Se nada estiver selecionado
                return # Sai da função
            iid = sel_i[0] # Pega o ID do primeiro item selecionado
            sel["num"] = tree.set(iid, "Número") # Obtém o número do projeto da coluna "Número"
            índice = int(tree.item(iid, "text")) # Obtém o índice da linha (armazenado como 'text' no item)
            sel["path"] = projetos[índice][1] # Usa o índice para obter o caminho completo do projeto original
            root.destroy() # Fecha a janela
            logger.debug("Selecionar projeto confirmar executed successfully") # Loga o sucesso
        except Exception:
            logger.exception("Erro ao confirmar projeto") # Loga qualquer exceção

    def filtrar(*_):
        # Função para filtrar a lista de projetos na treeview com base no texto digitado.
        termo = entrada.get().lower() # Obtém o texto da entrada e o converte para minúsculas
        tree.delete(*tree.get_children()) # Limpa todos os itens existentes na treeview
        for idx, (n, _, nome_disp) in enumerate(projetos): # Itera sobre a lista de projetos
            if termo in nome_disp.lower(): # Se o termo de busca estiver no nome de exibição
                # Insere o projeto na treeview
                tree.insert("", tk.END, text=str(idx), # 'text' armazena o índice original (invisível ao usuário)
                            values=(n, nome_disp)) # 'values' são as colunas visíveis

    # --------------- construção da UI ---------------
    # Rótulo de instrução para a entrada de filtro
    tk.Label(root, text="Digite nome ou parte do nome do projeto:"
             ).pack(anchor="w", padx=10, pady=5)

    entrada = tk.Entry(root) # Campo de entrada de texto para o filtro
    entrada.pack(fill=tk.X, padx=10) # Empacota o campo de entrada, preenchendo a largura
    entrada.bind("<KeyRelease>", filtrar) # Liga o evento de liberação de tecla à função de filtro
    
    def _on_return(_):
        # Callback para a tecla Enter no campo de entrada
        confirmar() # Chama a função de confirmação

    entrada.bind("<Return>", _on_return) # Liga o evento Return (Enter) à função _on_return

    cols = ("Número", "Nome do Projeto") # Nomes das colunas da treeview
    tree = ttk.Treeview(root, columns=cols, show="headings", height=15) # Cria a treeview
    for c in cols: # Para cada coluna
        tree.heading(c, text=c) # Define o cabeçalho da coluna
    tree.column("Número", width=60, anchor="w") # Define largura e alinhamento para a coluna "Número"
    tree.column("Nome do Projeto", anchor="w") # Define alinhamento para a coluna "Nome do Projeto"
    tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5) # Empacota a treeview, preenchendo e expandindo

    for idx, (n, _, nome_disp) in enumerate(projetos): # Popula a treeview com todos os projetos inicialmente
        tree.insert("", tk.END, text=str(idx), values=(n, nome_disp))

    tk.Button(root, text="Confirmar", command=confirmar).pack(pady=10) # Botão "Confirmar"
    root.mainloop() # Inicia o loop principal de eventos do Tkinter para esta janela
    return sel["num"], sel["path"] # Retorna o número e o caminho do projeto selecionado


def selecionar_pasta_entrega(diretorio_inicial: str):
    # Abre um diálogo para o usuário selecionar uma pasta de entrega.
    root = tk.Tk() # Cria uma instância Tkinter oculta
    root.withdraw() # Esconde a janela principal para não aparecer
    # Abre o diálogo de seleção de diretório
    pasta = filedialog.askdirectory(title="Selecione a pasta de entrega", initialdir=diretorio_inicial)
    root.destroy() # Destrói a instância Tkinter oculta
    return pasta # Retorna o caminho da pasta selecionada (ou string vazia se cancelado)


def janela_erro_revisao(arquivos_alterados):
    # Exibe uma janela de aviso quando arquivos modificados são encontrados sem alteração de revisão.
    janela = tk.Toplevel() # Cria uma nova janela de nível superior
    janela.title("Possível Erro de Revisão") # Define o título
    janela.configure(bg="#FFA07A") # Define a cor de fundo (laranja claro, para aviso)
    msg = (
        "Foi identificado que o tamanho dos arquivos abaixo mudou em relação à entrega anterior.\n"
        "Confirma que isso está correto?\n\n"
    )
    for rv, arq, *_ in arquivos_alterados: # Adiciona cada arquivo alterado à mensagem
        msg += f"{arq} - Revisão: {rv or 'Sem Revisão'}\n" # Inclui revisão ou "Sem Revisão"
    tk.Label(janela, text=msg, bg="#FFA07A", font=("Arial", 12)).pack(padx=10, pady=10) # Cria e empacota o rótulo da mensagem

    def _encerra():
        # Função chamada ao clicar em "Confirmar e sair".
        logger.debug("Erro revisao confirmar e sair clicked") # Loga o clique
        janela.destroy() # Fecha a janela
        logger.debug("Erro revisao confirm window closed") # Loga o fechamento
        sys.exit(0) # Sai do programa

    def _ignorar():
        # Função chamada ao clicar em "Ignorar".
        logger.debug("Erro revisao ignorar clicked") # Loga o clique
        janela.destroy() # Fecha a janela
        logger.debug("Erro revisao ignorar executed successfully") # Loga o sucesso

    # Botões para confirmar/sair ou ignorar
    tk.Button(janela, text="Confirmar e sair", command=_encerra).pack(pady=5)
    tk.Button(janela, text="Ignorar", command=janela.destroy).pack(pady=5)
    janela.grab_set() # Torna a janela modal
    janela.mainloop() # Inicia o loop principal para esta janela


def janela_selecao_disciplina(numero_proj: str, caminho_proj: str) -> tuple[str | None, bool]:
    # Exibe uma janela para o usuário selecionar uma disciplina dentro do projeto.
    # Retorna o caminho da pasta de entregas da disciplina selecionada e um flag para voltar.
    pasta_desenvol = os.path.join(caminho_proj, "3 Desenvolvimento") # Monta o caminho para a pasta "3 Desenvolvimento"
    if not os.path.isdir(pasta_desenvol): # Verifica se a pasta existe
        messagebox.showerror(
            "Erro",
            f"A pasta '3 Desenvolvimento' não foi encontrada em:\n{caminho_proj}"
        )
        return None, False # Retorna None e False se a pasta não for encontrada

    ignorar = ("3.0 compatibilização", "3.1 projetos externos") # Pastas de disciplina a serem ignoradas
    disciplinas = [] # Lista para armazenar os caminhos completos das disciplinas
    for nome in sorted(os.listdir(pasta_desenvol)): # Itera sobre os itens na pasta "3 Desenvolvimento"
        if any(nome.strip().lower().startswith(term) for term in ignorar): # Se o nome começa com um termo a ignorar
            continue # Pula para o próximo item
        p = os.path.join(pasta_desenvol, nome) # Monta o caminho completo da subpasta
        if os.path.isdir(p): # Se for um diretório
            disciplinas.append(p) # Adiciona à lista de disciplinas

    if not disciplinas: # Se nenhuma disciplina for encontrada
        messagebox.showerror("Erro", "Nenhuma disciplina encontrada.")
        return None, False # Retorna None e False

    root = tk.Tk() # Cria a janela principal para a seleção de disciplina
    root.title(f"Projeto {numero_proj} – Selecionar Disciplina") # Define o título
    root.geometry("700x500") # Define o tamanho inicial
    root.minsize(700, 400) # Define o tamanho mínimo

    # Rótulo e campo de entrada para filtrar disciplinas
    tk.Label(root, text="Filtrar disciplina:").pack(anchor="w", padx=10, pady=5)
    var_filtro = tk.StringVar() # Variável para o texto do filtro
    entrada = tk.Entry(root, textvariable=var_filtro) # Campo de entrada
    entrada.pack(fill=tk.X, padx=10) # Empacota, preenchendo a largura

    main_content_frame = tk.Frame(root) # Frame principal para o conteúdo rolável
    main_content_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    canvas = tk.Canvas(main_content_frame, borderwidth=0, highlightthickness=0) # Canvas para a área rolável
    vsb = tk.Scrollbar(main_content_frame, orient="vertical", command=canvas.yview) # Barra de rolagem vertical
    canvas.configure(yscrollcommand=vsb.set) # Configura o canvas para usar a barra de rolagem
    vsb.pack(side=tk.RIGHT, fill=tk.Y) # Empacota a barra de rolagem à direita
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True) # Empacota o canvas à esquerda

    frame = tk.Frame(canvas, bg="#f4f4f4") # Frame interno onde os cartões de disciplina serão colocados
    canvas.create_window((0, 0), window=frame, anchor="nw") # Cria uma janela no canvas para o frame

    def _on_mousewheel(event):
        # Função para rolagem com a roda do mouse.
        if event.num == 4 or event.delta > 0: # Rola para cima
            canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0: # Rola para baixo
            canvas.yview_scroll(1, "units")

    # Liga os eventos da roda do mouse ao canvas
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    canvas.bind_all("<Button-4>", _on_mousewheel) # Linux/Unix
    canvas.bind_all("<Button-5>", _on_mousewheel) # Linux/Unix

    def _unbind_wheel():
        # Desvincula os eventos da roda do mouse.
        canvas.unbind_all("<MouseWheel>")
        canvas.unbind_all("<Button-4>")
        canvas.unbind_all("<Button-5>")

    CARD_W, CARD_H = 120, 120 # Dimensões dos cartões de disciplina
    sel = {"widget": None, "path": None} # Dicionário para armazenar o cartão selecionado e seu caminho
    voltar_flag = {"val": False} # Flag para indicar se o usuário clicou em "Voltar"

    # Declara nonlocal_result antes de _confirmar, para ser acessível e modificável
    nonlocal_result = [None] 

    def _render():
        # Renderiza (ou re-renderiza) os cartões de disciplina com base no filtro.
        for w in frame.winfo_children(): # Destrói todos os widgets existentes no frame
            w.destroy()

        termo = var_filtro.get().lower() # Obtém o termo de filtro atual
        # Filtra as disciplinas que correspondem ao termo
        paths = [p for p in disciplinas if termo in os.path.basename(p).lower()]
        # Calcula o número de colunas com base na largura do canvas e largura do cartão
        cols = max(1, (canvas.winfo_width() // (CARD_W + 20)))

        for idx, path in enumerate(paths): # Itera sobre as disciplinas filtradas
            r, c = divmod(idx, cols) # Calcula a linha e coluna para o posicionamento na grade
            card = tk.Frame(
                frame,
                width=CARD_W,
                height=CARD_H,
                bg="white",
                highlightthickness=1, # Borda fina
                highlightbackground="#c0c0c0", # Cor da borda
                relief="flat", # Estilo da borda
            )
            card.grid(row=r, column=c, padx=10, pady=10) # Posiciona o cartão na grade
            card.grid_propagate(False) # Impede que o frame redimensione para caber seus filhos

            lbl = tk.Label(
                card,
                text=os.path.basename(path), # Nome da disciplina
                wraplength=CARD_W - 10, # Quebra o texto se for muito longo
                justify="center", # Centraliza o texto
                font=("TkDefaultFont", 10),
            )
            lbl.place(relx=0.5, rely=0.5, anchor="center") # Posiciona o rótulo no centro do cartão

            def _select(widget=card, caminho=path):
                # Função para selecionar um cartão (visual e logicamente).
                if sel["widget"]: # Se já houver um cartão selecionado, remove o destaque
                    sel["widget"].config(bg="white", highlightbackground="#c0c0c0")
                widget.config(bg="#dbe9ff", highlightbackground="#4e9af1") # Destaca o novo cartão selecionado
                sel["widget"], sel["path"] = widget, caminho # Armazena o widget e o caminho do cartão selecionado

            # Liga eventos de clique (simples e duplo) aos cartões e seus rótulos
            card.bind("<Button-1>", lambda e, w=card, p=path: _select(w, p))
            card.bind("<Double-Button-1>", lambda e, p=path: _confirmar(p)) # Duplo clique confirma
            lbl.bind("<Button-1>", lambda e, w=card, p=path: _select(w, p))
            lbl.bind("<Double-Button-1>", lambda e, p=path: _confirmar(p))

    def _confirmar(caminho_sel):
        # Função chamada ao confirmar a seleção de uma disciplina.
        logger.debug("Selecao disciplina confirmar clicked") # Loga o clique
        try:
            if not caminho_sel: # Se nenhum caminho foi selecionado
                return # Sai da função
            pasta_entregas = None # Variável para armazenar o caminho da pasta "1.ENTREGAS"
            for folder in os.listdir(caminho_sel): # Itera sobre os subdiretórios da disciplina
                if folder.strip().lower().replace(" ", "") == "1.entregas": # Procura a pasta "1.ENTREGAS" (ignorando case e espaços)
                    pasta_entregas = os.path.join(caminho_sel, folder) # Monta o caminho completo
                    break # Encontrou, sai do loop
            if not pasta_entregas or not os.path.isdir(pasta_entregas): # Se a pasta "1.ENTREGAS" não for encontrada ou não for um diretório
                messagebox.showerror(
                    "Erro",
                    "A subpasta '1.ENTREGAS' não foi encontrada dentro da disciplina."
                )
                return # Sai da função
            nonlocal_result[0] = pasta_entregas # Armazena o caminho da pasta de entregas no resultado não-local
            _unbind_wheel() # Desvincula os eventos da roda do mouse
            root.destroy() # Fecha a janela
            logger.debug("Selecao disciplina confirmar executed successfully") # Loga o sucesso
        except Exception:
            logger.exception("Erro ao confirmar disciplina") # Loga qualquer exceção

    footer = tk.Frame(root) # Frame para o rodapé (botões)
    footer.pack(side=tk.BOTTOM, fill=tk.X, pady=10) # Empacota no rodapé
    button_frame = tk.Frame(footer) # Frame para centralizar os botões
    button_frame.pack(expand=True) # Empacota, permitindo expansão

    def _on_voltar():
        # Função chamada ao clicar no botão "Voltar".
        logger.debug("Selecao disciplina voltar clicked") # Loga o clique
        # 1) destrói a janela atual
        root.destroy()
        # 2) abre imediatamente a tela de seleção de projeto
        num, path = janela_selecao_projeto() # Chama a janela de seleção de projeto
        # 3) se o usuário cancelou, encerra o programa ou retorna valores nulos
        if not num or not path: # Se o usuário cancelou a seleção de projeto
            sys.exit(0)   # Sai do programa (ou pode ser 'return None, False' dependendo do fluxo desejado)
        # 4) reenfila o fluxo chamando novamente disciplina para o novo projeto
        # Retorna o resultado da nova chamada, que fechará esta função também.
        return janela_selecao_disciplina(num, path)
        
    btn_voltar = tk.Button(
        button_frame,
        text="Voltar",
        width=15,
        command=_on_voltar, # Liga ao callback _on_voltar
    )
    btn_confirmar = tk.Button(
        button_frame,
        text="Confirmar",
        width=15,
        command=lambda: _confirmar(sel["path"]), # Liga ao callback _confirmar com o caminho selecionado
    )
    btn_voltar.pack(side=tk.LEFT, padx=10) # Empacota o botão "Voltar"
    btn_confirmar.pack(side=tk.LEFT, padx=10) # Empacota o botão "Confirmar"

    var_filtro.trace_add("write", lambda *_: _render()) # Liga a variável de filtro à função _render para atualizar a exibição
    # Liga o evento de configuração do frame interno ao canvas para ajustar a região de rolagem
    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    # Liga o evento de configuração do canvas para re-renderizar e ajustar a largura dos itens quando o canvas é redimensionado
    canvas.bind(
        "<Configure>",
        lambda e: (_render(), canvas.itemconfig("all", width=canvas.winfo_width())),
    )

    nonlocal_result = [None] # Usado para retornar o resultado de funções aninhadas
    _render() # Renderiza os cartões inicialmente
    root.mainloop() # Inicia o loop principal de eventos para esta janela
    return nonlocal_result[0], voltar_flag["val"] # Retorna o caminho da pasta de entregas e o flag de "voltar"


class TelaVisualizacaoEntregaAnterior(tk.Tk):
    """Mostra arquivos da entrega AP/PE mais recente e permite renomear/excluir."""

    def __init__(self, pasta_entregas: str, projeto_num: str, disciplina: str, lista_inicial=None, *args, **kwargs):
        # Construtor da classe. Inicializa a janela e seus componentes.
        super().__init__(*args, **kwargs) # Chama o construtor da classe base tk.Tk
        self.title("Entrega Anterior – Visualização") # Define o título da janela
        self.geometry("1000x600") # Define o tamanho inicial
        self.resizable(True, True) # Permite redimensionar
        self.configure(bg="#f5f5f5") # Cor de fundo

        self.pasta_entregas = pasta_entregas # Caminho para a pasta '1.ENTREGAS'
        self.projeto_num = projeto_num # Número do projeto
        self.disciplina = disciplina # Nome da disciplina
        self.lista_inicial = lista_inicial # Lista de arquivos inicial (opcional, para testes ou pré-carregamento)
        self.reabrir_disciplina = False # Flag para indicar se a janela de disciplina deve ser reaberta

        # ------ estado ------
        self.tipo_var = tk.StringVar(value="AP") # Variável para selecionar o tipo de entrega (AP/PE) na UI
        self.lista_arquivos: list[tuple[str, str, int, str, str]] = [] # Lista de arquivos exibidos (rev, nome, tam, cam, dt)

        # UI ------------------------------------------------------------------
        header = tk.Frame(self, bg="#2c3e50") # Frame do cabeçalho
        header.pack(fill=tk.X) # Empacota, preenchendo a largura
        # Rótulo com informações do projeto e disciplina no cabeçalho
        tk.Label(header, text=f"Projeto {projeto_num}  •  {disciplina}", fg="white", bg="#2c3e50",
                 font=("Helvetica", 14, "bold")).pack(padx=10, pady=6, anchor="w")

        ctrl = tk.Frame(self) # Frame para controles (combobox, botões)
        ctrl.pack(fill=tk.X, padx=10, pady=(10, 5))
        tk.Label(ctrl, text="Visualizar entregas de:").pack(side=tk.LEFT) # Rótulo para o combobox

        cmb_tipo = ttk.Combobox(ctrl, values=["AP", "PE"], textvariable=self.tipo_var,
                                width=4, state="readonly") # Combobox para selecionar AP/PE
        cmb_tipo.pack(side=tk.LEFT, padx=5) # Empacota o combobox
        cmb_tipo.bind("<<ComboboxSelected>>", lambda e: self._carregar_entrega()) # Liga evento de seleção para recarregar arquivos

        # Botões de ação
        ttk.Button(ctrl, text="Tornar Obsoletos", command=self._excluir_selecionados).pack(side=tk.RIGHT, padx=5)
        ttk.Button(ctrl, text="Adicionar Arquivos", command=self._avancar).pack(side=tk.RIGHT, padx=5)
        ttk.Button(ctrl, text="Voltar", command=self._voltar).pack(side=tk.RIGHT, padx=5)

        # Tabela (Treeview) --------------------------------------------------------------
        tbl_frame = tk.Frame(self) # Frame para a tabela e barra de rolagem
        tbl_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.sb_y = tk.Scrollbar(tbl_frame, orient="vertical") # Barra de rolagem vertical
        self.sb_y.pack(side=tk.RIGHT, fill=tk.Y) # Empacota à direita

        self.tree = ttk.Treeview(
            tbl_frame,
            columns=("sel", "nome", "dt", "tam"), # Define as colunas
            show="headings", # Mostra apenas os cabeçalhos
            yscrollcommand=self.sb_y.set, # Liga a barra de rolagem à treeview
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True) # Empacota a treeview, preenchendo e expandindo
        self.sb_y.config(command=self.tree.yview) # Configura a barra de rolagem para controlar a treeview

        # Define os cabeçalhos das colunas
        self.tree.heading("sel", text="Sel")
        self.tree.heading("nome", text="Nome do arquivo", anchor="w")
        self.tree.heading("dt", text="Data mod.")
        self.tree.heading("tam", text="Tamanho (KB)")

        # Define as propriedades das colunas (largura, alinhamento)
        self.tree.column("sel", width=40, anchor="center")
        self.tree.column("nome", width=400, anchor="w")
        self.tree.column("dt", width=130, anchor="center")
        self.tree.column("tam", width=100, anchor="e")

        # Liga eventos da treeview a callbacks
        self.tree.bind("<Double-1>", self._iniciar_edicao_nome) # Duplo clique para iniciar edição de nome
        self.tree.bind("<Button-1>", self._toggle_checkbox) # Clique simples para alternar checkbox de seleção

        self.checked = {} # Dicionário para armazenar o estado dos checkboxes (quais arquivos estão selecionados)

        if self.lista_inicial: # Se uma lista inicial de arquivos foi fornecida (ex: para testes)
            self._carregar_lista_inicial() # Carrega essa lista
        else:
            self._carregar_entrega() # Caso contrário, carrega a entrega mais recente

    # ------------------------------------------------------------------
    # lógica de descoberta de entrega mais recente
    # ------------------------------------------------------------------
    def _folder_mais_recente(self, tipo: str) -> str | None:
        # Wrapper para a função externa folder_mais_recente.
        return folder_mais_recente(self.pasta_entregas, tipo)
    def _carregar_entrega(self):
        # Este método é responsável por carregar e exibir os arquivos da entrega mais recente (AP ou PE)
        # na Treeview da interface.

        # Garante que a coluna 'cam_full' (caminho completo do arquivo) exista na Treeview.
        # Ela é usada internamente para operações de arquivo, mas é oculta (width=0).
        if "cam_full" not in self.tree["columns"]:
            self.tree["columns"] = self.tree["columns"] + ("cam_full",)
            self.tree.column("cam_full", width=0, stretch=False)
        
        # Limpa todos os itens da Treeview e as listas internas de controle.
        self.tree.delete(*self.tree.get_children())
        self.lista_arquivos.clear()
        self.checked.clear()

        # Lista todos os arquivos no diretório de entregas (incluindo subpastas).
        todos_root = listar_arquivos_no_diretorio(self.pasta_entregas)
        tipo = self.tipo_var.get() # Obtém o tipo de entrega selecionado no combobox (AP ou PE)

        todos = []
        # Filtra os arquivos para mostrar apenas os pertencentes ao tipo de entrega selecionado.
        for rv, nome, tam, cam, dt_mod in todos_root:
            caminho_norm = cam.replace("\\", "/").lower() # Normaliza o caminho para comparação
            # Verifica se o caminho contém a subpasta correspondente ao tipo de entrega (ex: /ap/ ou /pe/)
            if f"/{tipo.lower()}/" in caminho_norm:
                todos.append((rv, nome, tam, cam, dt_mod))
            else:
                # Trata casos onde a pasta de nível superior é "1.AP" ou "2.PE"
                pai = os.path.basename(os.path.dirname(cam)).lower()
                if tipo == "AP" and pai.startswith("1.ap"):
                    todos.append((rv, nome, tam, cam, dt_mod))
                elif tipo == "PE" and pai.startswith("2.pe"):
                    todos.append((rv, nome, tam, cam, dt_mod))

        if not todos:
            # Se nenhum arquivo for encontrado, exibe uma mensagem informativa.
            messagebox.showinfo(
                "Info",
                "Nenhum arquivo válido foi encontrado nas pastas de entrega."
            )
            return

        # Identifica arquivos obsoletos (ex: revisões antigas de um mesmo arquivo)
        obsoletos = set(identificar_obsoletos_custom(todos))
        # Filtra para manter apenas os arquivos "válidos" (não obsoletos)
        validos = [t for t in todos if t not in obsoletos]

        # Adiciona os arquivos válidos à Treeview, ordenados pelo nome do arquivo.
        for rv, nome, tam, cam, dt_mod in sorted(validos, key=lambda x: x[1].lower()):
            tam_kb = tam // 1024 # Converte tamanho de bytes para KB
            self.lista_arquivos.append((rv, nome, tam, cam, dt_mod)) # Armazena na lista interna
            # Insere o item na Treeview com um checkbox vazio (quadrado vazio unicode)
            iid = self.tree.insert("", tk.END, values=("\u2610", nome, dt_mod, tam_kb))
            self.checked[iid] = False # Inicializa o estado do checkbox como não marcado
            # Armazena o caminho completo do arquivo no item da Treeview para uso posterior
            self.tree.set(iid, "cam_full", cam)

    def _carregar_lista_inicial(self):
        """
        Preenche a Treeview usando self.lista_inicial, que é recebida do botão 'Voltar'
        da próxima tela (TelaAdicaoArquivos), permitindo persistir a seleção.
        """
        if not self.lista_inicial: # Se não houver lista inicial, não faz nada
            return
        
        # Garante a existência da coluna oculta 'cam_full'
        if "cam_full" not in self.tree["columns"]:
            self.tree["columns"] = self.tree["columns"] + ("cam_full",)
            self.tree.column("cam_full", width=0, stretch=False)

        self.tree.delete(*self.tree.get_children()) # Limpa a Treeview
        self.checked.clear() # Limpa os estados dos checkboxes
        
        for rv, nome, tam, cam, dt in self.lista_inicial:
            # Insere os arquivos na Treeview, similar a _carregar_entrega.
            iid = self.tree.insert("", tk.END,
                                   values=("\u2610", nome, dt, tam // 1024))
            self.tree.set(iid, "cam_full", cam)
            self.checked[iid] = False
        
        # Mantém a lista interna sincronizada com a lista inicial fornecida.
        self.lista_arquivos = self.lista_inicial.copy()

    def _toggle_checkbox(self, event):
        # Alterna o estado do checkbox (\u2610 = vazio, \u2611 = marcado) de um item na Treeview
        # quando o usuário clica na coluna "Sel".
        iid = self.tree.identify_row(event.y) # Identifica o item (linha) clicado
        col = self.tree.identify_column(event.x) # Identifica a coluna clicada
        
        if col != "#1" or not iid: # Só age se a coluna for a primeira ('#1') e um item válido for clicado
            return
        
        new_state = not self.checked.get(iid, False) # Inverte o estado atual do checkbox
        self.checked[iid] = new_state # Atualiza o dicionário de controle
        vals = list(self.tree.item(iid, "values")) # Obtém os valores atuais do item
        vals[0] = "\u2611" if new_state else "\u2610" # Atualiza o caractere do checkbox
        self.tree.item(iid, values=vals) # Atualiza o item na Treeview
        return "break" # Impede que o evento se propague para outros manipuladores

    # util simples para pegar revisao do nome (usa regex já presente no código original)
    @staticmethod
    def _identificar_rev(nome):
        # Função utilitária estática para extrair a revisão e extensão de um nome de arquivo.
        # Reutiliza a função 'identificar_nome_com_revisao' importada.
        nb, rev, ex = identificar_nome_com_revisao(nome)
        return rev, ex

    # --------------------------- edição ------------------------------
    def _iniciar_edicao_nome(self, event):
        # Permite ao usuário editar o nome de um arquivo diretamente na Treeview.
        iid = self.tree.identify_row(event.y) # Identifica o item clicado
        if not iid: # Se nenhum item, sai
            return
        col = self.tree.identify_column(event.x)
        if col != "#2":  # só permite editar a coluna 'nome' (coluna #2)
            return
        
        # Obtém as coordenadas e dimensões da célula clicada
        x, y, w, h = self.tree.bbox(iid, col)
        valor_antigo = self.tree.set(iid, "nome") # Obtém o nome atual do arquivo

        entry = tk.Entry(self.tree) # Cria um widget Entry (campo de texto)
        entry.place(x=x, y=y, width=w, height=h) # Posiciona o Entry sobre a célula da Treeview
        entry.insert(0, valor_antigo) # Preenche o Entry com o nome atual
        entry.focus_set() # Define o foco no Entry para o usuário digitar

        def _salvar(e):
            # Função interna para salvar o novo nome. Chamada ao pressionar Enter ou perder o foco.
            novo_nome = entry.get().strip() # Obtém o texto do Entry e remove espaços extras
            entry.destroy() # Destrói o Entry (removendo-o da interface)
            if not novo_nome or novo_nome == valor_antigo: # Se o nome não mudou ou está vazio, não faz nada
                return
            
            # Renomeia o arquivo no sistema de arquivos
            idx = self.tree.index(iid) # Obtém o índice do item na lista interna
            cam_antigo = self.lista_arquivos[idx][3] # Pega o caminho completo antigo
            novo_cam = os.path.join(os.path.dirname(cam_antigo), novo_nome) # Constrói o novo caminho completo
            try:
                os.rename(cam_antigo, novo_cam) # Tenta renomear o arquivo
            except OSError as err: # Captura erros do sistema operacional (permissão, arquivo em uso, etc.)
                messagebox.showerror("Erro", f"Falha ao renomear arquivo:\n{err}")
                return
            
            # Atualiza as estruturas de dados internas e a Treeview
            tam, dtmod = self.lista_arquivos[idx][2], self.lista_arquivos[idx][4] # Mantém tamanho e data de modificação
            self.lista_arquivos[idx] = (self.lista_arquivos[idx][0], novo_nome, tam, novo_cam, dtmod) # Atualiza a lista interna
            self.tree.set(iid, "nome", novo_nome) # Atualiza o nome na Treeview
        
        entry.bind("<Return>", _salvar) # Salva ao pressionar Enter
        entry.bind("<FocusOut>", _salvar) # Salva ao perder o foco

    # --------------------------- exclusão ----------------------------
    def _excluir_selecionados(self):
        # Move os arquivos selecionados para a lixeira (ou os exclui permanentemente se send2trash não estiver disponível).
        logger.debug("Visualizacao excluir selecionados clicked")
        marked = [iid for iid, val in self.checked.items() if val] # Obtém IDs dos itens marcados
        if not marked: # Se nenhum item marcado, sai
            return
        
        # Pede confirmação ao usuário
        if not messagebox.askyesno(
            "Confirmação",
            "Tem certeza de que deseja excluir os arquivos selecionados?\nEles serão enviados à lixeira.",
        ):
            return # Se não confirmar, sai
        
        erros = [] # Lista para armazenar erros de exclusão
        # Itera sobre os itens marcados em ordem reversa (para evitar problemas de índice ao remover da Treeview)
        for iid in sorted(marked, key=self.tree.index, reverse=True):
            idx = self.tree.index(iid) # Obtém o índice na lista interna
            _, nome, _, cam_full, _ = self.lista_arquivos[idx] # Pega informações do arquivo
            cam_full = os.path.normpath(cam_full) # Normaliza o caminho

            try:
                if send2trash: # Se send2trash estiver disponível
                    try:
                        send2trash(cam_full) # Tenta enviar para a lixeira
                    except OSError: # Se send2trash falhar (ex: sistema de arquivos diferente), remove diretamente
                        os.remove(cam_full)
                else: # Se send2trash não estiver disponível, remove diretamente
                    os.remove(cam_full)
            except OSError as err: # Captura erros de exclusão
                erros.append(str(err)) # Adiciona o erro à lista
                continue # Continua para o próximo arquivo
            
            # Se a exclusão foi bem-sucedida, remove da Treeview e das listas internas.
            self.tree.delete(iid)
            self.lista_arquivos.pop(idx)
            self.checked.pop(iid, None)
        
        if erros: # Se houve erros, exibe um aviso
            messagebox.showwarning(
                "Aviso",
                "Alguns arquivos não puderam ser excluídos:\n" + "\n".join(erros),
            )
        logger.debug("Visualizacao excluir selecionados executed successfully")

    # --------------------------- navegação ---------------------------
    def _avancar(self):
        # Avança para a próxima etapa da entrega (TelaAdicaoArquivos).
        logger.debug("Visualizacao avancar clicked")
        lista_init = []
        # Prepara uma lista dos arquivos selecionados (marcados) para passar para a próxima tela.
        for iid in self.tree.get_children():
            if self.checked.get(iid): # Se o checkbox do item estiver marcado
                idx = self.tree.index(iid)
                rv, nome, tam, cam, dt = self.lista_arquivos[idx]
                lista_init.append((rv, nome, tam, cam, dt))
        
        self.destroy() # Fecha a janela atual
        # Instancia e inicia a próxima tela, passando os arquivos selecionados.
        TelaAdicaoArquivos(lista_inicial=lista_init, pasta_entrega=self.pasta_entregas,
                           numero_projeto=self.projeto_num).mainloop()
        logger.debug("Visualizacao avancar executed successfully")

    def _voltar(self):
        # Volta para a tela de seleção de disciplina.
        logger.debug("Visualizacao voltar clicked")
        # Calcula o caminho do projeto a partir da pasta de entregas (3 níveis acima: 1.ENTREGAS -> Disciplina -> 3 Desenvolvimento -> Projeto).
        caminho_proj = os.path.dirname(os.path.dirname(
                         os.path.dirname(self.pasta_entregas)))

        self.withdraw() # Esconde a janela atual temporariamente
        # Chama a janela de seleção de disciplina
        nova_pasta, _ = janela_selecao_disciplina(self.projeto_num, caminho_proj)

        if not nova_pasta:  # Se o usuário cancelou a seleção de disciplina
            self.deiconify() # Mostra a janela atual novamente
            return

        # Se uma nova disciplina foi selecionada: fecha a atual e abre novamente com a nova pasta
        self.destroy() # Destrói a janela atual
        TelaVisualizacaoEntregaAnterior(
            pasta_entregas=nova_pasta,
            projeto_num=self.projeto_num,
            disciplina=os.path.basename(os.path.dirname(nova_pasta)) # Pega o nome da disciplina a partir do caminho
        ).mainloop() # Inicia a nova instância da tela de visualização
        logger.debug("Visualizacao voltar executed successfully")

# -----------------------------------------------------
# Janela 1: TelaAdicaoArquivos
# -----------------------------------------------------
class TelaAdicaoArquivos(tk.Tk):    
    """
    Janela inicial: o usuário seleciona os arquivos para entrega.
    Esta classe gerencia a interface e lógica para adicionar arquivos a uma nova entrega.
    """
    def __init__(self, lista_inicial=None, pasta_entrega: str | None = None, numero_projeto: str | None = None, *args, **kwargs):
        super().__init__(*args, **kwargs) # Chama o construtor da classe base Tkinter (tk.Tk)

        self.pasta_entrega   = pasta_entrega # Caminho da pasta de entregas (1.ENTREGAS)
        self.numero_projeto  = numero_projeto # Número do projeto
        # Obtém o nome da disciplina a partir do caminho da pasta_entrega
        self.disciplina      = (os.path.basename(os.path.dirname(pasta_entrega)) if pasta_entrega else "")
        
        root_frame = tk.Frame(self) # Frame raiz que contém todo o conteúdo da janela
        root_frame.pack(fill=tk.BOTH, expand=True) # Empacota para preencher toda a janela

        # Barra lateral esquerda (estilizada com cores escuras)
        bar_l = tk.Frame(root_frame, bg="#2c3e50", width=200)
        bar_l.pack(side=tk.LEFT, fill=tk.Y) # Empacota à esquerda, preenchendo verticalmente
        bar_l.pack_propagate(False) # Impede que a barra lateral se redimensione automaticamente

        # Rótulo do cabeçalho na barra lateral
        tk.Label(bar_l, text="OAE - Engenharia",
                 font=("Helvetica", 14, "bold"),
                 bg="#2c3e50", fg="white").pack(pady=10)

        # Rótulo da seção "PROJETOS" na barra lateral
        tk.Label(bar_l, text="PROJETOS",
                 font=("Helvetica", 10, "bold"),
                 bg="#34495e", fg="white", anchor="w", padx=10
                 ).pack(fill=tk.X, pady=(0, 5))

        # Listbox para exibir informações do projeto (atualmente apenas o número do projeto)
        lst_proj = tk.Listbox(bar_l, height=5,
                              bg="#ecf0f1", font=("Helvetica", 9))
        lst_proj.pack(fill=tk.X, padx=10, pady=5)
        if numero_projeto:
            lst_proj.insert(tk.END, f"Projeto {numero_projeto}")

        # Rótulo da seção "MEMBROS" na barra lateral (placeholder)
        tk.Label(bar_l, text="MEMBROS",
                 font=("Helvetica", 10, "bold"),
                 bg="#34495e", fg="white", anchor="w", padx=10
                 ).pack(fill=tk.X, pady=5)

        # Painel onde ficará TODO o conteúdo principal (direita da barra lateral)
        content = tk.Frame(root_frame, bg="#f5f5f5")
        content.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Título da janela principal
        lbl_title = tk.Label(content, text="Adicionar Arquivos para Entrega",
                             font=("Helvetica", 15, "bold"), bg="#f5f5f5",
                             anchor="w")
        lbl_title.pack(fill=tk.X, pady=(10, 5), padx=10)

        # Configurações da janela principal
        self.resizable(True, True) # Permite redimensionar a janela
        self.title("Adicionar Arquivos para Entrega") # Título da janela

        # Calcula o tamanho inicial da janela com base na resolução da tela
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        new_height = int(screen_h * 0.85)
        new_width = 1100
        self.geometry(f"{new_width}x{new_height}") # Define o tamanho da janela

        self.view_mode = "grouped" # Modo de exibição inicial dos arquivos (agrupado por tipo)

        container = tk.Frame(content) # Frame para o conteúdo principal dos arquivos
        container.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        lbl = tk.Label(container, text="Selecione arquivos para entrega. A listagem aparecerá conforme forem adicionados.")
        lbl.pack(pady=5)

        btn_frame = tk.Frame(container) # Frame para os botões de ação
        btn_frame.pack(fill=tk.X, pady=5)

        # Botões de ação para adicionar/remover arquivos e alternar visualização
        tk.Button(btn_frame, text="Adicionar arquivos", command=self.adicionar_arquivos).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Remover arquivos selecionados", command=self.remover_selecionados).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="Visualizar tudo", command=self.ativar_visualizar_tudo).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Agrupar por tipo", command=self.ativar_agrupado).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="Analisar", width=18, command=self.proxima_janela_nomenclatura).pack(side=tk.RIGHT, padx=5)

        self.canvas_global = tk.Canvas(content) # Canvas para permitir rolagem do conteúdo principal
        self.canvas_global.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar_global = tk.Scrollbar(container, orient="vertical", command=self.canvas_global.yview, width=20) # Barra de rolagem do canvas
        self.scrollbar_global.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas_global.configure(yscrollcommand=self.scrollbar_global.set) # Liga a barra de rolagem ao canvas

        self.inner_frame = tk.Frame(self.canvas_global) # Frame interno que conterá as tabelas de arquivos
        self.inner_frame_id = self.canvas_global.create_window((0,0), window=self.inner_frame, anchor="nw") # Cria uma janela no canvas para o frame interno

        # Liga eventos de redimensionamento para ajustar o canvas e a área de rolagem
        self.canvas_global.bind("<Configure>", self.on_canvas_configure)
        self.inner_frame.bind("<Configure>", self.on_frame_configure)

        bottom = tk.Frame(content) # Frame para os botões do rodapé
        # ancorado no rodapé:
        bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 10))

        tk.Button(bottom, text="Cancelar",
                  command=self._cancelar).pack(side=tk.LEFT, padx=5)

        tk.Button(bottom, text="Voltar",
                  command=self._voltar).pack(side=tk.RIGHT, padx=5)

        self.arquivos_por_grupo = {} # Dicionário para armazenar arquivos agrupados por tipo/extensão

        self.paned = None # Variável para o widget Panedwindow (usado em modo agrupado)
        self.table_all = None # Variável para a Treeview no modo "visualizar tudo"

        # Definição das colunas e seus títulos para as tabelas de arquivos
        self.colunas = ("num_arq", "nome_arq", "revisao", "dt_mod", "ext")
        self.coluna_titulos = {
            "num_arq": "Número do arquivo",
            "nome_arq": "Nome do arquivo",
            "revisao": "Revisão",
            "dt_mod": "Data de modificação",
            "ext": "Extensão",
        }
        self.tables = {} # Dicionário para armazenar as Treeviews para cada grupo de arquivos

        if lista_inicial: # Se uma lista inicial de arquivos for fornecida (ao voltar da próxima tela)
            self.recarregar_lista(lista_inicial) # Recarrega esses arquivos

        self.render_view() # Renderiza a visualização inicial dos arquivos

        if pasta_entrega and (not lista_inicial):
            # Se a pasta de entrega foi definida e não há lista inicial (primeira vez na tela),
            # abre o diálogo de seleção de arquivos automaticamente.
            self._abrir_filedialog_inicial(pasta_entrega)

    def _abrir_filedialog_inicial(self, pasta_entrega):
        # Abre o diálogo de seleção de arquivos automaticamente na inicialização da tela,
        # usando o último diretório salvo ou um diretório padrão.
        init_dir = carregar_ultimo_diretorio() # Tenta carregar o último diretório usado
        if not init_dir or not os.path.exists(init_dir):
            # Se não houver último diretório ou ele não existir, usa a pasta de entrega ou o diretório home.
            init_dir = pasta_entrega if os.path.isdir(pasta_entrega) else os.path.expanduser("~")
        
        paths = filedialog.askopenfilenames( # Abre o diálogo para múltiplos arquivos
            title="Selecione arquivos para entrega",
            initialdir=init_dir
        )
        if not paths: # Se o usuário cancelar, sai
            return
        
        dir_of_first = os.path.dirname(paths[0]) # Pega o diretório do primeiro arquivo selecionado
        salvar_ultimo_diretorio(dir_of_first) # Salva este diretório como o último usado

        for p in paths:
            if not os.path.isfile(p): # Ignora se não for um arquivo
                continue
            
            # Extrai informações do arquivo
            base, rev, ext = identificar_nome_com_revisao(os.path.basename(p))
            data_ts = os.path.getmtime(p) # Timestamp da última modificação
            data_mod = datetime.datetime.fromtimestamp(data_ts).strftime("%d/%m/%Y %H:%M") # Formata a data
            no_arq = extrair_numero_arquivo(base) # Extrai número do arquivo
            grupo = self.classificar_extensao(ext) # Classifica a extensão em um grupo (PDF, CAD, Outros)
            
            # Adiciona o arquivo ao dicionário agrupado por tipo, se ainda não estiver lá
            self.arquivos_por_grupo.setdefault(grupo, [])
            if p not in [x[0] for x in self.arquivos_por_grupo[grupo]]: # Evita duplicatas pelo caminho completo
                self.arquivos_por_grupo[grupo].append((p, no_arq, base, rev, data_mod, ext))

        self.render_view() # Re-renderiza a visualização para mostrar os novos arquivos

    def recarregar_lista(self, lista_inicial):
        """
        Popula o dicionário `arquivos_por_grupo` com base em uma `lista_inicial`
        (usado ao voltar da tela de visualização de entrega anterior).
        """
        for (rv, arq, tam, path, dmod) in lista_inicial:
            # Reconstrói as informações do arquivo a partir da lista inicial
            base, revisao, ext = identificar_nome_com_revisao(arq)
            no_arq = extrair_numero_arquivo(base)
            grupo = self.classificar_extensao(ext)
            
            # Adiciona o arquivo ao grupo correspondente
            if grupo not in self.arquivos_por_grupo:
                self.arquivos_por_grupo[grupo] = []
            self.arquivos_por_grupo[grupo].append((path, no_arq, base, revisao, dmod, ext))
        self.render_view() # Re-renderiza a visualização

    def on_canvas_configure(self, event):
        # Callback para ajustar a largura do frame interno quando o canvas é redimensionado.
        self.canvas_global.itemconfig(self.inner_frame_id, width=event.width)

    def on_frame_configure(self, event):
        # Callback para ajustar a região de rolagem do canvas quando o frame interno muda de tamanho.
        self.canvas_global.configure(scrollregion=self.canvas_global.bbox("all"))

    def _cancelar(self):
        # Cancela a operação e fecha a aplicação.
        logger.debug("Adicao arquivos cancelar clicked")
        self.destroy() # Fecha a janela
        logger.debug("Adicao arquivos cancelar executed successfully")
        sys.exit(0) # Sai do programa

    def _voltar(self):
        # Volta para a tela de visualização da entrega anterior, passando os arquivos atualmente adicionados.
        logger.debug("Adicao arquivos voltar clicked")
        lista_atual = []
        # Coleta todos os arquivos que foram adicionados (através do filedialog ou recarregados)
        for grupo, lst in self.arquivos_por_grupo.items():
            for (path, no_arq, base, rev, data_mod, ext) in lst:
                tam = os.path.getsize(path) # Obtém o tamanho atual do arquivo
                arq = os.path.basename(path) # Obtém o nome base do arquivo
                lista_atual.append((rev, arq, tam, path, data_mod)) # Formato esperado pela tela anterior

        self.destroy() # Fecha a janela atual

        # Reinstancia e exibe a tela de visualização de entrega anterior, passando a lista atual de arquivos.
        TelaVisualizacaoEntregaAnterior(
            pasta_entregas=self.pasta_entrega,
            projeto_num=self.numero_projeto,
            disciplina=self.disciplina,
            lista_inicial=lista_atual # Passa os arquivos para a tela anterior
        ).mainloop()
        logger.debug("Adicao arquivos voltar executed successfully")

    def adicionar_arquivos(self):
        # Abre um diálogo de seleção de arquivos para o usuário adicionar mais arquivos à lista.
        logger.debug("Adicao arquivos adicionar clicked")
        init_dir = carregar_ultimo_diretorio() # Carrega o último diretório usado
        if not init_dir or not os.path.exists(init_dir):
            init_dir = os.path.expanduser("~") # Se não houver, usa o diretório home

        paths = filedialog.askopenfilenames(title="Selecione arquivos para entrega", initialdir=init_dir)
        if not paths: # Se o usuário cancelar, sai
            return
        
        dir_of_first = os.path.dirname(paths[0])
        salvar_ultimo_diretorio(dir_of_first) # Salva o diretório do primeiro arquivo selecionado

        for p in paths:
            if not os.path.isfile(p): # Ignora se não for um arquivo válido
                continue
            
            # Extrai informações do arquivo, similar a _abrir_filedialog_inicial
            base, rev, ext = identificar_nome_com_revisao(os.path.basename(p))
            data_ts = os.path.getmtime(p)
            data_mod = datetime.datetime.fromtimestamp(data_ts).strftime("%d/%m/%Y %H:%M")
            no_arq = extrair_numero_arquivo(base)
            grupo = self.classificar_extensao(ext)

            # Adiciona o arquivo ao grupo, evitando duplicatas
            if grupo not in self.arquivos_por_grupo:
                self.arquivos_por_grupo[grupo] = []
            if p not in [x[0] for x in self.arquivos_por_grupo[grupo]]:  # Verifica duplicatas pelo caminho completo
                self.arquivos_por_grupo[grupo].append((p, no_arq, base, rev, data_mod, ext))

        self.render_view() # Re-renderiza a visualização com os novos arquivos
        logger.debug("Adicao arquivos adicionar executed successfully")

    def classificar_extensao(self, ext):
        # Classifica a extensão de um arquivo em um grupo predefinido (PDF, CAD, Outros).
        # Utiliza o dicionário GRUPOS_EXT importado.
        if ext == ".pdf": # PDF é um caso especial
            return "PDF"
        for grupo, exts in GRUPOS_EXT.items(): # Itera sobre os grupos definidos em GRUPOS_EXT
            if ext in exts:
                return grupo # Retorna o nome do grupo se a extensão for encontrada
        return "Outros" # Retorna "Outros" se a extensão não se encaixar em nenhum grupo

    def ativar_visualizar_tudo(self):
        # Altera o modo de visualização para "todos os arquivos" e re-renderiza.
        logger.debug("Adicao arquivos visualizar tudo clicked")
        self.view_mode = "all"
        self.render_view()
        logger.debug("Adicao arquivos visualizar tudo executed successfully")

    def ativar_agrupado(self):
        # Altera o modo de visualização para "agrupado por tipo" e re-renderiza.
        logger.debug("Adicao arquivos agrupado clicked")
        self.view_mode = "grouped"
        self.render_view()
        logger.debug("Adicao arquivos agrupado executed successfully")
        
    def render_view(self):
        # Destroi todos os widgets filhos do frame interno para limpar a visualização.
        for child in self.inner_frame.winfo_children():
            child.destroy()
        # Limpa as referências às tabelas e panedwindow.
        self.tables.clear()
        self.table_all = None
        self.paned = None

        # Renderiza a visualização com base no modo selecionado ("grouped" ou "all").
        if self.view_mode == "grouped":
            self.render_tables_grouped()
        else:
            self.render_table_all()

        # Atualiza o canvas para ajustar a região de rolagem ao novo conteúdo.
        self.inner_frame.update_idletasks()
        self.canvas_global.config(scrollregion=self.canvas_global.bbox("all"))

    def render_tables_grouped(self):
        """
        Cria uma Treeview por grupo de extensão.
        • Se houver só 1 grupo → coloca o quadro direto no inner_frame
          (evita colapso do Panedwindow).
        • Se houver 2+ grupos → mantém Panedwindow como antes.
        """
        # Filtra os grupos para incluir apenas aqueles que contêm arquivos.
        grupos = [(g, lst) for g, lst in self.arquivos_por_grupo.items() if lst]

        # ---------- CASO 1 GRUPO ----------------------------------
        # Se houver apenas um grupo, renderiza-o diretamente no frame interno.
        if len(grupos) == 1:
            grupo, lista = grupos[0]
            self._criar_quadro_grupo(self.inner_frame, grupo, lista)
            return  # Pronto, não precisa de Panedwindow.

        # ---------- CASO 2+ GRUPOS (comportamento original) -------
        # Cria uma Panedwindow vertical para organizar múltiplos grupos.
        self.paned = ttk.Panedwindow(self.inner_frame, orient="vertical")
        self.paned.pack(fill=tk.BOTH, expand=True)

        # Configura o estilo para as Treeviews (altura da linha).
        ttk.Style().configure("Treeview", rowheight=24)

        # Itera sobre cada grupo e cria um quadro com Treeview para ele.
        for grupo, lista in grupos:
            quadro = tk.Frame(self.paned)
            self._criar_quadro_grupo(quadro, grupo, lista)
            # Tenta adicionar o quadro à Panedwindow com minsize, com fallback se não suportado.
            try:
                self.paned.add(quadro, weight=1, minsize=120)
            except tk.TclError:
                self.paned.add(quadro, weight=1)

    # --------------------------------------------------------------
    # Função auxiliar reaproveitada pelos dois caminhos
    # --------------------------------------------------------------
    def _criar_quadro_grupo(self, parent, grupo, lista_arqs):
        """Monta label + Treeview para um grupo de arquivos."""
        # Cria um rótulo para o nome do grupo.
        tk.Label(parent, text=grupo, font=("Arial", 11, "bold")
                 ).pack(anchor="w", pady=5)

        # Cria um sub-frame para conter a Treeview e a barra de rolagem.
        sub = tk.Frame(parent); sub.pack(fill=tk.BOTH, expand=True)
        sub.config(height=120); sub.pack_propagate(False) # Define altura fixa e impede que sub redimensione com os filhos.

        # Cria a barra de rolagem vertical.
        sb = tk.Scrollbar(sub, orient="vertical", width=20)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Cria a Treeview e a associa à barra de rolagem.
        tree = ttk.Treeview(
            sub, columns=self.colunas, show="headings",
            yscrollcommand=sb.set
        )
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.config(command=tree.yview)

        # Configura os cabeçalhos das colunas e suas propriedades.
        for col in self.colunas:
            tree.heading(col, text=self.coluna_titulos[col],
                         command=lambda c=col, t=tree:
                             self.sort_column(t, c, False))
            tree.column(col, width=120, anchor='w', stretch=True)
        # Define uma largura específica para a coluna "nome_arq".
        tree.column("nome_arq", width=220)

        # Insere os dados dos arquivos na Treeview.
        for full, num, base, rev, dmod, ext in lista_arqs:
            tree.insert("", tk.END, values=(num, base, rev, dmod, ext))

        # Guarda uma referência para a Treeview para futuras manipulações (remoção/ordenação).
        self.tables[grupo] = tree

    def render_table_all(self):
        # Agrega todos os arquivos de todos os grupos em uma única lista.
        all_files = []
        for grupo, lista_arqs in self.arquivos_por_grupo.items():
            all_files.extend(lista_arqs)

        # Cria um frame para conter a Treeview e a barra de rolagem.
        frame_all = tk.Frame(self.inner_frame)
        frame_all.pack(fill=tk.BOTH, expand=True)

        # Cria a barra de rolagem vertical para a Treeview principal.
        scrollbar_all = tk.Scrollbar(frame_all, orient="vertical", width=20)
        scrollbar_all.pack(side=tk.RIGHT, fill=tk.Y)

        # Cria a Treeview principal para exibir todos os arquivos.
        self.table_all = ttk.Treeview(
            frame_all,
            columns=self.colunas,
            show="headings",
            yscrollcommand=scrollbar_all.set
        )
        self.table_all.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_all.config(command=self.table_all.yview)

        # Configura o estilo da Treeview (altura da linha).
        style = ttk.Style()
        style.configure("Treeview", rowheight=24)

        # Configura os cabeçalhos das colunas e suas propriedades, incluindo o comando de ordenação.
        for col in self.colunas:
            self.table_all.heading(
                col,
                text=self.coluna_titulos[col],
                command=lambda c=col: self.sort_column(self.table_all, c, False)
            )
            self.table_all.column(col, width=120, anchor='w', stretch=True)

        # Define uma largura específica para a coluna "nome_arq".
        self.table_all.column("nome_arq", width=220)

        # Insere os dados de todos os arquivos na Treeview.
        for (fullpath, no_arq, base, rev, data_mod, ext) in all_files:
            self.table_all.insert("", tk.END, values=(no_arq, base, rev, data_mod, ext))

    def sort_column(self, tree, col, reverse):
        # Função auxiliar para tentar converter valores para float para ordenação numérica.
        # Se não for numérico, retorna uma tupla com um indicador de tipo e o valor como string.
        def try_convert(v):
            try:
                return (0, float(v))  # (0, valor_numerico) para ordenação numérica
            except ValueError:
                return (1, str(v))    # (1, valor_string) para ordenação alfabética

        # Coleta os valores da coluna e os IDs dos itens da Treeview.
        l = []
        for k in tree.get_children(""):
            val = tree.set(k, col)
            l.append((try_convert(val), k))
        
        # Ordena a lista com base nos valores convertidos.
        l.sort(key=lambda x: x[0], reverse=reverse)

        # Move os itens na Treeview para refletir a nova ordem.
        for index, (_, iid) in enumerate(l):
            tree.move(iid, '', index)
        
        # Atualiza o comando do cabeçalho da coluna para alternar a ordem na próxima vez que for clicado.
        tree.heading(col, command=lambda: self.sort_column(tree, col, not reverse))

    def remover_selecionados(self):
        logger.debug("Adicao arquivos remover selecionados clicked")
        # Verifica o modo de visualização para remover os arquivos corretamente.
        if self.view_mode == "grouped":
            # Itera sobre cada Treeview nos grupos.
            for grupo, tree in self.tables.items():
                sel = tree.selection()
                if not sel:
                    continue  # Se não houver seleção, pula para o próximo grupo.
                
                # Itera sobre os itens selecionados (em ordem inversa para evitar problemas de índice ao deletar).
                for iid in reversed(sel):
                    vals = tree.item(iid,"values")
                    no_arq, base, rev, data_mod, ext = vals
                    idx_rm = None
                    # Encontra o arquivo correspondente na estrutura de dados interna.
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
                    # Se encontrado, remove o arquivo da lista interna e da Treeview.
                    if idx_rm is not None:
                        self.arquivos_por_grupo[grupo].pop(idx_rm)
                    tree.delete(iid)
        else: # Modo "all"
            if self.table_all:
                sel = self.table_all.selection()
                if not sel:
                    return # Se não houver seleção, retorna.
                
                # Itera sobre os itens selecionados (em ordem inversa).
                for iid in reversed(sel):
                    vals = self.table_all.item(iid,"values")
                    no_arq, base, rev, data_mod, ext = vals
                    self.table_all.delete(iid) # Remove o item da Treeview principal.
                    
                    # Procura e remove o arquivo da estrutura de dados interna em qualquer grupo.
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
                            break # Encontrou e removeu, pode parar de procurar em outros grupos.
        logger.debug("Adicao arquivos remover selecionados executed successfully")        

    def proxima_janela_nomenclatura(self):
        logger.debug("Adicao arquivos avancar clicked")
        final_list = []
        # Coleta todos os arquivos de todos os grupos para a lista final.
        for grupo, lista_arqs in self.arquivos_por_grupo.items():
            for (path, no_arq, base, rev, data_mod, ext) in lista_arqs:
                tam = os.path.getsize(path) # Obtém o tamanho do arquivo.
                arq = os.path.basename(path) # Obtém apenas o nome do arquivo.
                final_list.append((rev, arq, tam, path, data_mod))
        
        # Destrói a janela atual e abre a próxima janela de verificação de nomenclatura.
        self.destroy()
        tela = TelaVerificacaoNomenclatura(final_list)
        tela.mainloop()
        logger.debug("Adicao arquivos avancar executed successfully")


# -----------------------------------------------------
# Janela 2: TelaVerificacaoNomenclatura
# -----------------------------------------------------
class TelaVerificacaoNomenclatura(tk.Tk):
    def __init__(self, lista_arquivos, *args, **kwargs):
        super().__init__(*args, **kwargs)
        global NOMENCLATURA_GLOBAL, NUM_PROJETO_GLOBAL
        # Carrega a nomenclatura global do JSON se o número do projeto estiver definido.
        if NUM_PROJETO_GLOBAL:
            NOMENCLATURA_GLOBAL = carregar_nomenclatura_json(NUM_PROJETO_GLOBAL)
        self.lista_arquivos = lista_arquivos.copy()  # Copia a lista para evitar referência residual.

        # Mapas para armazenar tokens e tags pré-computados para cada arquivo.
        self._token_map = {}
        self._tags_map  = {}
        for rv, arq, tam, path, dmod in self.lista_arquivos:
            nome_sem_ext, _ = os.path.splitext(arq)
            # Divide o nome do arquivo em tokens e verifica a nomenclatura.
            tokens = split_including_separators(nome_sem_ext, NOMENCLATURA_GLOBAL)
            tags   = verificar_tokens(tokens, NOMENCLATURA_GLOBAL)
            self._token_map[arq] = tokens
            self._tags_map[arq]  = tags

        self.title("Verificação de Nomenclatura")
        self.resizable(True, True)
        self.lista_arquivos = lista_arquivos
        self.geometry("1200x700")

        # Cria o frame principal para organizar os widgets.
        container = tk.Frame(self)
        container.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Rótulo de instrução para o usuário.
        lbl = tk.Label(container, text="Confira a nomenclatura (campos e separadores). Caso haja erros, corrija antes de avançar.")
        lbl.pack(anchor="w", pady=5)

        # Frame para os botões de ação.
        frm_botoes = tk.Frame(container)
        frm_botoes.pack(fill=tk.X, pady=5)

        # Botões de "Mostrar Padrão", "Voltar" e "Avançar".
        tk.Button(frm_botoes, text="Mostrar Padrão", command=self.mostrar_nomenclatura_padrao).pack(side=tk.LEFT, padx=5)
        tk.Button(frm_botoes, text="Voltar", command=self.voltar).pack(side=tk.LEFT, padx=5)
        tk.Button(frm_botoes, text="Avançar", command=self.avancar).pack(side=tk.RIGHT, padx=5)

        # Legenda das cores para os status da nomenclatura.
        legend = (
            "Vermelho: token incorreto na nomenclatura. "
            "Amarelo: token faltando."
        )
        tk.Label(container, text=legend, font=("Arial", 10)).pack(anchor="w", pady=(0, 5))

        # Cria o Frame para a Treeview.
        frame_tv = tk.Frame(container)
        frame_tv.pack(fill=tk.BOTH, expand=True)

        # Cria a barra de rolagem vertical para a Treeview.
        self.tv_scroll_y = tk.Scrollbar(frame_tv, orient="vertical")
        self.tv_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # Cria a Treeview para exibir os tokens dos arquivos.
        self.tree = ttk.Treeview(frame_tv, show="headings", yscrollcommand=self.tv_scroll_y.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tv_scroll_y.config(command=self.tree.yview)

        # Definimos a quantidade de colunas = "máximo de tokens" que possamos ter
        # em qualquer arquivo. A cada arquivo, iremos tokenizar. Se um arquivo tiver 20 tokens,
        # e outro 24, tomamos o max.
        # quantidade de tokens esperados pela nomenclatura (campos + separadores)
        campos_cfg = NOMENCLATURA_GLOBAL.get("campos", []) if NOMENCLATURA_GLOBAL else []
        expected_tokens = 2 * len(campos_cfg) - 1 if campos_cfg else 0

        # também consideramos o maximo de tokens real dos arquivos lidos
        real_max = 0
        for (rv, arq, tam, path, dmod) in self.lista_arquivos:
            nome_sem_ext, _ = os.path.splitext(arq)
            tokens = split_including_separators(nome_sem_ext, NOMENCLATURA_GLOBAL)
            real_max = max(real_max, len(tokens))

        self.max_tokens = max(expected_tokens, real_max)

        # Cria as colunas da Treeview, nomeadas como T1, T2, etc.
        # A solicitação é para não ter a primeira coluna "Nome do arquivo",
        # então as colunas são apenas para os tokens.
        
        col_names = [f"T{i}" for i in range(1, self.max_tokens+1)]
        self.tree["columns"] = col_names
        for cn in col_names:
            self.tree.heading(cn, text="") # Cabeçalhos vazios.
            self.tree.column(cn, width=60, anchor="center")

        # Configuramos tags para colorir as linhas/células da Treeview.
        self.tree.tag_configure("mismatch", background="#FF9999")   # Vermelho clarinho para tokens incorretos.
        self.tree.tag_configure("missing", background="#FFFF99")    # Amarelo clarinho para tokens faltando.
        self.tree.tag_configure("ok", background="white")           # Fundo branco para tokens corretos.

        # Preenche a Treeview com os dados dos arquivos.
        self.preencher_arvore()

    def preencher_arvore(self):
        # Limpa todos os itens existentes na Treeview.
        self.tree.delete(*self.tree.get_children())

        # Itera sobre a lista de arquivos para preencher a Treeview.
        for (rv, arq, tam, path, dmod) in self.lista_arquivos:
            nome_sem_ext, _ = os.path.splitext(arq)
            # Recupera os tokens e tags de verificação pré-computados.
            tokens          = self._token_map[arq]
            tags_result     = self._tags_map[arq]
            
            # Monta a linha com valores e tags para cada célula, preenchendo até max_tokens.
            row_vals = []
            row_tags = []
            for i in range(self.max_tokens):
                if i < len(tokens):
                    row_vals.append(tokens[i])
                else:
                    row_vals.append("")  # Adiciona string vazia para tokens faltando.
                
                # Atribui a tag de cor para a célula.
                if i < len(tags_result):
                    if tags_result[i] == 'mismatch':
                        row_tags.append('mismatch')
                    elif tags_result[i] == 'missing':
                        row_tags.append('missing')
                    else:
                        row_tags.append('ok')
                else:
                    # Se sobrou espaço e não há mais tags de resultado, assume que é um token faltando.
                    row_tags.append('missing')

            # Determina a tag final da linha: "missing" > "mismatch" > "ok".
            if 'missing' in row_tags:
                final_tag = 'missing'
            elif 'mismatch' in row_tags:
                final_tag = 'mismatch'
            else:
                final_tag = 'ok'

            # Insere o item na Treeview com os valores da linha e a tag final.
            item_id = self.tree.insert("", tk.END, values=row_vals, tags=(final_tag,))

    def mostrar_nomenclatura_padrao(self):
        """Exibe uma janela com a nomenclatura padrão, cada campo + separador."""
        logger.debug("Nomenclatura mostrar padrao clicked")
        # Verifica se a nomenclatura global está definida.
        if not NOMENCLATURA_GLOBAL:
            messagebox.showinfo("Info", "Nomenclatura não definida para este projeto.")
            return

        # Cria uma nova janela Toplevel.
        win = tk.Toplevel(self)
        win.title("Nomenclatura Padrão")
        win.geometry("1000x300")

        campos = NOMENCLATURA_GLOBAL.get("campos", [])

        # Calcula o número de colunas necessárias para exibir campos e separadores.
        col_count = 2 * len(campos) - 1
        col_ids = [f"C{i}" for i in range(col_count)]
        
        # Cria um frame para a Treeview e suas barras de rolagem.
        frm_tree = tk.Frame(win)
        frm_tree.pack(fill=tk.BOTH, expand=True)

        # Cria as barras de rolagem horizontal e vertical.
        sbx = tk.Scrollbar(frm_tree, orient="horizontal")
        sby = tk.Scrollbar(frm_tree, orient="vertical")
        sby.pack(side=tk.RIGHT, fill=tk.Y)
        sbx.pack(side=tk.BOTTOM, fill=tk.X)

        # Cria a Treeview para exibir a nomenclatura padrão.
        tv = ttk.Treeview(frm_tree, columns=col_ids, show="headings",
                          xscrollcommand=sbx.set, yscrollcommand=sby.set)
        tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sbx.config(command=tv.xview)
        sby.config(command=tv.yview)

        # Configura os cabeçalhos e a largura das colunas.
        for cid in col_ids:
            tv.heading(cid, text="") # Cabeçalhos vazios.
            tv.column(cid, width=80, anchor="center")

        # Monta uma linha de exemplo com a nomenclatura padrão.
        row_vals = []
        for i, cinfo in enumerate(campos):
            # Exibe o primeiro valor fixo ou "(livre)".
            fixos = cinfo.get("valores_fixos",[])
            if fixos:
                if isinstance(fixos[0], dict):
                    row_vals.append(fixos[0].get("value",""))
                else:
                    row_vals.append(str(fixos[0]))
            else:
                row_vals.append("(livre)")

            # Adiciona o separador se não for o último campo.
            if i < len(campos)-1:
                sep_ = cinfo.get("separador","-")
                row_vals.append(sep_)

        # Insere a linha de exemplo na Treeview.
        tv.insert("", tk.END, values=row_vals)
        logger.debug("Nomenclatura mostrar padrao executed successfully")

    def voltar(self):
        logger.debug("Nomenclatura voltar clicked")
        # Destrói a janela atual.
        self.destroy()
        # Volta para a janela de adição de arquivos, passando os dados atuais.
        TelaAdicaoArquivos(
            lista_inicial=self.lista_arquivos,
            pasta_entrega=PASTA_ENTREGA_GLOBAL,
            numero_projeto=NUM_PROJETO_GLOBAL
        ).mainloop()
        logger.debug("Nomenclatura voltar executed successfully")

    def avancar(self):
        logger.debug("Nomenclatura avancar clicked")
        global TIPO_ENTREGA_GLOBAL
        # Verifica se há erros de nomenclatura (mismatch ou missing) antes de avançar.
        for iid in self.tree.get_children():
            tags_ = self.tree.item(iid, "tags")
            if "mismatch" in tags_ or "missing" in tags_:
                messagebox.showwarning(
                    "Atenção",
                    "Há campos com erro (vermelho) ou faltando (amarelo). "
                    "Corrija antes de avançar."
                )
                return # Impede o avanço se houver erros.

        # --- Escolha AP / PE ---
        # Chama uma função para o usuário escolher o tipo de entrega (AP ou PE).
        escolha = escolher_tipo_entrega(self)
        if escolha is None:        # Usuário cancelou a escolha.
            return
        TIPO_ENTREGA_GLOBAL = escolha # Armazena a escolha globalmente.
        
        # Destrói a janela atual e avança para a tela de verificação de revisão.
        self.destroy()
        TelaVerificacaoRevisao(self.lista_arquivos).mainloop()
        logger.debug("Nomenclatura avancar executed successfully")

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
        self.diretorio = PASTA_ENTREGA_GLOBAL # Obtém o diretório de entrega global.
        
        # Carrega dados de entregas anteriores e verifica se é a primeira entrega.
        self.dados_anteriores = carregar_dados_anteriores(self.diretorio)
        self.primeira_entrega = (len(self.dados_anteriores) == 0)
        
        # Analisa e classifica os arquivos (novos, revisados, alterados).
        (self.arquivos_novos,
         self.arquivos_revisados,
         self.arquivos_alterados) = analisar_comparando_estado(self.lista_arquivos, self.dados_anteriores)
        
        # Identifica arquivos obsoletos no diretório de entrega.
        todos_diretorio = listar_arquivos_no_diretorio(self.diretorio)
        self.obsoletos = identificar_obsoletos_custom(todos_diretorio)
        
        # Cria o frame principal para os widgets.
        container = tk.Frame(self)
        container.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # Exibe uma mensagem informativa sobre a análise da entrega.
        if self.primeira_entrega:
            info_text = "Essa é a primeira análise. Esses foram os arquivos:"
        else:
            info = obter_info_ultima_entrega(self.dados_anteriores)
            info_text = f"Esses foram os arquivos analisados a partir da última entrega, {info}."
        lbl = tk.Label(container, text=info_text, font=("Arial", 12, "bold"))
        lbl.pack(pady=5)
        
        # Cria as tabelas para arquivos novos e revisados.
        self.criar_tabela(container, "Arquivos novos", self.arquivos_novos)
        self.criar_tabela(container, "Arquivos revisados", self.arquivos_revisados)
        
        # Frame para os botões de ação.
        btn_frame = tk.Frame(container)
        btn_frame.pack(fill=tk.X, pady=5)
        
        # Botões "Voltar", "Confirmar" e "Cancelar".
        tk.Button(btn_frame, text="Voltar", command=self.voltar).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Confirmar", command=self.confirmar).pack(side=tk.RIGHT, padx=5)
        # O botão cancelar encerra a aplicação.
        tk.Button(btn_frame, text="Cancelar", command=lambda: (self.destroy(), sys.exit(0))).pack(side=tk.RIGHT, padx=5)

    def criar_tabela(self, parent, titulo, arr):
        # Cria um LabelFrame com o título fornecido para agrupar a tabela.
        lf = tk.LabelFrame(parent, text=titulo, font=("Arial", 11, "bold"))
        lf.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Cria a barra de rolagem vertical para a Treeview.
        scrollbar_local = tk.Scrollbar(lf, orient="vertical", width=20)
        scrollbar_local.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Define as colunas da Treeview.
        cols = ("Nome do arquivo","Revisão","Data de modificação")
        tree = ttk.Treeview(lf, columns=cols, show="headings", height=5, yscrollcommand=scrollbar_local.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_local.config(command=tree.yview)
        
        # Configura o estilo da Treeview (altura da linha).
        style = ttk.Style()
        style.configure("Treeview", rowheight=24)
        
        # Configura os cabeçalhos das colunas.
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=200, anchor='w', stretch=True)
        
        # Preenche a Treeview com os dados dos arquivos ou uma mensagem de "Nenhum".
        if not arr:
            tree.insert("", tk.END, values=("Nenhum","",""))
        else:
            for (rv,a,tam,cam,dmod) in arr:
                tree.insert("", tk.END, values=(a, rv if rv else "Sem Revisão", dmod))
        return tree

    def voltar(self):
        logger.debug("Revisao voltar clicked")
        # Destrói a janela atual.
        self.destroy()
        # Volta para a janela de verificação de nomenclatura.
        TelaVerificacaoNomenclatura(self.lista_arquivos).mainloop()
        logger.debug("Revisao voltar executed successfully")

    def confirmar(self):
        logger.debug("Revisao confirmar clicked")
        global TIPO_ENTREGA_GLOBAL
        # Se houver arquivos alterados (modificados com mesma revisão), exibe um erro.
        if self.arquivos_alterados:
            self.withdraw() # Esconde a janela principal temporariamente.
            janela_erro_revisao(self.arquivos_alterados) # Abre a janela de erro.
            self.deiconify() # Restaura a janela principal.
            return # Não prossegue se houver arquivos alterados não tratados.
        
        # Pede confirmação final ao usuário.
        if not messagebox.askyesno("Confirmação Final", "Confirma que estes arquivos estão corretos?"):
            self.destroy()
            sys.exit(0) # Se não confirmar, encerra a aplicação.
        
        # Pede confirmação se é uma entrega oficial.
        if not messagebox.askyesno("Entrega Oficial", "Essa é uma entrega oficial?"):
            self.destroy()
            sys.exit(0) # Se não for oficial, encerra a aplicação.
        
        try:
            # Cria a estrutura de pastas de entrega (AP/PE) se TIPO_ENTREGA_GLOBAL estiver definido.
            if TIPO_ENTREGA_GLOBAL:
                criar_pasta_entrega_ap_pe(
                    self.diretorio,
                    TIPO_ENTREGA_GLOBAL,
                    self.arquivos_novos + self.arquivos_revisados + self.arquivos_alterados
                )
            
            # Determina o subdiretório de destino (AP ou PE).
            subdir = "AP" if TIPO_ENTREGA_GLOBAL == "AP" else "PE"
            pasta_base = os.path.join(self.diretorio, subdir)
            
            # Lista as entregas ativas para encontrar a mais recente.
            entregas_ativas = [
                    d for d in os.listdir(pasta_base)
                if d.startswith(("1.AP - Entrega-", "2.PE - Entrega-"))
                and not d.endswith("-OBSOLETO")
            ]
            
            # Se houver entregas ativas, define a pasta de destino como a mais recente.
            if entregas_ativas:
                pasta_destino = os.path.join(
                    pasta_base,
                    max(
                        entregas_ativas,
                        key=lambda n: int(re.search(r"(\d+)$", n).group(1)), # Extrai o número da entrega para comparação.
                    ),
                )

                # Função auxiliar para redirecionar os caminhos dos arquivos para a pasta de destino.
                def _redir(lista):
                    for i, tup in enumerate(list(lista)):
                        nome = os.path.basename(tup[3]) # Obtém o nome base do arquivo.
                        lista[i] = tup[:3] + (
                            os.path.join(pasta_destino, nome), # Atualiza o caminho completo do arquivo.
                        ) + tup[4:]
                
                # Redireciona os caminhos para as listas de arquivos.
                _redir(self.arquivos_novos)
                _redir(self.arquivos_revisados)
                _redir(self.arquivos_alterados)

        except Exception as e:
            # Em caso de erro na criação/cópia da pasta, exibe uma mensagem de erro.
            messagebox.showerror(
                "Erro",
                f"Falha ao criar/copiar pasta de entrega AP/PE:\n{e}"
            )
            # A aplicação não é encerrada aqui, permitindo que o usuário possa tentar novamente ou cancelar.
            return # Retorna para não executar o pos_processamento em caso de erro.
        
        # Realiza o pós-processamento da entrega.
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
        logger.debug("Revisao confirmar executed successfully")
        sys.exit(0) # Encerra a aplicação após o sucesso.

def main():
    while True:  # Loop para seleção de projeto.
        num_proj, caminho_proj = janela_selecao_projeto()
        if not num_proj or not caminho_proj:
            return # Se o usuário cancelar a seleção do projeto, sai do programa.
        
        while True:  # Loop para seleção de disciplina dentro do projeto.
            pasta_entrega, voltar_proj = janela_selecao_disciplina(num_proj, caminho_proj)
            if voltar_proj:
                break  # Se o usuário quiser voltar para a seleção de projeto, sai do loop da disciplina.
            if not pasta_entrega:
                return  # Se o usuário cancelar a seleção da disciplina, sai do programa.

            global NOMENCLATURA_GLOBAL, PASTA_ENTREGA_GLOBAL, NUM_PROJETO_GLOBAL
            # Define as variáveis globais com base na seleção.
            NOMENCLATURA_GLOBAL = carregar_nomenclatura_json(num_proj)
            PASTA_ENTREGA_GLOBAL = pasta_entrega
            NUM_PROJETO_GLOBAL = num_proj

            # Abre a tela de visualização da entrega anterior.
            tela = TelaVisualizacaoEntregaAnterior(
                pasta_entregas=pasta_entrega,
                projeto_num=num_proj,
                disciplina=os.path.basename(os.path.dirname(pasta_entrega))
            )
            tela.mainloop()
            
            # Se a flag 'reabrir_disciplina' estiver definida, o loop da disciplina continua.
            if getattr(tela, "reabrir_disciplina", False):
                continue  # Reabrir seleção de disciplina.
            break  # Finaliza o ciclo do projeto, saindo do loop da disciplina.

        # Se o usuário optou por voltar ao projeto, o loop externo continua.
        if voltar_proj:
            continue
        return # Sai da função main.