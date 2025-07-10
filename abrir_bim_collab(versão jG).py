# abrir_bim_collab.py
import os
import re
import sys
import time
import tkinter as tk
from tkinter import ttk
from datetime import datetime
from PIL import Image, ImageTk
import pyautogui
import pygetwindow as gw
from pywinauto import Application

# ----------------------------------------------------------------------
# 0.  CONFIGURAÇÕES ORIGINAIS (INALTERADAS)
# ----------------------------------------------------------------------
ROOT_DIR = r"G:\Drives compartilhados"  # Diretório raiz para projetos BIMcollab
APONTAMENTOS_DIR = os.path.join(ROOT_DIR, "OAE - APONTAMENTOS") # Diretório para arquivos BCF
BCP_SUBFOLDER = "3 Desenvolvimento"  # Subpasta de BCPs dentro de cada projeto
ICON_PROJECT = r"G:\Drives compartilhados\OAE - SCRIPTS\SCRIPTS\BIMCOLLAB.jpeg" # Ícone para janela de projetos
ICON_BCF = r"G:\Drives compartilhados\OAE - SCRIPTS\SCRIPTS\BCF.png" # Ícone para janela de BCF

ICON_SMALL = 16            # Tamanho de ícone para Treeview
PROJECT_KEY_RE = re.compile(r"OAE-\d+") # Regex para extrair chave de projeto (e.g., OAE-123)
UI_DELAY = 1.0  # Atraso em segundos para automação de UI
ISSUES_TAB_OFFSET = (200, 40) # Offset para clique na aba "Issues" (fallback)
MENU_OFFSET = (400, 40) # Offset para clique no menu hambúrguer (fallback)
last_issues_rect = None # Armazena a última posição da aba "Issues" para automação

# ----------------------------------------------------------------------
# 1.  FUNÇÕES DE UTILIDADE (INALTERADAS + 2 helpers de formatação)
# ----------------------------------------------------------------------
def load_icon(path, size=ICON_SMALL): # Carrega e redimensiona um ícone para Tkinter
    img = Image.open(path) # Abre a imagem
    ratio = min(size / img.width, size / img.height) # Calcula a proporção de redimensionamento
    new_size = (int(img.width * ratio), int(img.height * ratio)) # Novo tamanho
    return ImageTk.PhotoImage(img.resize(new_size, Image.LANCZOS)) # Retorna PhotoImage redimensionado

def fmt_date(ts): # Formata um timestamp em data e hora
    return datetime.fromtimestamp(ts).strftime("%d/%m/%Y %H:%M") # Retorna string formatada

def fmt_size(bytes_): # Formata tamanho de arquivo em bytes para B, KB, MB, GB
    for unit in ("B", "KB", "MB", "GB"): # Itera sobre as unidades
        if bytes_ < 1024 or unit == "GB": # Se menor que 1KB ou for GB, retorna formatado
            return f"{bytes_:,.0f} {unit}" # Retorna string formatada
        bytes_ /= 1024 # Divide por 1024 para próxima unidade

def truncate(text, max_len=40): # Trunca texto para um comprimento máximo
    return text if len(text) <= max_len else text[: max_len - 3] + "..." # Retorna texto truncado ou original

# ----------------------------------------------------------------------
# 2.  AUTOMACAO BIMCOLLAB (100 % IGUAL AO ORIGINAL)
# ----------------------------------------------------------------------
def wait_for_bim_window(timeout=40): # Aguarda a janela do BIMcollab Zoom
    print("Aguardando janela do BIMcollab Zoom...") # Mensagem de status
    end = time.time() + timeout # Tempo limite
    while time.time() < end: # Loop até o tempo limite
        for title in gw.getAllTitles(): # Itera sobre os títulos das janelas
            if "bimcollab" in title.lower(): # Se "bimcollab" estiver no título (case-insensitive)
                win = gw.getWindowsWithTitle(title)[0] # Obtém o objeto da janela
                try:
                    win.activate() # Tenta ativar a janela
                except:  # noqa
                    pyautogui.click(win.left + win.width // 2, win.top + 10) # Clica no centro superior como fallback
                print(f"Janela detectada: {title}") # Mensagem de detecção
                return title # Retorna o título da janela
        time.sleep(0.5) # Pequeno atraso
    print("Timeout aguardando janela BIMcollab.") # Mensagem de timeout
    return None # Retorna None se a janela não for encontrada

def select_issues_tab(uia_win): # Seleciona a aba "Issues" no BIMcollab Zoom via UI Automation
    global last_issues_rect # Acessa a variável global
    try:
        elements = uia_win.descendants(control_type="CheckBox") + uia_win.descendants(
            control_type="TabItem"
        ) # Busca por CheckBox ou TabItem
        for e in elements: # Itera sobre os elementos
            name = e.window_text() or "" # Obtém o texto do elemento
            if "issue" in name.lower() or "problema" in name.lower(): # Se contiver "issue" ou "problema"
                last_issues_rect = e.rectangle() # Armazena a posição do elemento
                try:
                    e.invoke() # Tenta invocar o elemento
                except:  # noqa
                    e.click_input() # Clica no elemento como fallback
                return True # Retorna True se a aba for selecionada
    except Exception as e: # Captura exceções
        print(f"Erro em select_issues_tab: {e}") # Mensagem de erro
    return False # Retorna False se a aba não for selecionada

def select_hamburger_menu(uia_win): # Seleciona o menu hambúrguer no BIMcollab Zoom
    global last_issues_rect # Acessa a variável global
    if last_issues_rect: # Se a posição da aba "Issues" for conhecida
        width = last_issues_rect.right - last_issues_rect.left # Largura da aba
        x = int(last_issues_rect.right - width * 0.10) # Calcula a posição X do menu
        y = last_issues_rect.bottom + 5 # Calcula a posição Y do menu
        try:
            pyautogui.click(x, y) # Clica nas coordenadas calculadas
            return True # Retorna True se o clique for bem-sucedido
        except Exception as e: # Captura exceções
            print(f"Falha click hamburger dinâmico: {e}") # Mensagem de erro
    return False # Retorna False se o menu não for selecionado

def open_bcf_via_issues(bcf_path, win_title): # Abre um arquivo BCF no BIMcollab Zoom
    print(f"Aguardando e abrindo BCF: {bcf_path}") # Mensagem de status
    try:
        fallback_win = gw.getWindowsWithTitle(win_title)[0] # Obtém a janela do BIMcollab Zoom
    except IndexError: # Se a janela não for encontrada
        print("Janela BIMcollab não encontrada para fallback.") # Mensagem de erro
        return # Sai da função

    hwnd = fallback_win._hWnd # Handle da janela
    try:
        app = Application(backend="uia").connect(handle=hwnd) # Conecta à aplicação via UIA
        uia_win = app.window(handle=hwnd) # Obtém o objeto da janela UIA
    except Exception as e: # Captura exceções
        print(f"UIA connect erro: {e}") # Mensagem de erro
        uia_win = None # Define uia_win como None

    try:
        fallback_win.maximize() # Tenta maximizar a janela
    except:  # noqa
        pass # Ignora erro ao maximizar

    fallback_win.activate() # Ativa a janela
    pyautogui.moveTo(fallback_win.left + 10, fallback_win.top + 10) # Move o mouse para a janela
    pyautogui.FAILSAFE = False # Desabilita o failsafe do pyautogui
    time.sleep(0.5) # Pequeno atraso

    if not (uia_win and select_issues_tab(uia_win)): # Tenta selecionar a aba "Issues" via UIA, se falhar
        pyautogui.click(
            fallback_win.left + ISSUES_TAB_OFFSET[0], fallback_win.top + ISSUES_TAB_OFFSET[1]
        ) # Clica na aba "Issues" usando offset
    time.sleep(UI_DELAY) # Atraso

    if not (uia_win and select_hamburger_menu(uia_win)): # Tenta selecionar o menu hambúrguer via UIA, se falhar
        pyautogui.click(
            fallback_win.left + MENU_OFFSET[0], fallback_win.top + MENU_OFFSET[1]
        ) # Clica no menu hambúrguer usando offset
    time.sleep(UI_DELAY) # Atraso

    pyautogui.press("down") # Pressiona "seta para baixo"
    time.sleep(0.2) # Pequeno atraso
    pyautogui.press("enter") # Pressiona "Enter"
    time.sleep(UI_DELAY) # Atraso

    time.sleep(1) # Atraso
    try:
        import pyperclip # Importa pyperclip
        pyperclip.copy(bcf_path) # Copia o caminho do BCF para a área de transferência
        pyautogui.hotkey("ctrl", "v") # Cola o caminho
        time.sleep(0.5) # Pequeno atraso
        pyautogui.press("enter") # Pressiona "Enter"
    except Exception: # Se pyperclip falhar
        pyautogui.write(bcf_path) # Digita o caminho
        pyautogui.press("enter") # Pressiona "Enter"

# ----------------------------------------------------------------------
# 3.  WIDGET REUTILIZÁVEL: LISTA DE DETALHES COM SCROLL
# ----------------------------------------------------------------------
class DetailsList(ttk.Frame): # Widget Treeview com scrollbar para exibir detalhes
    """Treeview com colunas + scrollbar. Usa dict interno para payload."""
    def __init__(self, master, columns, on_open): # Construtor da classe
        super().__init__(master) # Chama o construtor da classe pai
        self.on_open = on_open # Função de callback para abrir item
        self._payload = {}              # Dicionário para mapear id do item no Treeview para o objeto de dados

        col_ids = [c[0] for c in columns] # Extrai os IDs das colunas
        self.tree = ttk.Treeview(self, columns=col_ids, show="headings",
                                 selectmode="browse") # Cria o Treeview
        vbar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview) # Cria a barra de rolagem vertical
        self.tree.configure(yscrollcommand=vbar.set) # Conecta a barra de rolagem ao Treeview
        self.tree.grid(row=0, column=0, sticky="nsew") # Posiciona o Treeview na grid
        vbar.grid(row=0, column=1, sticky="ns") # Posiciona a barra de rolagem na grid
        self.grid_rowconfigure(0, weight=1) # Configura o redimensionamento da linha
        self.grid_columnconfigure(0, weight=1) # Configura o redimensionamento da coluna

        for col, heading, width in columns: # Configura as colunas
            self.tree.heading(col, text=heading, anchor="w") # Define o cabeçalho da coluna
            self.tree.column(col, anchor="w", width=width) # Define o alinhamento e largura da coluna

        self.tree.bind("<Double-1>", self._open_selected) # Associa clique duplo ao _open_selected
        self.tree.bind("<Return>",    self._open_selected) # Associa tecla Enter ao _open_selected

    def populate(self, rows): # Popula o Treeview com dados
        """rows = list de (col1, col2, ... , payload_obj)."""
        self.tree.delete(*self.tree.get_children()) # Limpa os itens existentes
        self._payload.clear() # Limpa o payload
        for row in rows: # Itera sobre as linhas
            *visible, payload = row # Desempacota valores visíveis e o payload
            item_id = self.tree.insert("", "end", values=visible) # Insere o item no Treeview
            self._payload[item_id] = payload      # Armazena o objeto real (payload)

    def _open_selected(self, _e): # Método chamado ao selecionar um item
        sel = self.tree.selection() # Obtém o item selecionado
        if sel: # Se um item estiver selecionado
            self.on_open(self._payload[sel[0]]) # Chama a função on_open com o payload

# ----------------------------------------------------------------------
# 4.  JANELA DE PROJETOS  (só UI trocada)
# ----------------------------------------------------------------------
class Selector(tk.Tk): # Janela principal para seleção de projetos
    def __init__(self): # Construtor da classe
        super().__init__() # Chama o construtor da classe pai
        self.icon = load_icon(ICON_PROJECT) # Carrega o ícone da janela
        self.title("Selecionar Projeto") # Define o título da janela
        w = int(self.winfo_screenwidth() * 0.75) # Largura da janela (75% da tela)
        h = int(self.winfo_screenheight() * 0.75) # Altura da janela (75% da tela)
        self.geometry(f"{w}x{h}") # Define o tamanho da janela
        self.project_list = self.scan_projects() # Escaneia e obtém a lista de projetos
        self.create_widgets() # Cria os widgets da interface

    # ------------ lógica original ----------------
    def scan_projects(self): # Escaneia o diretório raiz em busca de projetos BIMcollab
        projects = [] # Lista para armazenar os projetos
        for name in os.listdir(ROOT_DIR): # Itera sobre os itens no diretório raiz
            base = os.path.join(ROOT_DIR, name, BCP_SUBFOLDER) # Caminho para a subpasta BCP
            if os.path.isdir(base): # Se a subpasta BCP existir
                bcps = [f for f in os.listdir(base) if f.lower().endswith(".bcp")] # Lista de arquivos .bcp
                if bcps: # Se houver arquivos .bcp
                    mod_ts = max(
                        os.path.getmtime(os.path.join(base, f)) for f in bcps
                    ) # Data de modificação do .bcp mais recente
                    projects.append(
                        {
                            "name": name, # Nome do projeto
                            "path": base, # Caminho da subpasta BCP
                            "count": len(bcps), # Contagem de arquivos .bcp
                            "mod": mod_ts, # Data de modificação
                        }
                    )
        return projects # Retorna a lista de projetos
    # ----------------------------------------------

    def create_widgets(self): # Cria os elementos da interface (widgets)
        # busca
        self.search_var = tk.StringVar() # Variável para o campo de busca
        self.search_var.trace_add("write", lambda *_: self.refresh_list()) # Associa evento de digitação à função refresh_list
        ttk.Entry(self, textvariable=self.search_var).pack(fill="x", padx=8, pady=4) # Campo de entrada para busca

        # lista de detalhes
        cols = [
            ("name", "Projeto", 300), # Coluna para o nome do projeto
            ("mod", "Modificado em", 150), # Coluna para a data de modificação
            ("count", "BCP(s)", 80), # Coluna para a contagem de BCPs
            ("payload", "", 0),  # Coluna invisível para o payload (objeto do projeto)
        ]
        self.dlist = DetailsList(self, cols, self._open_project) # Cria a lista de detalhes
        self.dlist.pack(fill="both", expand=True, padx=5, pady=5) # Posiciona a lista na janela
        self.refresh_list() # Atualiza a lista inicialmente

    def refresh_list(self): # Atualiza a lista de projetos com base na busca
        q = self.search_var.get().lower() # Obtém o texto de busca (minúsculas)
        rows = [] # Lista para armazenar as linhas da Treeview
        for proj in self.project_list: # Itera sobre a lista de projetos
            if q not in proj["name"].lower(): # Se o texto de busca não estiver no nome do projeto
                continue # Pula para o próximo projeto
            rows.append(
                (
                    proj["name"], # Nome do projeto
                    fmt_date(proj["mod"]), # Data de modificação formatada
                    proj["count"], # Contagem de BCPs
                    proj,  # Payload: dicionário completo do projeto
                )
            )
        self.dlist.populate(rows) # Popula a lista de detalhes com as linhas

    def _open_project(self, proj_dict): # Abre a janela de seleção de BCP para o projeto selecionado
        self.withdraw() # Esconde a janela atual
        BcpSelector(self, proj_dict) # Cria e exibe a janela de seleção de BCP

# ----------------------------------------------------------------------
# 5.  JANELA DE FEDERADOS (.BCP)  (só UI trocada)
# ----------------------------------------------------------------------
class BcpSelector(tk.Toplevel): # Janela para seleção de arquivos .bcp
    def __init__(self, master, project): # Construtor da classe
        super().__init__(master) # Chama o construtor da classe pai
        self.icon = load_icon(ICON_PROJECT) # Carrega o ícone da janela
        self.project = project # Dicionário do projeto selecionado
        self.title(f".BCP – {project['name']}") # Define o título da janela
        w = int(self.winfo_screenwidth() * 0.75) # Largura da janela
        h = int(self.winfo_screenheight() * 0.75) # Altura da janela
        self.geometry(f"{w}x{h}") # Define o tamanho da janela

        bcp_dir = project["path"] # Caminho do diretório BCP
        bcp_files = [f for f in os.listdir(bcp_dir) if f.lower().endswith(".bcp")] # Lista de arquivos .bcp

        if not bcp_files: # Se não houver arquivos .bcp
            master.selected = {"bcp": None, "bcf": None} # Define seleção como nula
            master.destroy() # Destrói a janela mestre
            process_selection(master.selected) # Processa a seleção (sem BCP/BCF)
            return # Sai do construtor

        cols = [
            ("file", "Arquivo", 350), # Coluna para o nome do arquivo
            ("mod", "Modificado em", 150), # Coluna para a data de modificação
            ("size", "Tamanho", 100), # Coluna para o tamanho do arquivo
            ("payload", "", 0), # Coluna invisível para o payload (caminho completo do BCP)
        ]
        self.dlist = DetailsList(self, cols, self._open_bcf) # Cria a lista de detalhes
        self.dlist.pack(fill="both", expand=True, padx=5, pady=5) # Posiciona a lista na janela

        rows = [] # Lista para armazenar as linhas da Treeview
        for fname in bcp_files: # Itera sobre os arquivos .bcp
            full = os.path.join(bcp_dir, fname) # Caminho completo do arquivo .bcp
            rows.append(
                (
                    fname, # Nome do arquivo
                    fmt_date(os.path.getmtime(full)), # Data de modificação formatada
                    fmt_size(os.path.getsize(full)), # Tamanho do arquivo formatado
                    full, # Payload: caminho completo do BCP
                )
            )
        self.dlist.populate(rows) # Popula a lista de detalhes

        ttk.Button(self, text="Voltar", command=self._go_back).pack(pady=4) # Botão para voltar

    def _go_back(self): # Volta para a janela de seleção de projetos
        self.destroy() # Destrói a janela atual
        self.master.deiconify() # Reexibe a janela mestre (Selector)

    def _open_bcf(self, full_bcp_path): # Abre a janela de seleção de BCF para o BCP selecionado
        BcfSelector(self, full_bcp_path) # Cria e exibe a janela de seleção de BCF

# ----------------------------------------------------------------------
# 6.  JANELA DE BCF (INALTERADA – grade foi mantida, mas poderia usar Tree)
# ----------------------------------------------------------------------
class BcfSelector(tk.Toplevel): # Janela para seleção de arquivos .bcf
    def __init__(self, master, bcp_path): # Construtor da classe
        super().__init__(master) # Chama o construtor da classe pai
        self.icon = load_icon(ICON_BCF) # Carrega o ícone da janela
        self.bcp_path = bcp_path # Caminho do arquivo .bcp selecionado
        self.title("Selecionar BCF") # Define o título da janela
        w = int(self.winfo_screenwidth() * 0.75) # Largura da janela
        h = int(self.winfo_screenheight() * 0.75) # Altura da janela
        self.geometry(f"{w}x{h}") # Define o tamanho da janela

        m = PROJECT_KEY_RE.search(os.path.basename(bcp_path)) # Busca a chave do projeto no nome do BCP
        key = m.group(0) if m else None # Extrai a chave se encontrada
        dir_bcf = os.path.join(APONTAMENTOS_DIR, f"{key}-BCF") if key else None # Caminho do diretório BCF
        if not dir_bcf or not os.path.isdir(dir_bcf): # Se o diretório BCF não existir
            root = master.master # Janela raiz (Selector)
            root.selected = {"bcp": bcp_path, "bcf": None} # Define seleção (apenas BCP)
            root.destroy() # Destrói a janela raiz
            process_selection(root.selected) # Processa a seleção
            return # Sai do construtor

        bcf_files = [
            os.path.join(dir_bcf, f) for f in os.listdir(dir_bcf) if f.lower().endswith(".bcf")
        ] # Lista de arquivos .bcf
        if not bcf_files: # Se não houver arquivos .bcf
            root = master.master # Janela raiz (Selector)
            root.selected = {"bcp": bcp_path, "bcf": None} # Define seleção (apenas BCP)
            root.destroy() # Destrói a janela raiz
            process_selection(root.selected) # Processa a seleção
            return # Sai do construtor

        # Lista em Treeview também, para consistência
        cols = [
            ("file", "BCF", 350), # Coluna para o nome do arquivo BCF
            ("mod", "Modificado em", 150), # Coluna para a data de modificação
            ("size", "Tamanho", 100), # Coluna para o tamanho
            ("payload", "", 0), # Coluna invisível para o payload (caminho completo do BCF)
        ]
        self.dlist = DetailsList(self, cols, self._finish_selection) # Cria a lista de detalhes
        self.dlist.pack(fill="both", expand=True, padx=5, pady=5) # Posiciona a lista na janela

        rows = [] # Lista para armazenar as linhas da Treeview
        for full in bcf_files: # Itera sobre os arquivos .bcf
            rows.append(
                (
                    os.path.basename(full), # Nome base do arquivo BCF
                    fmt_date(os.path.getmtime(full)), # Data de modificação formatada
                    fmt_size(os.path.getsize(full)), # Tamanho do arquivo formatado
                    full, # Payload: caminho completo do BCF
                )
            )
        self.dlist.populate(rows) # Popula a lista de detalhes

        ttk.Button(self, text="Voltar", command=self._go_back).pack(pady=4) # Botão para voltar

    def _go_back(self): # Volta para a janela de seleção de BCP
        self.destroy() # Destrói a janela atual
        self.master.deiconify() # Reexibe a janela mestre (BcpSelector)

    def _finish_selection(self, bcf_path): # Finaliza a seleção e inicia o processo de abertura
        root = self.master.master # Janela raiz (Selector)
        root.selected = {"bcp": self.bcp_path, "bcf": bcf_path} # Armazena os caminhos selecionados
        root.destroy() # Destrói a janela raiz
        process_selection(root.selected) # Processa a seleção final

# ----------------------------------------------------------------------
# 7.  PROCESSAMENTO FINAL (INALTERADO)
# ----------------------------------------------------------------------
def process_selection(selection): # Processa os arquivos BCP e BCF selecionados
    bcp = selection.get("bcp") # Obtém o caminho do BCP
    bcf = selection.get("bcf") # Obtém o caminho do BCF
    if bcp: # Se um BCP foi selecionado
        os.startfile(bcp) # Abre o arquivo .bcp com o aplicativo padrão
        wt = wait_for_bim_window() # Aguarda a janela do BIMcollab Zoom
        if not wt: # Se a janela não for encontrada
            print("Janela BIMcollab não encontrada.") # Mensagem de erro
            sys.exit(1) # Sai do programa com erro
        time.sleep(10) # Aguarda 10 segundos para o BIMcollab carregar
    if bcf: # Se um BCF foi selecionado
        open_bcf_via_issues(bcf, wt) # Abre o arquivo BCF dentro do BIMcollab Zoom

# ----------------------------------------------------------------------
# 8.  MAIN
# ----------------------------------------------------------------------
if __name__ == "__main__": # Executa o código principal quando o script é chamado diretamente
    app = Selector() # Cria uma instância da janela principal de seleção
    app.mainloop() # Inicia o loop de eventos do Tkinter para exibir a GUI