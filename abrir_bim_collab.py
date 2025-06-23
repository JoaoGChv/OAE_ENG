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
ROOT_DIR = r"G:\Drives compartilhados"
APONTAMENTOS_DIR = os.path.join(ROOT_DIR, "OAE - APONTAMENTOS")
BCP_SUBFOLDER = "3 Desenvolvimento"
ICON_PROJECT = r"G:\Drives compartilhados\OAE - SCRIPTS\SCRIPTS\BIMCOLLAB.jpeg"
ICON_BCF = r"G:\Drives compartilhados\OAE - SCRIPTS\SCRIPTS\BCF.png"

ICON_SMALL = 16          # ícone menor para Treeview
PROJECT_KEY_RE = re.compile(r"OAE-\d+")
UI_DELAY = 1.0
ISSUES_TAB_OFFSET = (200, 40)
MENU_OFFSET = (400, 40)
last_issues_rect = None

# ----------------------------------------------------------------------
# 1.  FUNÇÕES DE UTILIDADE (INALTERADAS + 2 helpers de formatação)
# ----------------------------------------------------------------------
def load_icon(path, size=ICON_SMALL):
    img = Image.open(path)
    ratio = min(size / img.width, size / img.height)
    new_size = (int(img.width * ratio), int(img.height * ratio))
    return ImageTk.PhotoImage(img.resize(new_size, Image.LANCZOS))

def fmt_date(ts):
    return datetime.fromtimestamp(ts).strftime("%d/%m/%Y %H:%M")

def fmt_size(bytes_):
    for unit in ("B", "KB", "MB", "GB"):
        if bytes_ < 1024 or unit == "GB":
            return f"{bytes_:,.0f} {unit}"
        bytes_ /= 1024

def truncate(text, max_len=40):
    return text if len(text) <= max_len else text[: max_len - 3] + "..."

# ----------------------------------------------------------------------
# 2.  AUTOMACAO BIMCOLLAB (100 % IGUAL AO ORIGINAL)
# ----------------------------------------------------------------------
def wait_for_bim_window(timeout=40):
    print("Aguardando janela do BIMcollab Zoom...")
    end = time.time() + timeout
    while time.time() < end:
        for title in gw.getAllTitles():
            if "bimcollab" in title.lower():
                win = gw.getWindowsWithTitle(title)[0]
                try:
                    win.activate()
                except:  # noqa
                    pyautogui.click(win.left + win.width // 2, win.top + 10)
                print(f"Janela detectada: {title}")
                return title
        time.sleep(0.5)
    print("Timeout aguardando janela BIMcollab.")
    return None

def select_issues_tab(uia_win):
    global last_issues_rect
    try:
        elements = uia_win.descendants(control_type="CheckBox") + uia_win.descendants(
            control_type="TabItem"
        )
        for e in elements:
            name = e.window_text() or ""
            if "issue" in name.lower() or "problema" in name.lower():
                last_issues_rect = e.rectangle()
                try:
                    e.invoke()
                except:  # noqa
                    e.click_input()
                return True
    except Exception as e:
        print(f"Erro em select_issues_tab: {e}")
    return False

def select_hamburger_menu(uia_win):
    global last_issues_rect
    if last_issues_rect:
        width = last_issues_rect.right - last_issues_rect.left
        x = int(last_issues_rect.right - width * 0.10)
        y = last_issues_rect.bottom + 5
        try:
            pyautogui.click(x, y)
            return True
        except Exception as e:
            print(f"Falha click hamburger dinâmico: {e}")
    return False

def open_bcf_via_issues(bcf_path, win_title):
    print(f"Aguardando e abrindo BCF: {bcf_path}")
    try:
        fallback_win = gw.getWindowsWithTitle(win_title)[0]
    except IndexError:
        print("Janela BIMcollab não encontrada para fallback.")
        return

    hwnd = fallback_win._hWnd
    try:
        app = Application(backend="uia").connect(handle=hwnd)
        uia_win = app.window(handle=hwnd)
    except Exception as e:
        print(f"UIA connect erro: {e}")
        uia_win = None

    try:
        fallback_win.maximize()
    except:  # noqa
        pass

    fallback_win.activate()
    pyautogui.moveTo(fallback_win.left + 10, fallback_win.top + 10)
    pyautogui.FAILSAFE = False
    time.sleep(0.5)

    if not (uia_win and select_issues_tab(uia_win)):
        pyautogui.click(
            fallback_win.left + ISSUES_TAB_OFFSET[0], fallback_win.top + ISSUES_TAB_OFFSET[1]
        )
    time.sleep(UI_DELAY)

    if not (uia_win and select_hamburger_menu(uia_win)):
        pyautogui.click(
            fallback_win.left + MENU_OFFSET[0], fallback_win.top + MENU_OFFSET[1]
        )
    time.sleep(UI_DELAY)

    pyautogui.press("down")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(UI_DELAY)

    time.sleep(1)
    try:
        import pyperclip

        pyperclip.copy(bcf_path)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.5)
        pyautogui.press("enter")
    except Exception:
        pyautogui.write(bcf_path)
        pyautogui.press("enter")

# ----------------------------------------------------------------------
# 3.  WIDGET REUTILIZÁVEL: LISTA DE DETALHES COM SCROLL
# ----------------------------------------------------------------------
class DetailsList(ttk.Frame):
    """Treeview com colunas + scrollbar. Usa dict interno para payload."""
    def __init__(self, master, columns, on_open):
        super().__init__(master)
        self.on_open = on_open
        self._payload = {}                       # id → objeto

        col_ids = [c[0] for c in columns]
        self.tree = ttk.Treeview(self, columns=col_ids, show="headings",
                                 selectmode="browse")
        vbar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        for col, heading, width in columns:
            self.tree.heading(col, text=heading, anchor="w")
            self.tree.column(col, anchor="w", width=width)

        self.tree.bind("<Double-1>", self._open_selected)
        self.tree.bind("<Return>",    self._open_selected)

    def populate(self, rows):
        """rows = list de (col1, col2, … , payload_obj)."""
        self.tree.delete(*self.tree.get_children())
        self._payload.clear()
        for row in rows:
            *visible, payload = row
            item_id = self.tree.insert("", "end", values=visible)
            self._payload[item_id] = payload      # guarda objeto real

    def _open_selected(self, _e):
        sel = self.tree.selection()
        if sel:
            self.on_open(self._payload[sel[0]])

# ----------------------------------------------------------------------
# 4.  JANELA DE PROJETOS  (só UI trocada)
# ----------------------------------------------------------------------
class Selector(tk.Tk):
    def __init__(self):
        super().__init__()
        self.icon = load_icon(ICON_PROJECT)          # ← agora é seguro
        self.title("Selecionar Projeto")
        w = int(self.winfo_screenwidth() * 0.75)
        h = int(self.winfo_screenheight() * 0.75)
        self.geometry(f"{w}x{h}")
        self.project_list = self.scan_projects()
        self.create_widgets()

    # ------------ lógica original ----------------
    def scan_projects(self):
        projects = []
        for name in os.listdir(ROOT_DIR):
            base = os.path.join(ROOT_DIR, name, BCP_SUBFOLDER)
            if os.path.isdir(base):
                bcps = [f for f in os.listdir(base) if f.lower().endswith(".bcp")]
                if bcps:
                    mod_ts = max(
                        os.path.getmtime(os.path.join(base, f)) for f in bcps
                    )
                    projects.append(
                        {
                            "name": name,
                            "path": base,
                            "count": len(bcps),
                            "mod": mod_ts,
                        }
                    )
        return projects
    # ----------------------------------------------

    def create_widgets(self):
        # busca
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh_list())
        ttk.Entry(self, textvariable=self.search_var).pack(fill="x", padx=8, pady=4)

        # lista de detalhes
        cols = [
            ("name", "Projeto", 300),
            ("mod", "Modificado em", 150),
            ("count", "BCP(s)", 80),
            ("payload", "", 0),  # payload invisível
        ]
        self.dlist = DetailsList(self, cols, self._open_project)
        self.dlist.pack(fill="both", expand=True, padx=5, pady=5)
        self.refresh_list()

    def refresh_list(self):
        q = self.search_var.get().lower()
        rows = []
        for proj in self.project_list:
            if q not in proj["name"].lower():
                continue
            rows.append(
                (
                    proj["name"],
                    fmt_date(proj["mod"]),
                    proj["count"],
                    proj,  # payload
                )
            )
        self.dlist.populate(rows)

    def _open_project(self, proj_dict):
        self.withdraw()
        BcpSelector(self, proj_dict)

# ----------------------------------------------------------------------
# 5.  JANELA DE FEDERADOS (.BCP)  (só UI trocada)
# ----------------------------------------------------------------------
class BcpSelector(tk.Toplevel):
    def __init__(self, master, project):
        super().__init__(master)
        self.icon = load_icon(ICON_PROJECT)          # ← seguro aqui
        self.project = project
        self.title(f".BCP – {project['name']}")
        w = int(self.winfo_screenwidth() * 0.75)
        h = int(self.winfo_screenheight() * 0.75)
        self.geometry(f"{w}x{h}")

        bcp_dir = project["path"]
        bcp_files = [f for f in os.listdir(bcp_dir) if f.lower().endswith(".bcp")]

        if not bcp_files:
            master.selected = {"bcp": None, "bcf": None}
            master.destroy()
            process_selection(master.selected)
            return

        cols = [
            ("file", "Arquivo", 350),
            ("mod", "Modificado em", 150),
            ("size", "Tamanho", 100),
            ("payload", "", 0),
        ]
        self.dlist = DetailsList(self, cols, self._open_bcf)
        self.dlist.pack(fill="both", expand=True, padx=5, pady=5)

        rows = []
        for fname in bcp_files:
            full = os.path.join(bcp_dir, fname)
            rows.append(
                (
                    fname,
                    fmt_date(os.path.getmtime(full)),
                    fmt_size(os.path.getsize(full)),
                    full,
                )
            )
        self.dlist.populate(rows)

        ttk.Button(self, text="Voltar", command=self._go_back).pack(pady=4)

    def _go_back(self):
        self.destroy()
        self.master.deiconify()

    def _open_bcf(self, full_bcp_path):
        BcfSelector(self, full_bcp_path)

# ----------------------------------------------------------------------
# 6.  JANELA DE BCF (INALTERADA – grade foi mantida, mas poderia usar Tree)
# ----------------------------------------------------------------------
class BcfSelector(tk.Toplevel):
    def __init__(self, master, bcp_path):
        super().__init__(master)
        self.icon = load_icon(ICON_BCF)              # ← seguro aqui
        self.bcp_path = bcp_path
        self.title("Selecionar BCF")
        w = int(self.winfo_screenwidth() * 0.75)
        h = int(self.winfo_screenheight() * 0.75)
        self.geometry(f"{w}x{h}")

        m = PROJECT_KEY_RE.search(os.path.basename(bcp_path))
        key = m.group(0) if m else None
        dir_bcf = os.path.join(APONTAMENTOS_DIR, f"{key}-BCF") if key else None
        if not dir_bcf or not os.path.isdir(dir_bcf):
            root = master.master
            root.selected = {"bcp": bcp_path, "bcf": None}
            root.destroy()
            process_selection(root.selected)
            return

        bcf_files = [
            os.path.join(dir_bcf, f) for f in os.listdir(dir_bcf) if f.lower().endswith(".bcf")
        ]
        if not bcf_files:
            root = master.master
            root.selected = {"bcp": bcp_path, "bcf": None}
            root.destroy()
            process_selection(root.selected)
            return

        # Lista em Treeview também, para consistência
        cols = [
            ("file", "BCF", 350),
            ("mod", "Modificado em", 150),
            ("size", "Tamanho", 100),
            ("payload", "", 0),
        ]
        self.dlist = DetailsList(self, cols, self._finish_selection)
        self.dlist.pack(fill="both", expand=True, padx=5, pady=5)

        rows = []
        for full in bcf_files:
            rows.append(
                (
                    os.path.basename(full),
                    fmt_date(os.path.getmtime(full)),
                    fmt_size(os.path.getsize(full)),
                    full,
                )
            )
        self.dlist.populate(rows)

        ttk.Button(self, text="Voltar", command=self._go_back).pack(pady=4)

    def _go_back(self):
        self.destroy()
        self.master.deiconify()

    def _finish_selection(self, bcf_path):
        root = self.master.master
        root.selected = {"bcp": self.bcp_path, "bcf": bcf_path}
        root.destroy()
        process_selection(root.selected)

# ----------------------------------------------------------------------
# 7.  PROCESSAMENTO FINAL (INALTERADO)
# ----------------------------------------------------------------------
def process_selection(selection):
    bcp = selection.get("bcp")
    bcf = selection.get("bcf")
    if bcp:
        os.startfile(bcp)
        wt = wait_for_bim_window()
        if not wt:
            print("Janela BIMcollab não encontrada.")
            sys.exit(1)
        time.sleep(10)
    if bcf:
        open_bcf_via_issues(bcf, wt)

# ----------------------------------------------------------------------
# 8.  MAIN
# ----------------------------------------------------------------------
if __name__ == "__main__":
    app = Selector()
    app.mainloop()
