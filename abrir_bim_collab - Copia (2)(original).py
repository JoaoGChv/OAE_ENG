import os
import re
import sys
import time
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import pyautogui
import pygetwindow as gw
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError

# Configurações de caminhos
ROOT_DIR = r"G:\Drives compartilhados"
APONTAMENTOS_DIR = os.path.join(ROOT_DIR, "OAE - APONTAMENTOS")
BCP_SUBFOLDER = "3 Desenvolvimento"
ICON_PROJECT = r"G:\Drives compartilhados\OAE - SCRIPTS\SCRIPTS\BIMCOLLAB.jpeg"
ICON_BCF = r"G:\Drives compartilhados\OAE - SCRIPTS\SCRIPTS\BCF.png"
# Tamanho máximo do ícone e largura extra para bounding box
ICON_MAX_WIDTH = 176
ICON_MAX_HEIGHT = 117
EXTRA_WIDTH = 50
# Grid
ICONS_PER_ROW = 4

# Regex para extrair chave do projeto
PROJECT_KEY_RE = re.compile(r"OAE-\d+")
# UI constants
UI_DELAY = 1.0
ISSUES_TAB_OFFSET = (200, 40)
MENU_OFFSET = (400, 40)
last_issues_rect = None


def load_icon(path, max_w=ICON_MAX_WIDTH, max_h=ICON_MAX_HEIGHT):
    img = Image.open(path)
    ratio = min(max_w / img.width, max_h / img.height)
    new_size = (int(img.width * ratio), int(img.height * ratio))
    resized = img.resize(new_size, Image.LANCZOS)
    return ImageTk.PhotoImage(resized)


def truncate(text, max_len=20):
    return text if len(text) <= max_len else text[:max_len-3] + '...'


def wait_for_bim_window(timeout=40):
    print("Aguardando janela do BIMcollab Zoom...")
    end = time.time() + timeout
    while time.time() < end:
        for title in gw.getAllTitles():
            if 'bimcollab' in title.lower():
                win = gw.getWindowsWithTitle(title)[0]
                try:
                    win.activate()
                except:
                    pyautogui.click(win.left + win.width//2, win.top + 10)
                print(f"Janela detectada: {title}")
                return title
        time.sleep(0.5)
    print("Timeout aguardando janela BIMcollab.")
    return None


def select_issues_tab(uia_win):
    global last_issues_rect
    try:
        elements = uia_win.descendants(control_type="CheckBox") + uia_win.descendants(control_type="TabItem")
        for e in elements:
            name = e.window_text() or ''
            if 'issue' in name.lower() or 'problema' in name.lower():
                rect = e.rectangle()
                last_issues_rect = rect
                try:
                    e.invoke()
                except:
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
        print(f"Clicando hamburger baseado na aba Issues em {(x, y)}")
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
        app = Application(backend='uia').connect(handle=hwnd)
        uia_win = app.window(handle=hwnd)
    except Exception as e:
        print(f"UIA connect erro: {e}")
        uia_win = None

    try:
        fallback_win.maximize()
    except:
        pass

    fallback_win.activate()
    pyautogui.moveTo(fallback_win.left + 10, fallback_win.top + 10)
    pyautogui.FAILSAFE = False
    time.sleep(0.5)

    if uia_win and select_issues_tab(uia_win):
        pass
    else:
        pyautogui.click(fallback_win.left + ISSUES_TAB_OFFSET[0], fallback_win.top + ISSUES_TAB_OFFSET[1])
    time.sleep(UI_DELAY)

    if uia_win and select_hamburger_menu(uia_win):
        pass
    else:
        pyautogui.click(fallback_win.left + MENU_OFFSET[0], fallback_win.top + MENU_OFFSET[1])
    time.sleep(UI_DELAY)

    pyautogui.press('down'); time.sleep(0.2)
    pyautogui.press('enter'); time.sleep(UI_DELAY)

    time.sleep(1)
    try:
        import pyperclip
        pyperclip.copy(bcf_path)
        pyautogui.hotkey('ctrl', 'v'); time.sleep(0.5)
        pyautogui.press('enter')
    except Exception:
        pyautogui.write(bcf_path); pyautogui.press('enter')


class Selector(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Selecionar Projeto")
        w = int(self.winfo_screenwidth() * 0.75)
        h = int(self.winfo_screenheight() * 0.75)
        self.geometry(f"{w}x{h}")
        self.icon_img = load_icon(ICON_PROJECT)
        self.project_list = self.scan_projects()
        self.create_widgets()

    def scan_projects(self):
        projects = []
        for name in os.listdir(ROOT_DIR):
            base = os.path.join(ROOT_DIR, name, BCP_SUBFOLDER)
            if os.path.isdir(base) and any(f.lower().endswith('.bcp') for f in os.listdir(base)):
                projects.append({"name": name, "path": base})
        return projects

    def create_widgets(self):
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self.update_list())
        ttk.Entry(self, textvariable=self.search_var).pack(fill="x", padx=10, pady=5)

        self.canvas = tk.Canvas(self)
        scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self.frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0,0), window=self.frame, anchor="nw")
        self.frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.update_list()

    def update_list(self):
        for w in self.frame.winfo_children(): w.destroy()
        q = self.search_var.get().lower()
        filtered = [p for p in self.project_list if q in p['name'].lower()]
        for i, proj in enumerate(filtered):
            text = truncate(proj['name'], 20)
            container = ttk.Frame(self.frame, width=ICON_MAX_WIDTH+EXTRA_WIDTH, height=ICON_MAX_HEIGHT)
            container.grid(row=i//ICONS_PER_ROW, column=i%ICONS_PER_ROW, padx=10, pady=10)
            container.grid_propagate(False)
            btn = ttk.Button(container, text=text, image=self.icon_img, compound="top",
                             command=lambda p=proj: self.open_bcp_selector(p))
            btn.pack(fill="both", expand=True)
        for col in range(ICONS_PER_ROW):
            self.frame.grid_columnconfigure(col, weight=1)

    def open_bcp_selector(self, project):
        self.withdraw()
        BcpSelector(self, project)


class BcpSelector(tk.Toplevel):
    def __init__(self, master, project):
        super().__init__(master)
        self.project = project
        self.title(f".BCP - {project['name']}")
        w = int(self.winfo_screenwidth() * 0.75)
        h = int(self.winfo_screenheight() * 0.75)
        self.geometry(f"{w}x{h}")
        self.icon_img = load_icon(ICON_PROJECT)

        bcp_dir = project['path']
        bcp_files = [f for f in os.listdir(bcp_dir) if f.lower().endswith('.bcp')]
        if not bcp_files:
            master.selected = {'bcp': None, 'bcf': None}
            master.master.destroy()
            process_selection(master.selected)
            return

        self.canvas = tk.Canvas(self)
        scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self.frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0,0), window=self.frame, anchor="nw")
        self.frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        for i, fname in enumerate(bcp_files):
            full = os.path.join(bcp_dir, fname)
            text = truncate(fname, 20)
            container = ttk.Frame(self.frame, width=ICON_MAX_WIDTH+EXTRA_WIDTH, height=ICON_MAX_HEIGHT)
            container.grid(row=i//ICONS_PER_ROW, column=i%ICONS_PER_ROW, padx=5, pady=5)
            container.grid_propagate(False)
            btn = ttk.Button(container, text=text, image=self.icon_img, compound="top",
                             command=lambda p=full: BcfSelector(self, p))
            btn.pack(fill="both", expand=True)
        for col in range(ICONS_PER_ROW):
            self.frame.grid_columnconfigure(col, weight=1)

        back = ttk.Button(self.frame, text="Voltar", command=self.go_back)
        back.grid(row=(len(bcp_files)//ICONS_PER_ROW)+1, column=0, columnspan=ICONS_PER_ROW, pady=10)

    def go_back(self):
        self.destroy()
        self.master.deiconify()


class BcfSelector(tk.Toplevel):
    def __init__(self, master, bcp_path):
        super().__init__(master)
        self.bcp_path = bcp_path
        self.title("Selecionar BCF")
        w = int(self.winfo_screenwidth() * 0.75)
        h = int(self.winfo_screenheight() * 0.75)
        self.geometry(f"{w}x{h}")
        self.icon_img = load_icon(ICON_BCF)

        m = PROJECT_KEY_RE.search(os.path.basename(bcp_path))
        key = m.group(0) if m else None
        dir_bcf = os.path.join(APONTAMENTOS_DIR, f"{key}-BCF") if key else None
        if not dir_bcf or not os.path.isdir(dir_bcf):
            root = master.master
            root.selected = {'bcp': bcp_path, 'bcf': None}
            root.destroy()
            process_selection(root.selected)
            return

        bcf_files = [f for f in os.listdir(dir_bcf) if f.lower().endswith('.bcf')]
        if not bcf_files:
            root = master.master
            root.selected = {'bcp': bcp_path, 'bcf': None}
            root.destroy()
            process_selection(root.selected)
            return

        self.canvas = tk.Canvas(self)
        scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self.frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0,0), window=self.frame, anchor="nw")
        self.frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        for i, fname in enumerate(bcf_files):
            full = os.path.join(dir_bcf, fname)
            text = truncate(fname, 20)
            container = ttk.Frame(self.frame, width=ICON_MAX_WIDTH+EXTRA_WIDTH, height=ICON_MAX_HEIGHT)
            container.grid(row=i//ICONS_PER_ROW, column=i%ICONS_PER_ROW, padx=5, pady=5)
            container.grid_propagate(False)
            btn = ttk.Button(container, text=text, image=self.icon_img, compound="top",
                             command=lambda p=full: self.finish_selection(p))
            btn.pack(fill="both", expand=True)
        for col in range(ICONS_PER_ROW):
            self.frame.grid_columnconfigure(col, weight=1)

        back = ttk.Button(self.frame, text="Voltar", command=self.go_back)
        back.grid(row=(len(bcf_files)//ICONS_PER_ROW)+1, column=0, columnspan=ICONS_PER_ROW, pady=10)

    def go_back(self):
        self.destroy()
        self.master.deiconify()

    def finish_selection(self, bcf_path):
        root = self.master.master
        root.selected = {'bcp': self.bcp_path, 'bcf': bcf_path}
        root.destroy()
        process_selection(root.selected)


def process_selection(selection):
    bcp = selection.get('bcp')
    bcf = selection.get('bcf')
    if bcp:
        os.startfile(bcp)
        wt = wait_for_bim_window()
        if not wt:
            print("Janela BIMcollab não encontrada.")
            sys.exit(1)
        time.sleep(15)
    if bcf:
        open_bcf_via_issues(bcf, wt)


if __name__ == '__main__':
    app = Selector()
    app.mainloop()
