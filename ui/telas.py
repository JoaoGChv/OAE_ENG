import os, tkinter as tk, json, shutil, sys
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from typing import Dict, Any
import ttkbootstrap
from ttkbootstrap.style import Style
from ttkbootstrap.constants import *

MARGIN_SIZE = 10
JSON_FILE_PATH = "dados_projetos.json"
PROJETOS_JSON = r"G:\Drives compartilhados\OAE-JSONS\diretorios_projetos.json"
ULTIMO_DIRETORIO_JSON = "ultimo_diretorio.json"
HISTORICO_JSON = "historico_arquivos.json"
PADROES_JSON = r"G:\Drives compartilhados\OAE - SCRIPTS\SCRIPTS\tmp_joaoG\Melhorias\Código_reformulado_teste\ui\padrões.json"
CONFIG_NOMENCLATURA_JSON = "config_nomenclatura.json"

# Mantemos as regras como no original.
NOMENCLATURA_RULES = {
    "Status": ["E", "C", "P", "A", "R"],
    "Fase": ["PE", "AP", "EX"],
    "Tipo": ["ARQ", "ELE", "HID", "SCO", "DTE", "MEM", "RLT"]
}

def carregar_config_nomenclatura():
    if os.path.exists(CONFIG_NOMENCLATURA_JSON):
        with open(CONFIG_NOMENCLATURA_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        return {"config_revisao": {}, "config_valores": {}}

def carregar_padroes_nomenclatura():
    if os.path.exists(PADROES_JSON):
        with open(PADROES_JSON, "r", encoding="utf-8") as f:
            return json.load(f).get("padroes", [])
    else:
        return []

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

def salvar_ultimo_diretorio(diretorio):
    with open(ULTIMO_DIRETORIO_JSON, "w", encoding="utf-8") as f:
        json.dump({"ultimo_diretorio": diretorio}, f)

def atualizar_historico(lista_arquivos, caminho_json=HISTORICO_JSON):
    historico = {}
    if os.path.exists(caminho_json):
        with open(caminho_json, "r", encoding="utf-8") as f:
            historico = json.load(f)

    for arquivo in lista_arquivos:
        data_modificacao = os.path.getmtime(arquivo)
        if arquivo not in historico or historico[arquivo]["data"] != data_modificacao:
            historico[arquivo] = {
                "numero": len(historico) + 1,
                "data": data_modificacao,
                "status": "Atual"
            }

    mais_recente = max(historico, key=lambda x: historico[x]["data"])
    for arq in historico:
        historico[arq]["status"] = "Atual" if arq == mais_recente else "Obsoleto"

    with open(caminho_json, "w", encoding="utf-8") as f:
        json.dump(historico, f, indent=4, ensure_ascii=False)
    return historico

def criar_pastas_organizacao():
    base_dir = r"C:\Users\PROJETOS\Downloads\OAE-467 - PETER-KD-ENG"
    pasta_revisados = os.path.join(base_dir, "Revisados")
    if not os.path.exists(pasta_revisados):
        os.makedirs(pasta_revisados)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    pasta_obsoletos = os.path.join(base_dir, f"Obsoleto_{timestamp}")
    if not os.path.exists(pasta_obsoletos):
        os.makedirs(pasta_obsoletos)
    return pasta_revisados, pasta_obsoletos

def mover_arquivos(lista_arquivos, destino):
    for idx, arq in enumerate(lista_arquivos):
        origem = arq.get("caminho")
        nome_arquivo = arq.get("Nome do Arquivo")
        if not origem:
            messagebox.showerror("Erro", f"O caminho para o arquivo '{nome_arquivo}' não foi encontrado.")
            continue
        if not os.path.exists(origem):
            messagebox.showerror("Erro", f"O arquivo '{nome_arquivo}' com caminho '{origem}' não existe ou não é acessível.")
            continue
        destino_arquivo = os.path.join(destino, nome_arquivo)
        try:
            shutil.move(origem, destino_arquivo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao mover o arquivo '{nome_arquivo}' com caminho '{origem}': {e}")

def mover_obsoletos(lista_obsoletos, destino):
    for idx, arq in enumerate(lista_obsoletos):
        origem = arq.get("caminho")
        nome_arquivo = arq.get("Nome do Arquivo")
        if not origem:
            messagebox.showerror("Erro", f"O caminho para o arquivo '{nome_arquivo}' não foi encontrado.")
            continue
        if not os.path.exists(origem):
            messagebox.showerror("Erro", f"O arquivo '{nome_arquivo}' com caminho '{origem}' não existe ou não é acessível.")
            continue
        base, ext = os.path.splitext(nome_arquivo)
        destino_arquivo = os.path.join(destino, base + "_OBSOLETO" + ext)
        try:
            shutil.move(origem, destino_arquivo)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao mover o arquivo obsoleto '{nome_arquivo}' com caminho '{origem}': {e}")

def pos_processamento(primeira_entrega, diretorio, dados_anteriores, arquivos_novos, arquivos_revisados, arquivos_alterados, obsoletos):
    pasta_revisados, pasta_obsoletos = criar_pastas_organizacao()
    if pasta_revisados is None or pasta_obsoletos is None:
        messagebox.showerror("Erro", "Não foi possível criar pastas para organizar os arquivos.")
        return
    mover_arquivos(arquivos_revisados, pasta_revisados)
    mover_obsoletos(obsoletos, pasta_obsoletos)
    messagebox.showinfo("Concluído", "Processo concluído com sucesso.")
    sys.exit(0)

def janela_selecao_projeto():
    root = tk.Tk()
    root.title("Selecionar Projeto")
    root.geometry("600x400")
    projetos_dict = carregar_projetos()
    style = ttkbootstrap.Style(theme="flatly")

    if not projetos_dict:
        messagebox.showerror("Erro", "Nenhum projeto encontrado ou erro no arquivo JSON.")
        root.destroy()
        return None, None

    projetos_convertidos = []
    for numero, caminho_completo in projetos_dict.items():
        nome_projeto = os.path.basename(caminho_completo)
        projetos_convertidos.append((numero, nome_projeto, caminho_completo))

    sel = {"numero": None, "caminho": None}

    def filtrar(*args):
        termo = entrada.get().lower()
        tree.delete(*tree.get_children())
        for numero, nome_projeto, caminho_original in projetos_convertidos:
            if termo in nome_projeto.lower():
                tree.insert("", tk.END, values=(numero, nome_projeto, caminho_original))
        all_iid = tree.get_children()
        if len(all_iid) == 1:
            tree.selection_set(all_iid[0])

    def confirmar():
        sel_i = tree.selection()
        if not sel_i:
            messagebox.showinfo("Info", "Selecione um projeto.")
            return
        vals = tree.item(sel_i[0], "values")
        sel["numero"] = vals[0]
        sel["caminho"] = vals[2]
        root.destroy()
        Disciplinas_Detalhes_Projeto(sel["numero"], sel["caminho"])

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

    for numero, nome_projeto, caminho_original in projetos_convertidos:
        tree.insert("", tk.END, values=(numero, nome_projeto, caminho_original))

    btn_frame = tk.Frame(frame)
    btn_frame.pack(pady=5)

    ttk.Button(btn_frame, text="Confirmar", command=confirmar, bootstyle="success").pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="Cancelar", command=root.destroy, bootstyle="danger").pack(side=tk.LEFT, padx=5)

    root.mainloop()
    return sel["numero"], sel["caminho"]

def Disciplinas_Detalhes_Projeto(numero, caminho):
    disciplinas_path = os.path.join(caminho, "3 Desenvolvimento")
    if not os.path.exists(disciplinas_path):
        messagebox.showerror("Erro", "A pasta de disciplinas não foi encontrada no projeto selecionado.")
        return

    nova_janela = tk.Tk()
    nova_janela.title(f"Gerenciador de Projetos - Projeto {numero}")
    nova_janela.geometry("900x600")

    header = tk.Label(nova_janela, text=f"Projeto {numero} - Selecione a disciplina para entrega", font=("Helvetica", 14, "bold"), anchor="w")
    header.pack(fill=tk.X, padx=10, pady=5)

    cols = ["Nome", "Data de Modificação", "Tipo", "Tamanho"]
    tree = ttk.Treeview(nova_janela, columns=cols, show="headings", height=20)
    tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=200 if col == "Nome" else 150, anchor="w")

    disciplinas = []
    for item in os.listdir(disciplinas_path):
        full_path = os.path.join(disciplinas_path, item)
        if os.path.isdir(full_path):
            mod_time = datetime.fromtimestamp(os.path.getmtime(full_path)).strftime("%d/%m/%Y %H:%M")
            disciplinas.append((item, mod_time, "Pasta", "--"))

    for disciplina in disciplinas:
        tree.insert("", tk.END, values=disciplina)

    def confirmar_selecao_arquivos():
        selecionados = tree.selection()
        if not selecionados:
            messagebox.showwarning("Atenção", "Nenhuma disciplina selecionada.")
            return
        valores = tree.item(selecionados[0])["values"]
        disciplina_nome = valores[0]

        pasta_disciplina = os.path.join(disciplinas_path, disciplina_nome)
        if not os.path.isdir(pasta_disciplina):
            messagebox.showerror("Erro", f"A pasta da disciplina '{pasta_disciplina}' não foi encontrada.")
            return

        # Tenta corrigir eventuais variações de "1.ENTREGAS"
        expected_folder = "1.ENTREGAS"
        match_entrega = None

        def normalize(n):
            return n.lower().replace(" ", "").replace("-", "").replace("_", "").replace(".", "")

        for folder in os.listdir(pasta_disciplina):
            if normalize(folder) == normalize(expected_folder):
                match_entrega = folder
                break

        if not match_entrega:
            messagebox.showerror("Erro", f"A pasta de entrega '{expected_folder}' não foi encontrada.")
            return

        pasta_entrega = os.path.join(pasta_disciplina, match_entrega)
        if not os.path.isdir(pasta_entrega):
            messagebox.showerror("Erro", f"A pasta de entrega '{pasta_entrega}' não pôde ser acessada.")
            return

        arquivos_selecionados = filedialog.askopenfilenames(title="Selecione arquivos para entrega", initialdir=pasta_entrega)
        if not arquivos_selecionados:
            messagebox.showwarning("Atenção", "Nenhum arquivo foi selecionado.")
            return

        arquivos_processados = []
        for arquivo in arquivos_selecionados:
            nome_arquivo = os.path.basename(arquivo)
            dados_extraidos = extrair_dados_arquivo(nome_arquivo)
            dados_extraidos["caminho"] = arquivo
            arquivos_processados.append(dados_extraidos)

        if not arquivos_processados:
            messagebox.showerror("Erro", "Nenhum dado foi processado dos arquivos selecionados.")
            return

        nova_janela.destroy()
        exibir_interface_tabela(numero, arquivos_previos=arquivos_processados)

    def voltar():
        nova_janela.destroy()
        janela_selecao_projeto()

    content = tk.Frame(nova_janela, bg="#f5f5f5", padx=20, pady=20)
    content.pack(fill=tk.BOTH, expand=True)

    btn_frame = tk.Frame(nova_janela)
    btn_frame.pack(fill=tk.X, pady=5, padx=10)

    ttk.Button(btn_frame, text="Voltar", command=voltar, bootstyle="warning").pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="Confirmar Seleção", command=confirmar_selecao_arquivos, bootstyle="success").pack(side=tk.RIGHT, padx=5)

    nova_janela.mainloop()

def carregar_json(filepath: str) -> Dict[str, Any]:
    if os.path.exists(filepath):
        with open(filepath, "r", encoding="utf-8") as file:
            return json.load(file)
    else:
        return {}

def salvar_json(filepath: str, data: Dict[str, Any]):
    try:
        with open(filepath, "w", encoding="utf-8") as file:
            json.dump(data, file, indent=4, ensure_ascii=False)
    except Exception as e:
        messagebox.showerror("Erro de Salvamento", f"Ocorreu um erro ao salvar o arquivo JSON:\n{e}")

def identificar_numero_arquivo(partes):
    elementos_relevantes = partes[5:]
    for parte in elementos_relevantes:
        if any(char.isdigit() for char in parte) and not parte.isalpha() and len(parte) <= 6:
            return parte
    return ""

# Adicionamos os 14 campos. Manteremos a lógica atual, mas criamos chaves para cada um.
def extrair_dados_arquivo(nome_arquivo):
    nome_base, extensao = os.path.splitext(nome_arquivo)
    partes = nome_base.split('-')
    try:
        # Separar conjunto e nº do documento (ex: G.001)
        conjunto_raw = partes[7] if len(partes) > 7 else ""
        if '.' in conjunto_raw:
            conjunto_split = conjunto_raw.split('.')
            conjunto = conjunto_split[0]
            numero_documento = conjunto_split[1]
        else:
            conjunto = conjunto_raw
            numero_documento = ""

        dados = {
            "Status": partes[0] if len(partes) > 0 else "",
            "Cliente": partes[1] if len(partes) > 1 else "",
            "Nº do Projeto": partes[2] if len(partes) > 2 else "",
            "Organização": partes[3] if len(partes) > 3 else "",
            "Sigla da Disciplina": partes[4] if len(partes) > 4 else "",
            "Fase": partes[5] if len(partes) > 5 else "",
            "Tipo de Documento": partes[6] if len(partes) > 6 else "",

            "Conjunto": conjunto,
            
            "Nº do Documento": numero_documento,

            "Bloco": partes[8] if len(partes) > 8 else "",
            "Pavimento": partes[9] if len(partes) > 9 else "",
            "Subsistema": partes[10] if len(partes) > 10 else "",
            "Tipo do Desenho": partes[11] if len(partes) > 11 else "",
            "Revisão": partes[12] if len(partes) > 12 else "",
            # Os campos abaixo fazem parte da estrutura já existente
            "Nome do Arquivo": nome_arquivo,
            "Extensão": extensao.strip('-'),
            "Modificação": datetime.now().strftime("%d/%m/%Y"),
            "Modificado por": "Usuário"
        }
    except IndexError:
        # Se ocorrer algum problema, preencha com strings vazias para não quebrar o fluxo.
        dados = {
            "Status": "",
            "Cliente": "",
            "Nº do Projeto": "",
            "Organização": "",
            "Sigla da Disciplina": "",
            "Fase": "",
            "Tipo de Documento": "",
            "Conjunto": "",
            "Nº do Documento": "",
            "Bloco": "",
            "Pavimento": "",
            "Subsistema": "",
            "Tipo do Desenho": "",
            "Revisão": "",
            "Nome do Arquivo": nome_arquivo,
            "Extensão": extensao.strip('-'),
            "Modificação": datetime.now().strftime("%d/%m/%Y"),
            "Modificado por": "Usuário"
        }
    return dados


def parts_if_exist(parts_list, index):
    return parts_list[index] if len(parts_list) > index else ""

def exibir_interface_tabela(numero, arquivos_previos=None, caminho_projeto=None):
    janela = tk.Tk()
    janela.title(f"Gerenciador de Projetos - Projeto {numero}")
    janela.geometry("1200x800")

    frame_principal = tk.Frame(janela)
    frame_principal.pack(fill=tk.BOTH, expand=True)

    barra_lateral = tk.Frame(frame_principal, bg="#2c3e50", width=200)
    barra_lateral.pack(side=tk.LEFT, fill=tk.Y)

    lbl_titulo = tk.Label(barra_lateral, text="OAE - Engenharia", font=("Helvetica", 14, "bold"), bg="#2c3e50", fg="white")
    lbl_titulo.pack(pady=10)

    lbl_projetos = tk.Label(barra_lateral, text="PROJETOS", font=("Helvetica", 10, "bold"), bg="#34495e", fg="white", anchor="w", padx=10)
    lbl_projetos.pack(fill=tk.X, pady=5)

    lst_projetos = tk.Listbox(barra_lateral, height=5, bg="#ecf0f1", font=("Helvetica", 9))
    lst_projetos.pack(fill=tk.X, padx=10, pady=5, anchor="center")
    lst_projetos.insert(tk.END, "Projeto ", numero)

    lbl_membros = tk.Label(barra_lateral, text="MEMBROS", font=("Helvetica", 10, "bold"), bg="#34495e", fg="white", anchor="w", padx=10)
    lbl_membros.pack(fill=tk.X, pady=5)

    barra_lateral.pack_propagate(False)

    conteudo_principal = tk.Frame(frame_principal)
    conteudo_principal.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    def fazer_analise_nomenclatura():
        lista_arquivos = []
        for item in tabela.get_children():
            valores = tabela.item(item)["values"]
            nome_arquivo = valores[1]
            caminho = valores[10]
            dados_extraidos = extrair_dados_arquivo(nome_arquivo)
            dados_extraidos["caminho"] = caminho
            lista_arquivos.append(dados_extraidos)
                
        if not lista_arquivos:
            messagebox.showinfo("Aviso", "Nenhum arquivo adicionado para análise.")
        else:
            janela.destroy()
            # Enviamos a lista_arquivos para a tela de análise.
            # Vamos converter para o novo formato (com 14 campos).
            # Mas, como parte do enunciado, manteremos a mesma transição e adaptaremos no "tela_analise_nomenclatura".
            tela_analise_nomenclatura(lista_arquivos)

    lbl_instrucao = tk.Label(conteudo_principal, text="Adicionar Arquivos para Entrega", font=("Helvetica", 15, "bold"), anchor="w")
    lbl_instrucao.place(x=10, y=10)

    frm_botoes = tk.Frame(conteudo_principal)
    frm_botoes.pack(side=tk.TOP, anchor="ne", pady=10, padx=10)

    ttk.Button(frm_botoes, text="Fazer análise da Nomenclatura", command=fazer_analise_nomenclatura).pack(side=tk.LEFT, padx=5)

    # Campos originais, mas iremos manter por compatibilidade.
    cols = [
        "Status", "Nome do Arquivo", "Extensão", "Nº do Arquivo",
        "Fase", "Tipo", "Revisão", "Modificação", "Modificado por", "Entrega", "caminho"
    ]
    tabela = ttk.Treeview(conteudo_principal, columns=cols, show="headings", height=20)
    for col in cols:
        tabela.heading(col, text=col)
        if col == "Nome do Arquivo":
            tabela.column(col, width=300)
        elif col == "caminho":
            tabela.column(col, width=0, stretch=False, minwidth=0)
        else:
            tabela.column(col, width=120)

    tabela["displaycolumns"] = (
        "Status", "Nome do Arquivo", "Extensão", "Nº do Arquivo",
        "Fase", "Tipo", "Revisão", "Modificação", "Modificado por", "Entrega"
    )
    tabela.pack(fill=tk.BOTH, expand=True)

    if arquivos_previos:
        for dados_extraidos in arquivos_previos:
            tabela.insert("", tk.END, values=(
                dados_extraidos.get("Status", ""),
                dados_extraidos.get("Nome do Arquivo", ""),
                dados_extraidos.get("Extensão", ""),
                dados_extraidos.get("Nº do Arquivo", ""),
                dados_extraidos.get("Fase", ""),
                dados_extraidos.get("Tipo", ""),
                dados_extraidos.get("Revisão", ""),
                dados_extraidos.get("Modificação", ""),
                dados_extraidos.get("Modificado por", ""),
                dados_extraidos.get("Entrega", ""),
                dados_extraidos.get("caminho", "")
            ))

    def adicionar_arquivos():
        arquivos = filedialog.askopenfilenames(title="Selecione arquivos")
        for arquivo in arquivos:
            nome_arquivo = os.path.basename(arquivo)
            dados_extraidos = extrair_dados_arquivo(nome_arquivo)
            dados_extraidos["caminho"] = arquivo
            tabela.insert("", tk.END, values=(
                dados_extraidos.get("Status", ""),
                dados_extraidos.get("Nome do Arquivo", ""),
                dados_extraidos.get("Extensão", ""),
                dados_extraidos.get("Nº do Arquivo", ""),
                dados_extraidos.get("Fase", ""),
                dados_extraidos.get("Tipo", ""),
                dados_extraidos.get("Revisão", ""),
                dados_extraidos.get("Modificação", ""),
                dados_extraidos.get("Modificado por", ""),
                dados_extraidos.get("Entrega", ""),
                dados_extraidos.get("caminho", "")
            ))

    def remover_arquivo():
        selecionados = tabela.selection()
        if selecionados:
            for item in selecionados:
                tabela.delete(item)
        else:
            messagebox.showinfo("Informação", "Nenhum item selecionado.")

    def voltar():
        janela.destroy()
        Disciplinas_Detalhes_Projeto(numero, caminho_projeto if caminho_projeto else "")

    btn_frame = tk.Frame(conteudo_principal)
    btn_frame.pack(side=tk.LEFT, pady=10, padx=10)

    ttk.Button(btn_frame, text="Adicionar Arquivo", command=adicionar_arquivos).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="Remover Arquivo", command=remover_arquivo).pack(side=tk.LEFT, padx=5)

    btn_frame2 = tk.Frame(conteudo_principal)
    btn_frame2.pack(side="bottom", anchor="e", pady=5, padx=10)

    ttk.Button(btn_frame2, text="Voltar", command=voltar).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame2, text="Sair", command=janela.destroy).pack(side=tk.RIGHT, padx=5)

def identificar_revisoes(lista_arquivos):
    # Mantemos a lógica de identificação de revisões do código original.
    grupos = {}
    for arq in lista_arquivos:
        nome_base, _ = os.path.splitext(arq["Nome do Arquivo"])
        tokens = nome_base.split("-")
        if len(tokens) < 2:
            continue
        identificador = "-".join(tokens[:-1])
        revisao = tokens[-1] if tokens[-1].startswith("R") and tokens[-1][1:].isdigit() else "R00"
        if identificador not in grupos:
            grupos[identificador] = []
        grupos[identificador].append((revisao, arq))

    arquivos_revisados = []
    arquivos_obsoletos = []
    for identificador, arquivos in grupos.items():
        arquivos.sort(key=lambda x: int(x[0][1:]) if x[0][1:].isdigit() else 0)
        revisao_mais_recente = arquivos[-1][1]
        arquivos_revisados.append(revisao_mais_recente)
        arquivos_obsoletos.extend([a[1] for a in arquivos[:-1]])
    return arquivos_revisados, arquivos_obsoletos

############################
# TELA DE ANÁLISE DE NOMENCLATURA
############################
def tela_analise_nomenclatura(lista_arquivos):
    # Vamos mapear a correspondência entre as chaves do dicionário e os rótulos amigáveis.
    # Os 14 campos, na ordem do "extrair_dados_arquivo":
    friendly_fields = [
        ("Status", "Status"),
        ("Cliente", "Cliente"),
        ("Nº do Projeto", "Nº do Projeto"),
        ("Organização", "Organização"),
        ("Sigla da Disciplina", "Sigla da Disciplina"),
        ("Fase do Projeto", "Fase do Projeto"),
        ("Tipo do Documento", "Tipo do Documento"),
        ("Conjunto", "Conjunto"),
        ("Nº do Documento", "Nº do Documento"),
        ("Bloco", "Bloco"),
        ("Pavimento", "Pavimento"),
        ("Subsistema", "Subsistema"),
        ("Tipo do Desenho", "Tipo do Desenho"),
        ("Revisão", "Revisão"),
    ]

    CAMPOS_NOMENCLATURA = [
        "Status",
        "Cliente",
        "Nº do Projeto",
        "Organização",
        "Sigla da Disciplina",
        "Fase",
        "Tipo de Documento",
        "Conjunto",
        "Nº do Documento",
        "Bloco",
        "Pavimento",
        "Subsistema",
        "Tipo do Desenho",
        "Revisão"
    ]


    config_nomenclatura = carregar_config_nomenclatura()
    conf_revisao = config_nomenclatura.get("config_revisao", {})
    conf_valores = config_nomenclatura.get("config_valores", {})

    janela = tk.Tk()
    janela.title("Verificação de Nomenclatura")
    janela.geometry("1400x800")

    lbl_instrucao = tk.Label(
        janela,
        text="Confira a nomenclatura. Clique no nome se desejar editar campos. Corrija antes de avançar, se necessário.",
        font=("Helvetica", 12)
    )
    lbl_instrucao.pack(pady=MARGIN_SIZE)

    frm_main = tk.Frame(janela)
    frm_main.pack(fill=tk.BOTH, expand=True)

    canvas = tk.Canvas(frm_main)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(frm_main, orient="vertical", command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    canvas.configure(yscrollcommand=scrollbar.set)
    container = tk.Frame(canvas, bg="#ECE2E2")
    container_id = canvas.create_window((0, 0), window=container, anchor="n")

    def on_canvas_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas_width = event.width
        canvas.itemconfig(container_id, width=canvas_width)

    canvas.bind("<Configure>", on_canvas_configure)
    fields_in_error = set()
    expanded_cards = set()

    def expand_or_collapse(card_id):
        if card_id in expanded_cards:
            expanded_cards.remove(card_id)
        else:
            expanded_cards.add(card_id)
        render_cards()

    def validate_with_rules(campo, typed_value):
        # Campo refere-se a "Status", "Fase", etc. Se houver alguma regra, consultamos NOMENCLATURA_RULES.
        if campo in NOMENCLATURA_RULES:
            if typed_value not in NOMENCLATURA_RULES[campo]:
                return False
        return True

    def validate_entry(entry_var, campo, error_label):
        typed_value = entry_var.get().strip()
        # Se quisermos usar config_valores, poderíamos verificar:
        valor_certo = conf_valores.get(campo, {}).get("valor_aceito", "")
        msg_erro = conf_valores.get(campo, {}).get("mensagem_erro", "")

        # Se valor_certo estiver definido e typed_value != valor_certo, erro
        if valor_certo and typed_value != valor_certo:
            fields_in_error.add((campo, entry_var))
            error_label.config(text=msg_erro if msg_erro else "Valor não corresponde ao esperado.", fg="red")
            return

        # Checamos a regra geral do NOMENCLATURA_RULES
        if not typed_value:
            # Vazio é inválido, de acordo com a lógica do code original (bloqueia avanço).
            fields_in_error.add((campo, error_label))
            error_label.config(text="Campo obrigatório.", fg="red")
        else:
            # Verificamos se existe alguma restrição:
            if not validate_with_rules(campo, typed_value):
                fields_in_error.add((campo, entry_var))
                error_label.config(
                    text=f"O valor '{typed_value}' não consta em NOMENCLATURA_RULES para '{campo}'.",
                    fg="red"
                )
            else:
                # Se passou na validação, removemos do set de errors (caso estivesse).
                fields_in_error.discard((campo, entry_var))
                error_label.config(text="", fg="red")

    # Converter a lista_arquivos anterior para o novo modelo de 14 campos,
    # caso ainda não possua. Se já tiver, preserva.
    def ensure_14_fields(arquivo):
        # Já definimos a base. Se algum campo não existe, definimos vazio.
        for field_key, _friendly in friendly_fields:
            if field_key not in arquivo:
                arquivo[field_key] = ""
        return arquivo

    lista_arquivos = [ensure_14_fields(arq) for arq in lista_arquivos]

    def render_cards():

        for widget in container.winfo_children():
            widget.destroy()

        for idx, arq in enumerate(lista_arquivos):
            card_id = idx
            card_frame = tk.Frame(container, bd=1, relief=tk.RIDGE, padx=MARGIN_SIZE, pady=MARGIN_SIZE, bg="#ECE2E2")
            card_frame.pack(padx=MARGIN_SIZE, pady=MARGIN_SIZE, fill=tk.X)

            # Usamos o campo "Nome do Arquivo" como título. Se não existir, fallback em "Status" ou algo assim.
            titulo = arq.get("Nome do Arquivo", f"Arquivo {idx+1}")
            header_button = ttk.Button(card_frame, text=titulo, command=lambda cid=card_id: expand_or_collapse(cid))
            header_button.pack(fill=tk.X)

            if card_id in expanded_cards:
                detail_frame = tk.Frame(card_frame, bd=1, relief=tk.GROOVE, padx=MARGIN_SIZE, pady=MARGIN_SIZE, bg="#ECE2E2")
                detail_frame.pack(fill=tk.X)

                # Exibiremos os 14 campos amigáveis numa grade.
                row_max = 7  # 7 linhas
                col_max = 2  # 2 colunas => total 14 posições
                idx_campos = 0

                CAMPOS_NOMENCLATURA = [
                    "Status", "Cliente", "Nº do Projeto", "Organização", "Sigla da Disciplina",
                    "Fase", "Tipo de Documento", "Conjunto", "Nº do Documento", "Bloco",
                    "Pavimento", "Subsistema", "Tipo do Desenho", "Revisão"
                ]

                total_campos = len(CAMPOS_NOMENCLATURA)
                linhas = (total_campos + 1) // 2  # duas colunas por linha

                for rr in range(linhas):
                    for cc in range(2):
                        idx = rr * 2 + cc
                        if idx < total_campos:
                            nome_campo = CAMPOS_NOMENCLATURA[idx]
                            valor_campo = arq.get(nome_campo, "")
                            c_frame = tk.Frame(detail_frame, bd=1, relief=tk.FLAT, bg="#ECE2E2")
                            c_frame.grid(row=rr, column=cc, padx=MARGIN_SIZE, pady=MARGIN_SIZE, sticky="nsew")

                            lab = tk.Label(c_frame, text=nome_campo + ":", bg="#ECE2E2")
                            lab.pack(anchor="w")

                            val_var = tk.StringVar(value=valor_campo)
                            error_lbl = tk.Label(c_frame, text="", fg="red", bg="#ECE2E2")

                            def callback(svar=val_var, c=nome_campo, el=error_lbl):
                                validate_entry(svar, c, el)

                            val_entry = tk.Entry(c_frame, textvariable=val_var, width=30)
                            val_entry.bind("<KeyRelease>", lambda e, cb=callback: cb())
                            val_entry.pack(fill=tk.X)
                            error_lbl.pack(anchor="w")

                for rr in range(row_max):
                    detail_frame.rowconfigure(rr, weight=0)
                for cc in range(col_max):
                    detail_frame.columnconfigure(cc, weight=1)

    render_cards()

    def ajustar_canvas():
        janela.update_idletasks()
        width_atual = canvas.winfo_width()
        canvas.itemconfig(container_id, width=width_atual)
        canvas.configure(scrollregion=canvas.bbox("all"))

    janela.after(100, ajustar_canvas)

    def avancar():
        # Ao avançar, verificamos se há campos em erro. Se houver, não deixa prosseguir.
        if fields_in_error:
            messagebox.showerror("Erro", "Existem campos obrigatórios vazios ou inválidos. Corrija antes de prosseguir.")
        else:
            janela.destroy()
            tela_verificacao_revisao(lista_arquivos)

    def voltar():
        janela.destroy()
        # Retorna para a lista, se necessário
        exibir_interface_tabela("467", lista_arquivos)

    btn_frame = tk.Frame(janela)
    btn_frame.pack(side="bottom", anchor="e", pady=MARGIN_SIZE, padx=MARGIN_SIZE)

    ttk.Button(btn_frame, text="Voltar", command=voltar).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="Avançar", command=avancar).pack(side=tk.RIGHT, padx=5)

    janela.mainloop()

############################
# TELA DE VERIFICAÇÃO DE REVISÃO
############################
def tela_verificacao_revisao(lista_arquivos):
    arquivos_revisados, arquivos_obsoletos = identificar_revisoes(lista_arquivos)

    janela = tk.Tk()
    janela.title("Verificação de Revisão")
    janela.geometry("1000x700")

    lbl_instrucao = tk.Label(janela, text="Confira os arquivos revisados e obsoletos antes da entrega.")
    lbl_instrucao.pack(pady=10)

    frame_revisados = tk.LabelFrame(janela, text="Arquivos Revisados")
    frame_revisados.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    tree_revisados = ttk.Treeview(frame_revisados, columns=["Nome do Arquivo", "Revisão"], show="headings", height=10)
    tree_revisados.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    tree_revisados.heading("Nome do Arquivo", text="Nome do Arquivo")
    tree_revisados.heading("Revisão", text="Revisão")

    for arq in arquivos_revisados:
        tree_revisados.insert("", tk.END, values=(arq.get("Nome do Arquivo", ""), arq.get("Revisão", "")))

    frame_obsoletos = tk.LabelFrame(janela, text="Arquivos Obsoletos")
    frame_obsoletos.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    tree_obsoletos = ttk.Treeview(frame_obsoletos, columns=["Nome do Arquivo", "Revisão"], show="headings", height=10)
    tree_obsoletos.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    tree_obsoletos.heading("Nome do Arquivo", text="Nome do Arquivo")
    tree_obsoletos.heading("Revisão", text="Revisão")

    for arq in arquivos_obsoletos:
        tree_obsoletos.insert("", tk.END, values=(arq.get("Nome do Arquivo", ""), arq.get("Revisão", "")))

    def voltar():
        janela.destroy()
        tela_analise_nomenclatura(lista_arquivos)

    def confirmar():
        messagebox.showinfo("Confirmação", "Arquivos revisados e obsoletos identificados com sucesso.")
        pasta_revisados, pasta_obsoletos = criar_pastas_organizacao()
        mover_arquivos(arquivos_revisados, pasta_revisados)
        mover_obsoletos(arquivos_obsoletos, pasta_obsoletos)
        janela.destroy()

    bf = tk.Frame(janela)
    bf.pack(side="bottom", anchor="e", pady=5, padx=10)

    ttk.Button(bf, text="Voltar", command=voltar).pack(side=tk.LEFT, padx=5)
    ttk.Button(bf, text="Confirmar", command=confirmar).pack(side=tk.RIGHT, padx=5)

    # Como pedido, salvamos tudo em formato fixo (JSON).
    # Vamos usar a mesma estrutura do code original.
    salvar_json("saida_final_nomenclaturas.json", {"arquivos": lista_arquivos})

    janela.mainloop()

if __name__ == "__main__":
    # Mantido como no original, apenas chamando a interface principal.
    exibir_interface_tabela("467")
