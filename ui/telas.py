import os
import tkinter as tk
import json
import shutil
import sys
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from typing import Dict, Any
import ttkbootstrap
from ttkbootstrap.style import Style
from ttkbootstrap.constants import *

MARGIN_SIZE = 10

# Caminhos e JSONs
JSON_FILE_PATH = "dados_projetos.json"
PROJETOS_JSON = r"G:\Drives compartilhados\OAE-JSONS\diretorios_projetos.json"
ULTIMO_DIRETORIO_JSON = "ultimo_diretorio.json"
HISTORICO_JSON = "historico_arquivos.json"

# Novo JSON de regras de nomenclatura
NOMENCLATURA_REGRAS_JSON = "config/nomenclatura_regras.json"

# Nosso novo arquivo de cache de subdiretórios
PASTAS_PROJETOS_JSON = "pastas_projetos.json"

###############################################################################
# Carrega as regras de nomenclatura a partir do novo JSON
###############################################################################
def carregar_regras_nomenclatura():
    if not os.path.exists(NOMENCLATURA_REGRAS_JSON):
        print("[ERRO] Arquivo de regras não encontrado:", NOMENCLATURA_REGRAS_JSON)
        return {}
    try:
        with open(NOMENCLATURA_REGRAS_JSON, "r", encoding="utf-8") as f:
            regras = json.load(f)
            print("[DEBUG] Regras carregadas:", regras)  # <---- DEBUG
            return regras
    except json.JSONDecodeError as e:
        print("[ERRO] Falha ao ler JSON de regras:", e)
        return {}

###############################################################################
# NOVA FUNÇÃO: Atualiza o cache de subdiretórios em PASTAS_PROJETOS_JSON
# e sincroniza com PROJETOS_JSON se houver mudanças
###############################################################################
def atualizar_cache_projetos():
    """
    1. Lê o cache existente em pastas_projetos.json (se existir).
    2. Faz a leitura real dos subdiretórios em G:/Drives compartilhados/.
    3. Se houver diferença, salva novamente o cache e sobrescreve PROJETOS_JSON
       para manter compatibilidade com o restante do código.
    4. Retorna um dicionário no formato {numeroProjeto: caminhoCompleto}, 
       igual ao antigo PROJETOS_JSON, mas atualizado dinamicamente.
    """

    cache_antigo = {}
    if os.path.exists(PASTAS_PROJETOS_JSON):
        try:
            with open(PASTAS_PROJETOS_JSON, "r", encoding="utf-8") as f:
                cache_antigo = json.load(f)
        except:
            cache_antigo = {}

    # Subdiretórios do G:/Drives compartilhados/
    raiz_projetos = r"G:\Drives compartilhados"
    diretorios_atuais = {}
    try:
        # Vamos numerar (1,2,3...) ou criar alguma lógica de chave
        idx = 1
        for item in os.listdir(raiz_projetos):
            full_path = os.path.join(raiz_projetos, item)
            if os.path.isdir(full_path):
                # Ex.: { "1": "G:\\Drives compartilhados\\Projeto1" }
                if item.upper().startswith("OAE"):
                    diretorios_atuais[str(idx)] = full_path
                    idx += 1
    except Exception as e:
        print("[ERRO] Falha ao ler G:/Drives compartilhados/:", e)
        return cache_antigo  # Retornamos o que tinha

    # Compara se houve mudança
    if diretorios_atuais != cache_antigo:
        print("[DEBUG] Detectamos mudança real na pasta => Atualizando JSON de cache.")
        # Salva em pastas_projetos.json
        with open(PASTAS_PROJETOS_JSON, "w", encoding="utf-8") as f:
            json.dump(diretorios_atuais, f, indent=4, ensure_ascii=False)

        # Também sobrescreve PROJETOS_JSON para manter fluxo antigo
        with open(PROJETOS_JSON, "w", encoding="utf-8") as f:
            json.dump(diretorios_atuais, f, indent=4, ensure_ascii=False)

        return diretorios_atuais
    else:
        print("[DEBUG] Nenhuma mudança detectada, reaproveitando cache existente.")
        if not os.path.exists(PROJETOS_JSON):
            with open(PROJETOS_JSON, "w", encoding="utf-8") as f:
                json.dump(cache_antigo, f, indent=4, ensure_ascii=False)
        return cache_antigo

###############################################################################
# Funções para carregar outros dados do sistema
###############################################################################
def carregar_projetos():
    """
    Antes de carregar do PROJETOS_JSON, chamamos atualizar_cache_projetos()
    para garantir que o PROJETOS_JSON esteja sincronizado com a pasta real.
    Depois, lemos PROJETOS_JSON normalmente, preservando toda a lógica antiga.
    """
    # Chama a função que atualiza cache e reescreve PROJETOS_JSON se necessário
    atualizar_cache_projetos()

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

###############################################################################
# Criação de pastas para revisados e obsoletos (mantido da versão anterior)
###############################################################################
def criar_pastas_organizacao():
    base_dir = r"C:\Users\DEV\Downloads\OAE-467 - PETER-KD-ENG"
    pasta_revisados = os.path.join(base_dir, "Revisados")
    if not os.path.exists(pasta_revisados):
        os.makedirs(pasta_revisados)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    pasta_obsoletos = os.path.join(base_dir, f"Obsoleto_{timestamp}")
    if not os.path.exists(pasta_obsoletos):
        os.makedirs(pasta_obsoletos)
    return pasta_revisados, pasta_obsoletos

###############################################################################
# Movimentação de arquivos
###############################################################################
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

###############################################################################
# Seleção de Projeto (mantida original, mas agora projetos_dict já atualizado)
###############################################################################
def janela_selecao_projeto():
    root = tk.Tk()
    root.title("Selecionar Projeto")
    root.geometry("600x400")
    projetos_dict = carregar_projetos()  # Vai chamar atualizar_cache_projetos()
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

###############################################################################
# Disciplinas
###############################################################################
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

        # Monta caminho da disciplina
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
        exibir_interface_tabela(numero, arquivos_previos=arquivos_processados, caminho_projeto=caminho)

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

###############################################################################
# Salvamento e carregamento de JSON (histórico, etc.)
###############################################################################
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
        messagebox.showerror("Erro", f"Falha ao salvar dados em JSON: {e}")

###############################################################################
# Extrair dados do arquivo, adaptado para 14 campos, dividindo G.001 etc.
###############################################################################
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

###############################################################################
# Interface que exibe a lista de arquivos selecionados e chama análise
###############################################################################
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
            # Re-extrair a nomenclatura, garantindo todos os 14 campos
            dados_extraidos = extrair_dados_arquivo(nome_arquivo)
            dados_extraidos["caminho"] = caminho
            lista_arquivos.append(dados_extraidos)

        if not lista_arquivos:
            messagebox.showinfo("Aviso", "Nenhum arquivo adicionado para análise.")
        else:
            janela.destroy()
            tela_analise_nomenclatura(lista_arquivos)

    lbl_instrucao = tk.Label(conteudo_principal, text="Adicionar Arquivos para Entrega", font=("Helvetica", 15, "bold"), anchor="w")
    lbl_instrucao.place(x=10, y=10)

    frm_botoes = tk.Frame(conteudo_principal)
    frm_botoes.pack(side=tk.TOP, anchor="ne", pady=10, padx=10)

    ttk.Button(frm_botoes, text="Fazer análise da Nomenclatura", command=fazer_analise_nomenclatura).pack(side=tk.LEFT, padx=5)

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
                dados_extraidos.get("N° do Arquivo", ""),  # compatibilidade com antigo
                dados_extraidos.get("Fase", ""),
                dados_extraidos.get("Tipo de Documento", ""),  # adaptado
                dados_extraidos.get("Revisão", ""),
                dados_extraidos.get("Modificação", ""),
                dados_extraidos.get("Modificado por", ""),
                "",  # Entrega não está sendo usada
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
                dados_extraidos.get("N° do Arquivo", ""),
                dados_extraidos.get("Fase", ""),
                dados_extraidos.get("Tipo de Documento", ""),
                dados_extraidos.get("Revisão", ""),
                dados_extraidos.get("Modificação", ""),
                dados_extraidos.get("Modificado por", ""),
                "",
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

###############################################################################
# Identificação de revisões e telas finais
###############################################################################
def identificar_revisoes(lista_arquivos):
    grupos = {}
    for arq in lista_arquivos:
        nome_base, _ = os.path.splitext(arq["Nome do Arquivo"])
        tokens = nome_base.split("-")
        if len(tokens) < 2:
            continue
        # Último token = revisao (R01). Se não começar com R, assume R00
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

###############################################################################
# Tela de análise de nomenclatura com validação via JSON
###############################################################################
def tela_analise_nomenclatura(lista_arquivos):
    # Carrega as regras do JSON
    regras_nomenclatura = carregar_regras_nomenclatura()

    janela = tk.Tk()
    janela.title("Verificação de Nomenclatura")
    janela.geometry("1400x800")

    lbl_instrucao = tk.Label(
        janela,
        text="Confira a nomenclatura. Clique no nome para editar campos. Corrija antes de avançar, se necessário.",
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

    # Lista dos 14 campos em ordem
    CAMPOS_NOMENCLATURA = [
        "Status", "Cliente", "N° do Projeto", "Organização", "Sigla da Disciplina",
        "Fase", "Tipo de Documento", "Conjunto", "N° do Documento", "Bloco",
        "Pavimento", "Subsistema", "Tipo do Desenho", "Revisão"
    ]

    def expand_or_collapse(card_id):
        if card_id in expanded_cards:
            expanded_cards.remove(card_id)
        else:
            expanded_cards.add(card_id)
        render_cards()

    def validate_with_json_rules(campo, typed_value):
        # Se campo não existe no JSON, retorna (considerar livre?)
        if campo not in regras_nomenclatura:
            return (True, "")  # sem erro

        regra = regras_nomenclatura[campo]
        typed_value = typed_value.strip()

        # 1. Checa obrigatoriedade
        if regra.get("obrigatorio", False) and not typed_value:
            return (False, "Campo obrigatório.")

        # 2. Se tiver valores_permitidos
        valores_ok = regra.get("valores_permitidos", [])
        if valores_ok:
            if typed_value not in valores_ok and typed_value != "":
                return (False, f"Valor inválido para {campo}. Opções: {valores_ok}")

        # 3. Se tiver valor_fixo
        valor_fixo = regra.get("valor_fixo")
        if valor_fixo and typed_value != valor_fixo:
            return (False, f"Valor fixo exigido: {valor_fixo}")

        # 4. Se tiver regex
        import re
        pattern = regra.get("regex")
        if pattern and typed_value != "":
            if not re.match(pattern, typed_value):
                return (False, f"O valor '{typed_value}' não atende ao regex: {pattern}")

        # Se passou em tudo, está válido
        return (True, "")

    def validate_entry(entry_var, campo, error_label):
        typed_value = entry_var.get().strip()

        # Aplica regras do JSON
        is_valid, msg_erro = validate_with_json_rules(campo, typed_value)

        if not is_valid:
            error_label.config(text=msg_erro, fg="red")
            fields_in_error.add((campo, error_label))
        else:
            error_label.config(text="", fg="red")
            # Se estava em erro, remove
            if (campo, error_label) in fields_in_error:
                fields_in_error.remove((campo, error_label))

    def render_cards():
        for widget in container.winfo_children():
            widget.destroy()

        for idx, arq in enumerate(lista_arquivos):
            card_id = idx
            card_frame = tk.Frame(container, bd=1, relief=tk.RIDGE, padx=MARGIN_SIZE, pady=MARGIN_SIZE, bg="#ECE2E2")
            card_frame.pack(padx=MARGIN_SIZE, pady=MARGIN_SIZE, fill=tk.X)

            header_button = ttk.Button(card_frame, text=arq["Nome do Arquivo"], command=lambda cid=card_id: expand_or_collapse(cid))
            header_button.pack(fill=tk.X)

            if card_id in expanded_cards:
                detail_frame = tk.Frame(card_frame, bd=1, relief=tk.GROOVE, padx=MARGIN_SIZE, pady=MARGIN_SIZE, bg="#ECE2E2")
                detail_frame.pack(fill=tk.X)

                row_max = 7
                col_max = 2
                idx_campos = 0

                def make_callback(var, c, el):
                    return lambda e: validate_entry(var, c, el)

                # Monta tuples (campo amigável, valor)
                campos_arq = []
                for nome_campo in CAMPOS_NOMENCLATURA:
                    campos_arq.append((nome_campo, arq.get(nome_campo, "")))

                for rr in range(row_max):
                    for cc in range(col_max):
                        if idx_campos < len(campos_arq):
                            nome_campo, valor_campo = campos_arq[idx_campos]
                            c_frame = tk.Frame(detail_frame, bd=1, relief=tk.FLAT, bg="#ECE2E2")
                            c_frame.grid(row=rr, column=cc, padx=MARGIN_SIZE, pady=MARGIN_SIZE, sticky="nsew")

                            lab = tk.Label(c_frame, text=nome_campo + ":", bg="#ECE2E2")
                            lab.pack(anchor="w")

                            val_var = tk.StringVar(value=valor_campo)
                            error_lbl = tk.Label(c_frame, text="", fg="red", bg="#ECE2E2")

                            entry = tk.Entry(c_frame, textvariable=val_var, width=30)
                            entry.bind("<KeyRelease>", make_callback(val_var, nome_campo, error_lbl))
                            entry.pack(fill=tk.X)
                            error_lbl.pack(anchor="w")

                            # Atualiza o valor em arq quando muda
                            def update_dict_value(*args, ac=arq, field=nome_campo, v=val_var):
                                ac[field] = v.get()

                            val_var.trace("w", update_dict_value)

                            idx_campos += 1

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
        # Verifica se há campos em erro
        if fields_in_error:
            messagebox.showerror("Erro", "Existem campos vazios ou inválidos. Corrija antes de prosseguir.")
        else:
            janela.destroy()
            tela_verificacao_revisao(lista_arquivos)

    def voltar():
        janela.destroy()
        exibir_interface_tabela("467", lista_arquivos)

    btn_frame = tk.Frame(janela)
    btn_frame.pack(side="bottom", anchor="e", pady=MARGIN_SIZE, padx=MARGIN_SIZE)

    ttk.Button(btn_frame, text="Voltar", command=voltar).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="Avançar", command=avancar).pack(side=tk.RIGHT, padx=5)

    janela.mainloop()

def criar_arquivo_excel(diretorio, arquivos_novos, arquivos_revisados, arquivos_obsoletos):
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    now = datetime.now()
    data_str = now.strftime("%d_%m_%Y")
    hora_str = now.strftime("%Hh%Mmin")
    nome_arquivo = f"GRD {data_str}_{hora_str}.xlsx"
    caminho_excel = os.path.join(diretorio, nome_arquivo)

    wb = Workbook()
    ws = wb.active
    ws.title = "GRD"

    bold_font = Font(bold=True)

    # Cabeçalho
    ws.append(["OLIVEIRA ARAÚJO ENGENHARIA"])
    ws.append(["Lista de arquivos de projetos entregues com controle de versão."])
    ws.append(["Diretório:", diretorio])
    ws.append(["Data de emissão:", f"{data_str}_{hora_str}"])

    # Seção Revisados
    ws.append(["Arquivo Revisado"])
    if arquivos_revisados:
        ws.append(["Nome do arquivo", "Revisão", "Data de modificação"])
        for arq in arquivos_revisados:
            ws.append([arq["Nome do Arquivo"], arq["Revisão"], arq.get("Modificação", "")])
    else:
        ws.append(["Nenhum arquivo revisado"])

    # Seção Novos
    ws.append([])
    ws.append(["Arquivo Novo"])
    if arquivos_novos:
        ws.append(["Nome do arquivo", "Revisão", "Data de modificação"])
        for arq in arquivos_novos:
            ws.append([arq["Nome do Arquivo"], arq["Revisão"], arq.get("Modificação", "")])
    else:
        ws.append(["Nenhum arquivo novo"])

    # Seção Obsoletos
    ws.append([])
    ws.append(["Arquivo Obsoleto"])
    if arquivos_obsoletos:
        ws.append(["Nome do arquivo", "Revisão", "Data de modificação"])
        for arq in arquivos_obsoletos:
            ws.append([arq["Nome do Arquivo"], arq["Revisão"], arq.get("Modificação", "")])
    else:
        ws.append(["Nenhum arquivo obsoleto"])

    # Negrito para cabeçalhos
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            if isinstance(cell.value, str) and (
                "OLIVEIRA" in cell.value or
                "Lista de arquivos" in cell.value or
                "Diretório:" in cell.value or
                "Data de emissão:" in cell.value or
                "Arquivo Revisado" in cell.value or
                "Arquivo Novo" in cell.value or
                "Arquivo Obsoleto" in cell.value or
                cell.value in ["Nome do arquivo", "Revisão", "Data de modificação"]
            ):
                cell.font = bold_font

    # Ajusta larguras
    for col in ws.columns:
        max_length = 0
        column = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column].width = max_length + 2

    wb.save(caminho_excel)
    return caminho_excel

###############################################################################
# Tela de verificação final de revisão
###############################################################################
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
        tree_revisados.insert("", tk.END, values=(arq["Nome do Arquivo"], arq["Revisão"]))

    frame_obsoletos = tk.LabelFrame(janela, text="Arquivos Obsoletos")
    frame_obsoletos.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    tree_obsoletos = ttk.Treeview(frame_obsoletos, columns=["Nome do Arquivo", "Revisão"], show="headings", height=10)
    tree_obsoletos.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    tree_obsoletos.heading("Nome do Arquivo", text="Nome do Arquivo")
    tree_obsoletos.heading("Revisão", text="Revisão")

    for arq in arquivos_obsoletos:
        tree_obsoletos.insert("", tk.END, values=(arq["Nome do Arquivo"], arq["Revisão"]))

    def voltar():
        janela.destroy()
        tela_analise_nomenclatura(lista_arquivos)

    def confirmar():
        # Salva dados de revisão em JSON
        dados_para_salvar = {
            "arquivos_revisados": arquivos_revisados,
            "arquivos_obsoletos": arquivos_obsoletos
        }
        try:
            salvar_json("resultado_revisao.json", dados_para_salvar)
        except Exception:
            messagebox.showerror("Erro", "Falha ao salvar dados em JSON.")

        # Cria pastas
        pasta_revisados, pasta_obsoletos = criar_pastas_organizacao()

        # Move arquivos
        mover_arquivos(arquivos_revisados, pasta_revisados)
        mover_obsoletos(arquivos_obsoletos, pasta_obsoletos)

        # Gera a planilha GRD no mesmo diretório da pasta revisados
        try:
            caminho_planilha = criar_arquivo_excel(
                diretorio=pasta_revisados,
                arquivos_novos=[],  # Pode atualizar futuramente com lista de arquivos novos
                arquivos_revisados=arquivos_revisados,
                arquivos_obsoletos=arquivos_obsoletos
            )
            messagebox.showinfo("Planilha Gerada", f"GRD salva com sucesso:\n{caminho_planilha}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar a planilha GRD:\n{e}")

        janela.destroy()

    bf = tk.Frame(janela)
    bf.pack(side="bottom", anchor="e", pady=5, padx=10)

    ttk.Button(bf, text="Voltar", command=voltar).pack(side=tk.LEFT, padx=5)
    ttk.Button(bf, text="Confirmar", command=confirmar).pack(side=tk.RIGHT, padx=5)

    janela.mainloop()

###############################################################################
# MAIN
###############################################################################
if __name__ == "__main__":
    exibir_interface_tabela("467")
