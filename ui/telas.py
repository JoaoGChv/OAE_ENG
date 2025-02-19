import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from typing import Dict, Any
import ttkbootstrap
from ttkbootstrap.style import Style
from ttkbootstrap.constants import *


# Caminho do arquivo JSON
JSON_FILE_PATH = "dados_projetos.json"

# Caminho para o arquivo JSON
PROJETOS_JSON = r"G:\Drives compartilhados\OAE-JSONS\diretorios_projetos.json"
ULTIMO_DIRETORIO_JSON = "ultimo_diretorio.json"  # Arquivo para armazenar o último diretório acessado
HISTORICO_JSON = "historico_arquivos.json"  # Arquivo JSON para o histórico

# Caminho para o JSON com padrões de nomenclatura
PADROES_JSON = r"G:\Drives compartilhados\OAE - SCRIPTS\SCRIPTS\tmp_joaoG\Melhorias\Código_reformulado_teste\ui\padrões.json"

def carregar_padroes_nomenclatura():
    """
    Carrega os padrões de nomenclatura do arquivo JSON.
    Retorna uma lista de padrões, onde cada padrão é uma lista de tokens.
    """
    if os.path.exists(PADROES_JSON):
        with open(PADROES_JSON, 'r', encoding='utf-8') as f:
            return json.load(f).get("padroes", [])
    else:
        return []
    
# Função para carregar os projetos do arquivo JSON
def carregar_projetos():
    if not os.path.exists(PROJETOS_JSON):
        messagebox.showerror("Erro", f"Arquivo não encontrado: {PROJETOS_JSON}")
        return {}
    try:
        with open(PROJETOS_JSON, 'r', encoding='utf-8') as f:
            projetos = json.load(f)
            return projetos
    except json.JSONDecodeError as e:
        messagebox.showerror("Erro", f"Erro ao decodificar o JSON: {e}")
        return {}

# Função para carregar o último diretório acessado
def carregar_ultimo_diretorio():
    if os.path.exists(ULTIMO_DIRETORIO_JSON):
        with open(ULTIMO_DIRETORIO_JSON, 'r', encoding='utf-8') as f:
            return json.load(f).get("ultimo_diretorio", os.getcwd())
    return os.getcwd()

# Função para salvar o último diretório acessado
def salvar_ultimo_diretorio(diretorio):
    with open(ULTIMO_DIRETORIO_JSON, 'w', encoding='utf-8') as f:
        json.dump({"ultimo_diretorio": diretorio}, f)

# Função para atualizar o histórico de arquivos e determinar números
def atualizar_historico(lista_arquivos, caminho_json=HISTORICO_JSON):
    historico = {}
    if os.path.exists(caminho_json):
        with open(caminho_json, "r", encoding="utf-8") as f:
            historico = json.load(f)
    
    # Verificar mudanças e atualizar histórico
    for arquivo in lista_arquivos:
        data_modificacao = os.path.getmtime(arquivo)
        if arquivo not in historico or historico[arquivo]["data"] != data_modificacao:
            historico[arquivo] = {
                "numero": len(historico) + 1,
                "data": data_modificacao,
                "status": "Atual"
            }
    
    # Marcar obsoletos
    mais_recente = max(historico, key=lambda x: historico[x]["data"])
    for arq in historico:
        historico[arq]["status"] = "Atual" if arq == mais_recente else "Obsoleto"
    
    # Salvar histórico
    with open(caminho_json, "w", encoding="utf-8") as f:
        json.dump(historico, f, indent=4, ensure_ascii=False)
    
    return historico

##############################################
# Funções de Movimentação de Arquivos
##############################################
def criar_pastas_organizacao():
    """
    Cria as pastas para arquivos revisados e obsoletos.
    """
    base_dir = r"C:\Users\PROJETOS\Downloads\OAE-467 - PETER-KD-ENG"
    pasta_revisados = os.path.join(base_dir, "Revisados")
    if not os.path.exists(pasta_revisados):
        os.makedirs(pasta_revisados)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    pasta_obsoletos = os.path.join(base_dir, f"Obsoletos_{timestamp}")
    if not os.path.exists(pasta_obsoletos):
        os.makedirs(pasta_obsoletos)
    return pasta_revisados, pasta_obsoletos

def mover_arquivos(lista_arquivos, destino):
    import shutil
    print("[DEBUG] mover_arquivos() chamado.")
    print(f"[DEBUG] Pasta de destino: {destino}")
    print(f"[DEBUG] Quantidade de arquivos para mover: {len(lista_arquivos)}")

    for idx, arq in enumerate(lista_arquivos):
        print(f"[DEBUG] Analisando arquivo {idx+1}/{len(lista_arquivos)}: {arq}")
        
        origem = arq.get("caminho_completo")  # Recupera usando a chave correta
        nome_arquivo = arq.get("Nome do Arquivo")

        if not origem:
            print(f"[DEBUG] -> 'caminho_completo' não encontrado no dicionário. Pulando.")
            continue
        
        if not os.path.exists(origem):
            print(f"[DEBUG] -> O arquivo '{origem}' não existe. Pulando.")
            continue

        destino_arquivo = os.path.join(destino, nome_arquivo)
        print(f"[DEBUG] -> Movendo de '{origem}' para '{destino_arquivo}'...")
        try:
            shutil.move(origem, destino_arquivo)
            print("[DEBUG] -> Movimento bem-sucedido!")
        except Exception as e:
            print(f"[ERRO] -> Falha ao mover '{origem}' para '{destino_arquivo}': {e}")

def mover_obsoletos(lista_obsoletos, destino):
    import shutil
    print("[DEBUG] mover_obsoletos() chamado.")
    print(f"[DEBUG] Pasta de destino: {destino}")
    print(f"[DEBUG] Quantidade de arquivos obsoletos: {len(lista_obsoletos)}")

    for idx, arq in enumerate(lista_obsoletos):
        print(f"[DEBUG] Analisando arquivo obsoleto {idx+1}/{len(lista_obsoletos)}: {arq}")

        origem = arq.get("caminho_completo")  # Usar a mesma chave
        nome_arquivo = arq.get("Nome do Arquivo")

        if not origem:
            print(f"[DEBUG] -> 'caminho_completo' não encontrado no dicionário. Pulando.")
            continue
        
        if not os.path.exists(origem):
            print(f"[DEBUG] -> O arquivo '{origem}' não existe. Pulando.")
            continue

        base, ext = os.path.splitext(nome_arquivo)
        destino_arquivo = os.path.join(destino, base + "_OBSOLETO" + ext)
        print(f"[DEBUG] -> Movendo obsoleto de '{origem}' para '{destino_arquivo}'...")
        try:
            shutil.move(origem, destino_arquivo)
            print("[DEBUG] -> Movimento bem-sucedido!")
        except Exception as e:
            print(f"[ERRO] -> Falha ao mover '{origem}' para '{destino_arquivo}': {e}")

def pos_processamento(primeira_entrega, diretorio, dados_anteriores, arquivos_novos, arquivos_revisados, arquivos_alterados, obsoletos):
    print("[DEBUG] pos_processamento() chamado.")
    print(f"[DEBUG] primeira_entrega={primeira_entrega}, diretorio={diretorio}")
    print(f"[DEBUG] arquivos_novos={len(arquivos_novos)}, arquivos_revisados={len(arquivos_revisados)}, "
          f"arquivos_alterados={len(arquivos_alterados)}, obsoletos={len(obsoletos)}")

    pasta_revisados, pasta_obsoletos = criar_pastas_organizacao()
    if pasta_revisados is None or pasta_obsoletos is None:
        messagebox.showerror("Erro", "Não foi possível criar pastas para organizar os arquivos. Verifique as permissões.")
        return

    # Chamadas efetivas de movimentação:
    mover_arquivos(arquivos_revisados, pasta_revisados)
    mover_obsoletos(obsoletos, pasta_obsoletos)

    messagebox.showinfo("Concluído", "Processo concluído com sucesso.")
    import sys
    sys.exit(0)


# Interface inicial para seleção de projetos
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
    projetos = list(projetos_dict.items())
    sel = {"numero": None, "caminho": None}

    def filtrar(*args):
        termo = entrada.get().lower()
        tree.delete(*tree.get_children())
        for numero, caminho in projetos:
            if termo in caminho.lower():
                tree.insert("", tk.END, values=(numero, caminho))
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
        sel["caminho"] = vals[1]
        root.destroy()
        Disciplinas_Detalhes_Projeto(sel["numero"], sel["caminho"])  # Abrir nova interface

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
    tk.Label(frame, text="Digite nome ou parte do nome do projeto:", font=("Arial", 10)).pack(anchor="w")
    entrada = tk.Entry(frame)
    entrada.pack(fill=tk.X)
    entrada.bind("<KeyRelease>", filtrar)
    entrada.bind("<Return>", lambda e: confirmar())
    cols = ("Número", "Caminho/Nome")
    tree = ttk.Treeview(frame, columns=cols, show="headings", height=10)
    for c in cols:
        tree.heading(c, text=c)
    tree.pack(fill=tk.BOTH, expand=True)
    for numero, caminho in projetos:
        tree.insert("", tk.END, values=(numero, caminho))
    btn_frame = tk.Frame(frame)
    btn_frame.pack(pady=5)
    ttk.Button(btn_frame, text="Confirmar", command=confirmar, bootstyle="sucess").pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="Cancelar", command=root.destroy, bootstyle="danger").pack(side=tk.LEFT, padx=5)
    root.mainloop()
    return sel["numero"], sel["caminho"]

# Nova interface para exibir as disciplinas do projeto
def Disciplinas_Detalhes_Projeto(numero, caminho):
    """
    Exibe as disciplinas do projeto selecionado.
    Permite ao usuário selecionar uma disciplina e prosseguir para a pasta de entregas.
    """
    disciplinas_path = os.path.join(caminho, "3 Desenvolvimento")
    
    if not os.path.exists(disciplinas_path):
        messagebox.showerror("Erro", "A pasta de disciplinas não foi encontrada no projeto selecionado.")
        return

    # Criar janela
    nova_janela = tk.Tk()
    nova_janela.title(f"Gerenciador de Projetos - Projeto {numero}")
    nova_janela.geometry("900x600")

    # Cabeçalho
    header = tk.Label(
         nova_janela,
        text=f"Projeto {numero} - Selecione a disciplina para entrega",
        font=("Helvetica", 14, "bold"),
        anchor="w"
    )
    header.pack(fill=tk.X, padx=10, pady=5)

    # Configuração do Treeview para exibir as disciplinas
    cols = ["Nome", "Data de Modificação", "Tipo", "Tamanho"]
    tree = ttk.Treeview(nova_janela, columns=cols, show="headings", height=20)
    tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=200 if col == "Nome" else 150, anchor="w")

    # Carregar disciplinas dinamicamente
    disciplinas = []
    for item in os.listdir(disciplinas_path):
        full_path = os.path.join(disciplinas_path, item)
        if os.path.isdir(full_path):
            mod_time = datetime.fromtimestamp(os.path.getmtime(full_path)).strftime("%d/%m/%Y %H:%M")
            size = "--"  # Pastas não têm tamanho direto
            disciplinas.append((item, mod_time, "Pasta", size))
    
    # Inserir disciplinas no Treeview
    for disciplina in disciplinas:
        tree.insert("", tk.END, values=disciplina)

    # Função para abrir a pasta de entregas da disciplina selecionada
    def confirmar_selecao_arquivos():
        """
        Após selecionar os arquivos no explorador de arquivos,
        extrai os dados e envia para a interface `exibir_interface_tabela`.
        """
        # Verificar a disciplina selecionada no Treeview
        selecionados = tree.selection()
        if not selecionados:
            messagebox.showwarning("Atenção", "Nenhuma disciplina selecionada.")
            return

        valores = tree.item(selecionados[0])["values"]
        disciplina_nome = valores[0]  # Nome da disciplina
        pasta_entrega = os.path.join(disciplinas_path, disciplina_nome, "1.ENTREGAS")

        # Verificar se a pasta de entrega existe
        if not os.path.exists(pasta_entrega):
            messagebox.showerror("Erro", f"A pasta de entrega '{pasta_entrega}' não foi encontrada.")
            return

        # Abrir o explorador de arquivos com a pasta de entrega como diretório inicial
        arquivos_selecionados = filedialog.askopenfilenames(
            title="Selecione arquivos para entrega",
            initialdir=pasta_entrega
        )

        if not arquivos_selecionados:
            messagebox.showwarning("Atenção", "Nenhum arquivo foi selecionado.")
            return

        # Processar os arquivos selecionados
        arquivos_processados = []
        for arquivo in arquivos_selecionados:
            nome_arquivo = os.path.basename(arquivo)
            dados_extraidos = extrair_dados_arquivo(nome_arquivo)
            dados_extraidos["caminho_completo"] = arquivo  # Adiciona o caminho completo
            arquivos_processados.append(dados_extraidos)

        if not arquivos_processados:
            messagebox.showerror("Erro", "Nenhum dado foi processado dos arquivos selecionados.")
            return

        # Chamar a tela `exibir_interface_tabela` com os arquivos processados
        nova_janela.destroy()  # Fecha a janela atual
        exibir_interface_tabela(numero, arquivos_previos=arquivos_processados)

    content = tk.Frame(nova_janela, bg="#f5f5f5", padx=20, pady=20)
    content.pack(fill=tk.BOTH, expand=True)

    btn_frame = tk.Frame( nova_janela)
    btn_frame.pack(fill=tk.X, pady=5, padx=10)

    ttk.Button(btn_frame, text="Confirmar Seleção", command=confirmar_selecao_arquivos, bootstyle="success").pack(side=tk.RIGHT, padx=5)    
    ttk.Button(btn_frame, text="Cancelar", command= nova_janela.destroy, bootstyle="danger").pack(side=tk.LEFT, padx=5)

    nova_janela.mainloop()


'''
 - Função modificada para verificação de nomenclaturas
 - Função para carregar o JSON como dicionário
 '''
def carregar_json(filepath: str) -> Dict[str, Any]:
    if os.path.exists(filepath):
        with open(filepath, 'r', encoding='utf-8') as file:
            return json.load(file)
    else:
        return {}

def salvar_json(filepath: str, data: Dict[str, Any]):
    with open(filepath, 'w', encoding='utf-8') as file:
        json.dump(data, file, indent=4, ensure_ascii=False)

def identificar_numero_arquivo(partes):
    elementos_relevantes = partes[5:]

    for parte in elementos_relevantes:
        if (
            any(char.isdigit() for char in parte)  # Contém números
            and not parte.isalpha()               # Não é apenas letras (ex.: BLD)
            and len(parte) <= 6                   # Limita o tamanho (evita capturar coisas como "467")
        ):
            return parte
    return ""
    
# Função revisada para extrair dados do nome do arquivo
def extrair_dados_arquivo(nome_arquivo):
    nome_base, extensão = os.path.splitext(nome_arquivo)
    partes = nome_base.split('-')

    try:
        dados = {
            "Status": partes[0] if len(partes) > 0 else "",
            "Nome do Arquivo": nome_arquivo,
            "Extensão": extensão.strip('-'),

            "Nº do Arquivo": identificar_numero_arquivo(partes),  # Chama a nova função
            
            "Fase": partes[5] if len(partes) > 5 else "",  # Capturar "PE"
            "Tipo": partes[6] if len(partes) > 6 else "",  # Capturar "DTE"
            "Revisão": partes[-1].split('.')[0] if '.' in partes[-1] else partes[-1],  # Capturar "R01"
            "Modificação": datetime.now().strftime("%d/%m/%Y"),  # Data atual dinâmica
            "Modificado por": "Usuário",  # Placeholder
            "Entrega": f"Entrega.{partes[7].split('.')[0]}" if len(partes) > 7 else ""  # Ajustado para "Entrega.001"
        }
    except IndexError:
        # Caso algum elemento não seja encontrado, preencher com vazio
        dados = {
            "Status": "",
            "Nome do Arquivo": nome_arquivo,
            "Extensão": "",
            "Nº do Arquivo": "",
            "Fase": "",
            "Tipo": "",
            "Revisão": "",
            "Modificação": datetime.now().strftime("%d/%m/%Y"),
            "Modificado por": "Usuário",
            "Entrega": ""
        }
    return dados

#Exibe uma interface gráfica com a tabela baseada na lógica revisada.
def exibir_interface_tabela(numero, arquivos_previos=None):

    janela = tk.Tk()
    janela.title(f"Gerenciador de Projetos - Projeto {numero}")
    janela.geometry("1200x800")

    # Frame principal que separa a barra lateral do restante do conteúdo
    frame_principal = tk.Frame(janela)
    frame_principal.pack(fill=tk.BOTH, expand=True)

    # Adiciona a barra lateral
    barra_lateral = tk.Frame(frame_principal, bg="#2c3e50", width=200)
    barra_lateral.pack(side=tk.LEFT, fill=tk.Y)

    # Cabeçalho da barra lateral
    lbl_titulo = tk.Label(
        barra_lateral,
        text="OAE - Engenharia",
        font=("Helvetica", 14, "bold"),
        bg="#2c3e50",
        fg="white"
    )
    lbl_titulo.pack(pady=10)

    # Sessão de projetos
    lbl_projetos = tk.Label(
        barra_lateral,
        text="PROJETOS",
        font=("Helvetica", 10, "bold"),
        bg="#34495e",
        fg="white",
        anchor="w",
        padx=10
    )
    lbl_projetos.pack(fill=tk.X, pady=5)

    # Lista de projetos
    lst_projetos = tk.Listbox(barra_lateral, height=5, bg="#ecf0f1", font=("Helvetica", 9))
    lst_projetos.pack(fill=tk.X, padx=10, pady=5)
    lst_projetos.insert(tk.END, "OAE-467 - PETER-KD-ENG")

    # Sessão de membros
    lbl_membros = tk.Label(
        barra_lateral,
        text="MEMBROS",
        font=("Helvetica", 10, "bold"),
        bg="#34495e",
        fg="white",
        anchor="w",
        padx=10
    )
    lbl_membros.pack(fill=tk.X, pady=5)

    # Preenchendo espaço restante
    barra_lateral.pack_propagate(False)
    
    # Conteúdo principal (restante da interface)
    conteudo_principal = tk.Frame(frame_principal)
    conteudo_principal.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    def fazer_analise_nomenclatura():
        lista_arquivos = []
        for item in tabela.get_children():
            valores = tabela.item(item)["values"]
            lista_arquivos.append({
                "Status": valores[0],
                "Nome do Arquivo": valores[1],
                "Extensão": valores[2],
                "Nº do Arquivo": valores[3],
                "Fase": valores[4],
                "Tipo": valores[5],
                "Revisão": valores[6],
                "Modificação": valores[7],
                "Modificado por": valores[8],
                "Entrega": valores[9]
            })
        if not lista_arquivos:
            messagebox.showinfo("Aviso", "Nenhum arquivo adicionado para análise.")
        else:
            janela.destroy()
            tela_analise_nomenclatura(lista_arquivos)
    
    # Cabeçalho
    lbl_instrucao = tk.Label(
        conteudo_principal, 
        text="Adicionar Arquivos para Entrega", 
        font=("Helvetica", 15, "bold"),
        anchor="w"
    )
    lbl_instrucao.place(x=10, y=10)  # Define a posição no canto superior esquerdo

    # Frame para botões
    frm_botoes = tk.Frame(conteudo_principal)
    frm_botoes.pack(side=tk.TOP, anchor="ne", pady=10, padx=10)
    ttk.Button(frm_botoes, text="Fazer análise da Nomenclatura", command=fazer_analise_nomenclatura, bootstyle="info").pack(side=tk.LEFT, padx=5)

    cols = ["Status", "Nome do Arquivo", "Extensão", "Nº do Arquivo", "Fase", "Tipo", "Revisão", 
            "Modificação", "Modificado por", "Entrega"]

    tabela = ttk.Treeview(conteudo_principal, columns=cols, show="headings", height=20)
    for col in cols:
        tabela.heading(col, text=col)
        tabela.column(col, width=120 if col != "Nome do Arquivo" else 300)

    tabela.pack(fill=tk.BOTH, expand=True)

    # Adicionar arquivos prévios, se existirem
    if arquivos_previos:
        for dados_extraidos in arquivos_previos:
            tabela.insert("", tk.END, values=(
                dados_extraidos["Status"], dados_extraidos["Nome do Arquivo"], dados_extraidos["Extensão"], 
                dados_extraidos["Nº do Arquivo"], dados_extraidos["Fase"], dados_extraidos["Tipo"], 
                dados_extraidos["Revisão"], dados_extraidos["Modificação"], 
                dados_extraidos["Modificado por"], dados_extraidos["Entrega"]
            ))

    def adicionar_arquivos():
        arquivos = filedialog.askopenfilenames(title="Selecione arquivos")
        for arquivo in arquivos:
            nome_arquivo = os.path.basename(arquivo)
            dados_extraidos = extrair_dados_arquivo(nome_arquivo)
            dados_extraidos["caminho_completo"] = arquivo
            tabela.insert("", tk.END, values=(
                dados_extraidos["Status"], dados_extraidos["Nome do Arquivo"], dados_extraidos["Extensão"], 
                dados_extraidos["Nº do Arquivo"], dados_extraidos["Fase"], dados_extraidos["Tipo"], 
                dados_extraidos["Revisão"], dados_extraidos["Modificação"], 
                dados_extraidos["Modificado por"], dados_extraidos["Entrega"]
            ))

    def remover_arquivo():
        selecionados = tabela.selection()
        if selecionados:
            for item in selecionados:
                tabela.delete(item)
        else:
            messagebox.showinfo("Informação", "Nenhum item selecionado.")
    
    btn_frame = tk.Frame(conteudo_principal)
    btn_frame.pack(side=tk.LEFT, pady=10, padx=10)
    ttk.Button(btn_frame, text="Adicionar Arquivo", command=adicionar_arquivos, bootstyle="success").pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="Remover Arquivo", command=remover_arquivo, bootstyle="warning").pack(side=tk.LEFT, padx=5)

    btn_frame = tk.Frame(conteudo_principal)
    btn_frame.pack(side=tk.RIGHT, pady=10, padx=10)
    ttk.Button(btn_frame, text="Sair", command=janela.destroy, bootstyle="danger").pack(side=tk.RIGHT, padx=5)

def tela_analise_nomenclatura(lista_arquivos):
    padroes_nomenclatura = carregar_padroes_nomenclatura()

    if not padroes_nomenclatura:
        messagebox.showerror("Erro", "Nenhum padrão de nomenclatura carregado.")
        return

    def verificar_tokens(tokens):
        for padrao in padroes_nomenclatura:
            if len(tokens) != len(padrao):
                continue
            tags = [
                "ok" if tokens[i] == padrao[i] else "mismatch"
                for i in range(len(padrao))
            ]
            if "mismatch" not in tags:
                return tags, padrao
        return ["mismatch"] * len(tokens), None

    def mostrar_nomenclatura_padrao():
        win = tk.Toplevel(janela)
        win.title("Padrões de Nomenclatura")
        win.geometry("800x300")
        lbl_padrao = tk.Label(win, text="Padrões de Nomenclatura:")
        lbl_padrao.pack(pady=10)

        frm_padrao = tk.Frame(win)
        frm_padrao.pack(fill=tk.BOTH, expand=True)

        for padrao in padroes_nomenclatura:
            lbl_linha = tk.Label(frm_padrao, text=" - ".join(padrao), relief=tk.RIDGE, padx=5, pady=5)
            lbl_linha.pack(anchor="w", padx=5, pady=2)

    def avancar():
        for iid in tree.get_children():
            tags = tree.item(iid, "tags")
            if "mismatch" in tags:
                messagebox.showerror("Erro", "Corrija os erros antes de avançar.")
                return
        janela.destroy()
        tela_verificacao_revisao(lista_arquivos)
    
    def voltar():
        janela.destroy()
        exibir_interface_tabela("467", lista_arquivos)

    janela = tk.Tk()
    janela.title("Verificação de Nomenclatura")
    janela.geometry("1200x600")

    lbl_instrucao = tk.Label(
        janela,
        text="Confira a nomenclatura (campos e separadores). Caso haja erros, corrija antes de avançar.",
        font=("Helvetica", 12)
    )
    lbl_instrucao.pack(pady=10)

    frm_botoes = tk.Frame(janela)
    frm_botoes.pack(pady=10)
    ttk.Button(frm_botoes, text="Mostrar Padrões", command=mostrar_nomenclatura_padrao, bootstyle="secondary").pack(side=tk.LEFT, padx=5)
    ttk.Button(frm_botoes, text="Voltar", command=voltar, bootstyle="warning").pack(side=tk.LEFT, padx=5)
    ttk.Button(frm_botoes, text="Avançar", command=avancar, bootstyle="success").pack(side=tk.RIGHT, padx=5)

    tree = ttk.Treeview(janela, show="headings", height=20)
    tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    colunas = [f"Token {i+1}" for i in range(len(padroes_nomenclatura[0]))]
    tree["columns"] = colunas
    for col in colunas:
        tree.heading(col, text=col)
        tree.column(col, width=80, anchor="center")

    for arq in lista_arquivos:
        nome_base, _ = os.path.splitext(arq["Nome do Arquivo"])
        tokens = nome_base.split("-")
        tags, padrao_usado = verificar_tokens(tokens)
        item_id = tree.insert("", tk.END, values=tokens)
        if "mismatch" in tags:
            tree.item(item_id, tags=("error",))
        else:
            tree.item(item_id, tags=("ok",))

    tree.tag_configure("error", background="#FFCCCC")
    tree.tag_configure("ok", background="#CCFFCC")

'''
Identifica a revisão mais recente para cada conjunto de arquivos com o mesmo nome base
(excluindo a revisão). Marca os arquivos antigos como "Obsoletos" e os mais recentes como "Revisados".
'''
def identificar_revisoes(lista_arquivos):
    grupos = {}
    
    for arq in lista_arquivos:
        nome_base, _ = os.path.splitext(arq["Nome do Arquivo"])
        tokens = nome_base.split("-")
        if len(tokens) < 2:
            continue  # Ignorar arquivos com nome inválido
        
        identificador = "-".join(tokens[:-1])  # Ignora apenas o campo da revisão
        revisao = tokens[-1] if tokens[-1].startswith("R") and tokens[-1][1:].isdigit() else "R00"

        if identificador not in grupos:
            grupos[identificador] = []

        grupos[identificador].append((revisao, arq))

    arquivos_revisados = []
    arquivos_obsoletos = []

    for identificador, arquivos in grupos.items():
        arquivos.sort(key=lambda x: int(x[0][1:]))  # Ordena pela numeração da revisão (exemplo: "R01" -> 1)
        revisao_mais_recente = arquivos[-1][1]
        
        arquivos_revisados.append(revisao_mais_recente)
        arquivos_obsoletos.extend([arq[1] for arq in arquivos[:-1]])

    return arquivos_revisados, arquivos_obsoletos

def tela_verificacao_revisao(lista_arquivos):
    arquivos_revisados, arquivos_obsoletos = identificar_revisoes(lista_arquivos)

    janela = tk.Tk()
    janela.title("Verificação de Revisão")
    janela.geometry("1000x700")

    lbl_instrucao = tk.Label(janela, text="Confira os arquivos revisados e obsoletos antes da entrega.", font=("Helvetica", 12))
    lbl_instrucao.pack(pady=10)

    # Frame para os arquivos revisados
    frame_revisados = tk.LabelFrame(janela, text="Arquivos Revisados", font=("Helvetica", 11, "bold"))
    frame_revisados.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    tree_revisados = ttk.Treeview(frame_revisados, columns=["Nome do Arquivo", "Revisão"], show="headings", height=10)
    tree_revisados.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    tree_revisados.heading("Nome do Arquivo", text="Nome do Arquivo")
    tree_revisados.heading("Revisão", text="Revisão")

    for arq in arquivos_revisados:
        tree_revisados.insert("", tk.END, values=(arq["Nome do Arquivo"], arq["Revisão"]))

    # Frame para os arquivos obsoletos
    frame_obsoletos = tk.LabelFrame(janela, text="Arquivos Obsoletos", font=("Helvetica", 11, "bold"))
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
        messagebox.showinfo("Confirmação", "Arquivos revisados e obsoletos identificados com sucesso.")
        # Aqui você de fato chama as funções de movimentação
        # 1) Criar as pastas
        pasta_revisados, pasta_obsoletos = criar_pastas_organizacao()
        # 2) Mover revisados
        mover_arquivos(arquivos_revisados, pasta_revisados)
        # 3) Mover obsoletos
        mover_obsoletos(arquivos_obsoletos, pasta_obsoletos)
        janela.destroy()
        

    # Botões de ação
    btn_frame = tk.Frame(janela)
    btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text="Voltar", command=voltar, bootstyle="warning").pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="Confirmar", command=confirmar, bootstyle="success").pack(side=tk.RIGHT, padx=5)

    janela.mainloop()

# Executa a interface
if __name__ == "__main__":
    exibir_interface_tabela()
