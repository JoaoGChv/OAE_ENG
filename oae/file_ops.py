"""File processing utilities for OAE."""
from __future__ import annotations
import datetime
import json
import os
import re
import shutil
from typing import Dict, List, Sequence, Tuple
from utils.planilha_gerador import criar_ou_atualizar_planilha, _key # Importa funções para geração de planilhas

try:
    from openpyxl import Workbook # Importa a classe Workbook para manipular arquivos Excel
except Exception as exc:
    raise ImportError("openpyxl is required") from exc


def _resolve_json_path(env_var: str, default_path: str) -> str:
    # Resolve o caminho de um arquivo JSON, usando variável de ambiente ou caminho padrão
    return os.getenv(env_var, default_path)

# Caminho para o arquivo JSON de diretórios de projetos
PROJETOS_JSON: str = _resolve_json_path(
    "OAE_PROJETOS_JSON",
    r"G:\Drives compartilhados\OAE-JSONS\diretorios_projetos.json",
)

# Caminho para o arquivo JSON de nomenclaturas
NOMENCLATURAS_JSON: str = _resolve_json_path(
    "OAE_NOMENCLATURAS_JSON",
    r"G:\Drives compartilhados\OAE-JSONS\nomenclaturas.json",
)

# Caminho para o arquivo JSON do último diretório utilizado
ARQ_ULTIMO_DIR: str = _resolve_json_path(
    "OAE_ULTIMO_DIR_JSON",
    r"G:\Drives compartilhados\OAE-JSONS\ultimo_diretorio_arqs.json",
)

# Mapeia grupos de extensões de arquivos para facilitar a categorização
GRUPOS_EXT: Dict[str, Sequence[str]] = {
    "DWG/DXF": [".dwg", ".dxf"],
    "DOC/DOCX": [".doc", ".docx"],
    "XLS/XLSX": [".xls", ".xlsx"],
    "ZIP/RAR": [".zip", ".rar", ".7z"],
    "RVT": [".rvt"],
    "IFC": [".ifc"],
    "NWC": [".nwc"],
    "NWD": [".nwd"],
}

# Tupla com os nomes dos meses em português
MESES = (
    "janeiro","fevereiro","março","abril","maio","junho",
    "julho","agosto","setembro","outubro","novembro","dezembro"
)

# Nomes dos arquivos de planilha consolidada para entregas AP e PE
GRD_MASTER_AP_NOME = "GRD_ENTREGAS_AP.xlsx" # Nome da planilha mestra para entregas AP
GRD_MASTER_PE_NOME = "GRD_ENTREGAS_PE.xlsx" # Nome da planilha mestra para entregas PE
GRD_MASTER_NOME = "GRD_ENTREGAS.xlsx" # Mantido para compatibilidade com versões antigas


def criar_pasta_entrega_ap_pe(
    pasta_entrega_disc: str, # Diretório base da disciplina
    tipo: str, # Tipo de entrega ("AP" ou "PE")
    arquivos: list[Tuple[str, str, int, str, str]], # Lista de arquivos a serem copiados
) -> None:
    # Cria uma nova pasta de entrega para AP ou PE, renomeando a anterior como obsoleta
    prefixo = "1.AP - Entrega-" if tipo == "AP" else "2.PE - Entrega-" # Prefixo da pasta
    subdir = "AP" if tipo == "AP" else "PE" # Subdiretório para o tipo de entrega
    pasta_base = os.path.join(pasta_entrega_disc, subdir) # Caminho completo da pasta base
    os.makedirs(pasta_base, exist_ok=True) # Garante que a pasta base existe

    # Lista e ordena as entregas ativas para encontrar a última
    entregas_ativas = sorted(
        [d for d in os.listdir(pasta_base)
         if d.startswith(prefixo) and not d.endswith("-OBSOLETO")],
        key=lambda n: int(re.search(r"(\d+)$", n).group(1)) # Extrai o número da entrega para ordenação
    )
    # Calcula o número da próxima entrega
    n_prox = (
        int(re.search(r"(\d+)$", entregas_ativas[-1]).group(1)) + 1
        if entregas_ativas else 1
    )

    if entregas_ativas:
        # Renomeia a pasta da entrega anterior para "OBSOLETO"
        ant_path = os.path.join(pasta_base, entregas_ativas[-1]) # Caminho da pasta anterior
        novo_ant = ant_path + "-OBSOLETO" # Novo nome com sufixo OBSOLETO
        seq = 1
        while os.path.exists(novo_ant):
            seq += 1
            novo_ant = f"{ant_path}-OBSOLETO{seq}" # Adiciona número sequencial se já existir OBSOLETO
        os.rename(ant_path, novo_ant) # Renomeia a pasta antiga

    nova_pasta = os.path.join(pasta_base, f"{prefixo}{n_prox}") # Caminho da nova pasta de entrega
    os.makedirs(nova_pasta, exist_ok=False) # Cria a nova pasta

    for (_, nome, _, caminho_full, _) in arquivos:
        # Copia os arquivos para a nova pasta de entrega
        try:
            shutil.copy2(caminho_full, os.path.join(nova_pasta, nome)) # Copia arquivo, mantendo metadados
        except FileNotFoundError:
            # Ignora arquivos não encontrados (por exemplo, se foram movidos ou excluídos)
            continue


def _safe_json_load(fp) -> dict:
    # Carrega dados de um arquivo JSON com tratamento de erro
    try:
        return json.load(fp) # Tenta carregar o JSON
    except json.JSONDecodeError:
        return {} # Retorna dicionário vazio se houver erro de decodificação

def folder_mais_recente(base: str, tipo: str) -> str | None:
    """Return the most recent delivery folder for AP or PE inside *base*.

    Parameters
    ----------
    base:
        Directory containing the ``AP`` or ``PE`` subfolder.
    tipo:
        Either ``"AP"`` or ``"PE"`` to indicate which folder to inspect.

    Returns
    -------
    str | None
        Full path to the most recent delivery folder or ``None`` if none exist.
    """
    # Retorna o caminho da pasta de entrega mais recente (AP ou PE)
    pasta_tipo = os.path.join(base, tipo) # Caminho da pasta específica de tipo (AP/PE)
    if not os.path.isdir(pasta_tipo):
        return None # Retorna None se a pasta não existir
    candidatas = [] # Lista para armazenar pastas candidatas
    for d in os.listdir(pasta_tipo):
        if d.endswith("-OBSOLETO"):
            continue # Ignora pastas obsoletas
        m = re.search(r"Entrega-(\d+)$", d) # Busca o número da entrega no nome da pasta
        if m:
            candidatas.append((int(m.group(1)), d)) # Adiciona tupla (número, nome da pasta)
    if not candidatas:
        return None # Retorna None se não houver pastas de entrega
    _, pasta_nome = max(candidatas, key=lambda t: t[0]) # Encontra a pasta com o maior número de entrega
    return os.path.join(pasta_tipo, pasta_nome) # Retorna o caminho completo da pasta mais recente

def carregar_nomenclatura_json(numero_projeto: str) -> Dict | None:
    # Carrega as regras de nomenclatura de um projeto específico do arquivo JSON
    if not os.path.exists(NOMENCLATURAS_JSON):
        return None # Retorna None se o arquivo de nomenclaturas não existir
    with open(NOMENCLATURAS_JSON, "r", encoding="utf-8") as f:
        data = _safe_json_load(f) # Carrega os dados JSON com segurança
    return data.get(numero_projeto) # Retorna a nomenclatura para o número do projeto

def salvar_ultimo_diretorio(ultimo_dir: str) -> None:
    # Salva o último diretório utilizado em um arquivo JSON
    try:
        with open(ARQ_ULTIMO_DIR, "w", encoding="utf-8") as f:
            json.dump({"ultimo_diretorio": ultimo_dir}, f, ensure_ascii=False, indent=4) # Salva o diretório
    except OSError:
        pass # Ignora erros de sistema de arquivo

def carregar_ultimo_diretorio() -> str | None:
    # Carrega o último diretório utilizado de um arquivo JSON
    if os.path.exists(ARQ_ULTIMO_DIR):
        try:
            with open(ARQ_ULTIMO_DIR, "r", encoding="utf-8") as f:
                data = _safe_json_load(f) # Carrega os dados JSON com segurança
                return data.get("ultimo_diretorio") # Retorna o último diretório
        except OSError:
            return None # Retorna None em caso de erro de sistema de arquivo
    return None # Retorna None se o arquivo não existir

def extrair_numero_arquivo(nome_base: str) -> str:
    # Extrai um número de 3 dígitos do nome base de um arquivo
    if len(nome_base) <= 11:
        return "" # Retorna vazio se o nome for muito curto
    substring: str = nome_base[11:] # Pega a substring a partir do 12º caractere
    match = re.search(r"(\d{3})", substring) # Busca por 3 dígitos
    return match.group(1) if match else "" # Retorna o número encontrado ou vazio

REV_REGEX: re.Pattern[str] = re.compile(r"^(.*?)[-_]R(\d{1,3})$", re.IGNORECASE) # Expressão regular para identificar revisões

def identificar_nome_com_revisao(nome_arquivo: str) -> Tuple[str, str, str]:
    # Identifica o nome base, revisão e extensão de um arquivo
    nome_sem_extensao, extensao = os.path.splitext(nome_arquivo) # Separa nome e extensão
    nome_normalizado = nome_sem_extensao.replace("_", "-") # Normaliza o nome (troca _ por -)
    match = REV_REGEX.match(nome_normalizado) # Tenta casar com o padrão de revisão
    if match:
        nome_base = match.group(1) # Parte do nome antes da revisão
        revisao = "R" + match.group(2).zfill(2) # Formata a revisão (ex: R01)
        return nome_base, revisao, extensao.lower() # Retorna nome base, revisão e extensão
    return nome_sem_extensao, "", extensao.lower() # Retorna sem revisão se não encontrar

def _parse_rev(rev: str) -> int:
    # Converte uma string de revisão (ex: "R01") para um número inteiro
    if not rev:
        return -1 # Retorna -1 se a revisão for vazia
    digits = re.findall(r"\d+", rev) # Encontra todos os dígitos na revisão
    return int(digits[0]) if digits else -1 # Converte o primeiro grupo de dígitos em int

def comparar_revisoes(r1: str, r2: str) -> int:
    # Compara duas revisões numericamente
    try:
        return _parse_rev(r1) - _parse_rev(r2) # Retorna a diferença entre as revisões
    except ValueError:
        return 0 # Retorna 0 em caso de erro de valor

DEFAULT_SEPARATORS: set[str] = {"-", "."} # Separadores padrão para nomes de arquivos

def _obter_separadores_do_json(nomenclatura: Dict | None) -> set[str]:
    # Obtém os separadores definidos no JSON de nomenclatura
    seps: set[str] = set() # Conjunto para armazenar separadores
    if nomenclatura:
        for campo in nomenclatura.get("campos", []):
            sep = campo.get("separador") # Obtém o separador do campo
            if sep and isinstance(sep, str):
                seps.add(sep) # Adiciona separador ao conjunto
    return seps or DEFAULT_SEPARATORS # Retorna separadores encontrados ou os padrão

def split_including_separators(nome_sem_ext: str, nomenclatura: Dict | None) -> List[str]:
    # Divide um nome de arquivo em tokens, incluindo os separadores como tokens separados
    tokens: List[str] = [] # Lista para armazenar os tokens
    seps = _obter_separadores_do_json(nomenclatura) # Obtém os separadores
    i = 0
    while i < len(nome_sem_ext):
        c = nome_sem_ext[i] # Caractere atual
        if c in seps:
            tokens.append(c) # Adiciona o separador como um token
            i += 1
            continue
        j = i
        while j < len(nome_sem_ext) and nome_sem_ext[j] not in seps:
            j += 1 # Avança enquanto não encontra um separador
        tokens.append(nome_sem_ext[i:j]) # Adiciona a parte do nome como um token
        i = j
    return tokens # Retorna a lista de tokens

def verificar_tokens(tokens: Sequence[str], nomenclatura: Dict | None) -> List[str]:
    # Verifica se os tokens de um nome de arquivo correspondem à nomenclatura definida
    if not nomenclatura:
        return ["mismatch"] * len(tokens) # Retorna "mismatch" para todos se não houver nomenclatura

    campos_cfg = nomenclatura.get("campos", []) # Configuração dos campos da nomenclatura
    tokens_esperados: List[Tuple[str, object]] = [] # Lista de tokens esperados (campo ou separador)
    for idx, cinfo in enumerate(campos_cfg):
        tokens_esperados.append(("campo", cinfo)) # Adiciona o campo esperado
        if idx < len(campos_cfg) - 1:
            sep = cinfo.get("separador", "-") # Separador esperado entre campos
            tokens_esperados.append(("sep", sep)) # Adiciona o separador esperado

    result_tags: List[str] = [] # Lista de tags de resultado ("ok", "mismatch", "missing")
    idx_exp = idx_tok = 0 # Índices para tokens esperados e tokens reais
    while idx_tok < len(tokens) and idx_exp < len(tokens_esperados):
        token = tokens[idx_tok] # Token atual do arquivo
        tipo_esp, conteudo_esp = tokens_esperados[idx_exp] # Tipo e conteúdo esperado

        if tipo_esp == "sep":
            result_tags.append("ok" if token == conteudo_esp else "mismatch") # Verifica se o separador é o esperado
            idx_tok += 1
            idx_exp += 1
            continue

        tipo_campo = conteudo_esp.get("tipo", "Fixo") # Tipo do campo (ex: "Fixo")
        fixos = conteudo_esp.get("valores_fixos", []) # Valores fixos permitidos para o campo
        if tipo_campo == "Fixo" and fixos:
            valores_permitidos = [f.get("value") if isinstance(f, dict) else str(f) for f in fixos] # Lista de valores permitidos
            result_tags.append("ok" if token in valores_permitidos else "mismatch") # Verifica se o token está nos valores fixos
        else:
            result_tags.append("ok") # Se não for fixo ou não tiver valores fixos, considera "ok"
        idx_tok += 1
        idx_exp += 1

    while idx_tok < len(tokens):
        result_tags.append("mismatch") # Marca tokens extras como "mismatch"
        idx_tok += 1
    while idx_exp < len(tokens_esperados):
        result_tags.append("missing") # Marca tokens esperados que faltam como "missing"
        idx_exp += 1
    return result_tags # Retorna as tags de verificação

def identificar_obsoletos_custom(lista_arqs: Sequence[Tuple[str, str, int, str, str]]):
    # Identifica arquivos obsoletos com base em suas revisões (versões mais antigas)
    grouping: Dict[Tuple[str, str], List[Tuple[str, str, int, str, str]]] = {} # Agrupa arquivos por nome base e extensão
    for rv, a, tam, cam, dmod in lista_arqs:
        base, revision, ext = identificar_nome_com_revisao(a) # Extrai nome base, revisão e extensão
        key = (base.lower(), ext.lower()) # Chave para agrupamento (nome base + extensão)
        grouping.setdefault(key, []).append((rv, a, tam, cam, dmod)) # Adiciona o arquivo ao grupo

    obsoletos: List[Tuple[str, str, int, str, str]] = [] # Lista para armazenar arquivos obsoletos
    for arr in grouping.values():
        arr.sort(key=lambda x: _parse_rev(x[0]), reverse=True) # Ordena arquivos por revisão (mais recente primeiro)
        obsoletos.extend(arr[1:]) # Adiciona todos exceto o mais recente à lista de obsoletos
    return obsoletos # Retorna a lista de arquivos obsoletos

def carregar_dados_anteriores(diretorio: str) -> Dict:
    # Carrega dados de execução anterior de um arquivo JSON
    caminho = os.path.join(diretorio, "dados_execucao_anterior.json") # Caminho do arquivo de dados
    if os.path.exists(caminho):
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                return _safe_json_load(f) # Carrega os dados JSON com segurança
        except OSError:
            pass # Ignora erros de sistema de arquivo
    return {} # Retorna dicionário vazio se o arquivo não existir ou houver erro

def salvar_dados(diretorio: str, dados: Dict) -> None:
    # Salva dados (estado atual dos arquivos) em um arquivo JSON
    caminho = os.path.join(diretorio, "dados_execucao_anterior.json") # Caminho do arquivo de dados
    try:
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=4) # Salva os dados JSON formatados
    except OSError:
        pass # Ignora erros de sistema de arquivo

def obter_info_ultima_entrega(dados_anteriores: Dict) -> str:
    # Obtém informações formatadas sobre a última entrega
    entregas_oficiais = dados_anteriores.get("entregas_oficiais", 0) # Número de entregas oficiais
    ultima_execucao = dados_anteriores.get("ultima_execucao") # Data da última execução
    if ultima_execucao:
        dt = datetime.datetime.strptime(ultima_execucao, "%Y-%m-%d %H:%M:%S") # Converte string para datetime
        return f"Entrega {entregas_oficiais} de dia {dt.day} de {MESES[dt.month-1]} de {dt.year}" # Retorna info formatada
    return f"Entrega {entregas_oficiais}" # Retorna info básica se não houver data

def tentar_novamente_operacao(operacao, *args, **kwargs):
    # Tenta executar uma operação, relançando PermissionError
    while True:
        try:
            return operacao(*args, **kwargs) # Executa a operação
        except PermissionError:
            raise # Relança PermissionError

def gerar_nomes_entrega(num_entrega: int):
    # Gera nomes padronizados para arquivos de entrega e planilhas
    data_atual = datetime.datetime.now().strftime("%d_%m_%Y") # Data atual formatada
    return (
        f"GRD-ENTREGA.{num_entrega:02d}-{data_atual}.xlsx", # Nome do arquivo Excel da entrega
        f"ENTREGA.{num_entrega:02d}-{data_atual}", # Nome da planilha interna
        data_atual, # Apenas a data atual
    )

_SEP_PATTERN = r"[-_.]" # Padrão para separadores em nomes de arquivos

def renomear_para_arquivado(nome_arquivo: str) -> str:
    # Renomeia um arquivo para indicar que ele foi arquivado (ex: E-arq.dwg -> A-arq.dwg)
    base, ext = os.path.splitext(nome_arquivo) # Separa nome base e extensão
    m = re.match(rf"^([ECPR])({_SEP_PATTERN}.+)$", base, re.IGNORECASE) # Busca por E, C, P ou R no início
    if not m:
        return nome_arquivo # Retorna o nome original se não casar
    novo_base = "A" + m.group(2) # Substitui a primeira letra por 'A'
    return novo_base + ext # Retorna o novo nome com extensão

def mover_obsoletos_e_grd_anterior(obsoletos, diretorio: str, num_entrega_atual: int):
    # Move arquivos obsoletos e a planilha GRD da entrega anterior
    n_anterior = num_entrega_atual - 1 # Número da entrega anterior
    data_atual = datetime.datetime.now().strftime("%d_%m_%Y") # Data atual formatada
    pasta_pai = os.path.dirname(diretorio) # Diretório pai da entrega atual
    pasta_obsoletos = os.path.join(pasta_pai, f"Entrega_{n_anterior:02d}-Obsoletos-{data_atual}") # Caminho da pasta de obsoletos
    os.makedirs(pasta_obsoletos, exist_ok=True) # Garante que a pasta de obsoletos existe

    nome_arquivo_anterior, _, _ = gerar_nomes_entrega(n_anterior) # Nome do arquivo GRD da entrega anterior
    grd_anterior = os.path.join(diretorio, nome_arquivo_anterior) # Caminho completo da GRD anterior
    if os.path.exists(grd_anterior):
        # Move a planilha GRD da entrega anterior para a pasta de obsoletos
        tentar_novamente_operacao(shutil.move, grd_anterior, os.path.join(pasta_obsoletos, nome_arquivo_anterior))

    with open(os.path.join(pasta_obsoletos, "lista_obsoletos.txt"), "w", encoding="utf-8") as f:
        # Cria um arquivo de texto listando os arquivos obsoletos
        for rv, arq, *_ in obsoletos:
            f.write(arq + "\n")

    for rv, arq, _, cam, _ in obsoletos:
        # Move cada arquivo obsoleto para a pasta de obsoletos
        try:
            novo_nome = renomear_para_arquivado(arq) # Renomeia o arquivo para "arquivado"
            destino = os.path.join(pasta_obsoletos, novo_nome) # Caminho de destino do arquivo obsoleto

            if os.path.exists(destino):
                # Trata duplicatas adicionando sufixo (_dupX)
                base, ext = os.path.splitext(novo_nome)
                seq = 1
                while True:
                    cand = f"{base}_dup{seq}{ext}"
                    destino = os.path.join(pasta_obsoletos, cand)
                    if not os.path.exists(destino):
                        break
                    seq += 1

            tentar_novamente_operacao(shutil.move, cam, destino) # Move o arquivo
        except FileNotFoundError:
            continue # Ignora arquivos não encontrados

def criar_arquivo_excel(diretorio: str, num_entrega: int, arquivos):
    # Cria um novo arquivo Excel com uma lista básica de arquivos da entrega
    nome_arquivo, nome_planilha, _ = gerar_nomes_entrega(num_entrega) # Nomes para o arquivo e planilha
    caminho_excel = os.path.join(diretorio, nome_arquivo) # Caminho completo do novo arquivo Excel
    wb = Workbook() # Cria uma nova pasta de trabalho Excel
    ws = wb.active # Obtém a planilha ativa
    ws.title = nome_planilha # Define o título da planilha
    ws.append(["Nome do arquivo", "Revisão", "Caminho completo"]) # Adiciona cabeçalho
    for rv, arq, _, cam, _ in arquivos:
        ws.append([arq, rv or "", cam]) # Adiciona os dados dos arquivos
    wb.save(caminho_excel) # Salva o arquivo Excel
    return caminho_excel # Retorna o caminho do arquivo criado

def listar_arquivos_no_diretorio(diretorio):
    # Lista arquivos em um diretório, filtrando específicos
    ignorar = {
        "dados_execucao_anterior.json",
        GRD_MASTER_AP_NOME, # Ignora a planilha mestra AP
        GRD_MASTER_PE_NOME, # Ignora a planilha mestra PE
    }
    for f in os.listdir(diretorio):
        if f.startswith("GRD-ENTREGA."):
            ignorar.add(f) # Ignora planilhas de entrega anteriores
    saida = [] # Lista para armazenar os arquivos processados
    for raiz, _dirs, files in os.walk(diretorio):
        for a in files:
            if a in ignorar:
                continue # Pula arquivos a serem ignorados
            nb, rv, ex = identificar_nome_com_revisao(a) # Extrai nome base, revisão e extensão
            if ex in ['.jpg','.jpeg','.dwl','.dwl2','.png','.ini']:
                continue # Pula extensões de arquivo a serem ignoradas
            cam = os.path.join(raiz, a) # Caminho completo do arquivo
            tam = os.path.getsize(cam) # Tamanho do arquivo
            dmod_ts = os.path.getmtime(cam) # Timestamp da última modificação
            dmod = datetime.datetime.fromtimestamp(dmod_ts).strftime("%d/%m/%Y %H:%M") # Data de modificação formatada
            saida.append((rv, a, tam, cam, dmod)) # Adiciona tupla de informações do arquivo
    return saida # Retorna a lista de arquivos e suas informações

def analisar_comparando_estado(lista_de_arquivos, dados_anteriores):
    # Analisa a lista de arquivos comparando com o estado anterior para identificar mudanças
    grouping = {} # Agrupa arquivos por nome base e extensão
    for rv, a, tam, cam, dmod in lista_de_arquivos:
        nb, ex = _key(a) # Obtém chave de agrupamento (nome base, extensão)
        grouping.setdefault((nb, ex), []).append((rv, a, tam, cam, dmod)) # Adiciona ao grupo
    novos = [] # Lista de arquivos novos
    revisados = [] # Lista de arquivos revisados
    alterados = [] # Lista de arquivos alterados (mesma revisão, conteúdo diferente)
    for key, items in grouping.items():
        items.sort(key=lambda x: comparar_revisoes(x[0], 'R99')) # Ordena arquivos por revisão (crescente)
        ant = dados_anteriores.get(f"{key[0]}|{key[1]}", None) # Dados do arquivo na execução anterior
        rev_ant = ant["revisao"] if ant else "" # Revisão anterior
        tam_ant = ant["tamanho"] if ant else None # Tamanho anterior
        ts_ant = ant.get("timestamp") if ant else None # Timestamp anterior
        if not ant:
            # Se o arquivo não existia na execução anterior
            if items:
                novos.append(items[0]) # Primeiro item é novo
            for it in items[1:]:
                revisados.append(it) # Itens subsequentes são revisões
        else:
            # Se o arquivo já existia
            maior_rev = items[-1][0] # Maior revisão atual
            comp = comparar_revisoes(maior_rev, rev_ant) # Comparação da revisão atual com a anterior
            num_rev_ant = comparar_revisoes(rev_ant, '') # Número da revisão anterior
            if comp > 0:
                # Se a revisão atual é maior
                for (rvx, arqx, tamx, camx, dmodx) in items:
                    nr = comparar_revisoes(rvx, '') # Número da revisão do item atual
                    if nr > num_rev_ant:
                        revisados.append((rvx, arqx, tamx, camx, dmodx)) # Marca como revisado
                    elif nr == num_rev_ant:
                        ts_now = os.path.getmtime(camx) # Timestamp de modificação atual
                        if tamx != tam_ant or (ts_ant is not None and ts_now != ts_ant):
                            alterados.append((rvx, arqx, tamx, camx, dmodx)) # Marca como alterado se tamanho/timestamp mudou
            elif comp == 0:
                # Se a revisão é a mesma
                for (rvx, arqx, tamx, camx, dmodx) in items:
                    if rvx == rev_ant:
                        ts_now = os.path.getmtime(camx) # Timestamp de modificação atual
                        if tamx != tam_ant or (ts_ant is not None and ts_now != ts_ant):
                            alterados.append((rvx, arqx, tamx, camx, dmodx)) # Marca como alterado se tamanho/timestamp mudou
            else:
                # Se a revisão atual é menor (deve ser um caso de erro ou arquivo antigo)
                for (rvx, arqx, tamx, camx, dmodx) in items:
                    if rvx == rev_ant:
                        ts_now = os.path.getmtime(camx) # Timestamp de modificação atual
                        if tamx != tam_ant or (ts_ant is not None and ts_now != ts_ant):
                            alterados.append((rvx, arqx, tamx, camx, dmodx)) # Marca como alterado se tamanho/timestamp mudou
    return (novos, revisados, alterados) # Retorna as listas de arquivos por categoria

def pos_processamento(
    primeira_entrega, # Booleano indicando se é a primeira entrega
    diretorio, # Diretório da entrega
    dados_anteriores, # Dados da execução anterior
    arquivos_novos, # Lista de arquivos novos
    arquivos_revisados, # Lista de arquivos revisados
    arquivos_alterados, # Lista de arquivos alterados
    obsoletos, # Lista de arquivos obsoletos
    tipo_entrega: str | None = None, # Tipo de entrega ("AP" ou "PE")
):
    """Atualiza registros e planilhas após a análise de entrega."""
    # Realiza o pós-processamento da entrega: move obsoletos, atualiza planilhas e salva o estado
    num_entrega_atual = dados_anteriores.get("entregas_oficiais", 0) + 1 # Calcula o número da entrega atual

    master = GRD_MASTER_AP_NOME # Nome da planilha mestra padrão (AP)
    if (tipo_entrega or "AP") == "PE":
        master = GRD_MASTER_PE_NOME # Usa a planilha mestra PE se o tipo for PE
    caminho_excel_master = os.path.join(diretorio, master) # Caminho completo da planilha mestra

    if not primeira_entrega:
        if obsoletos or dados_anteriores.get("entregas_oficiais", 0) >= 1:
            # Move obsoletos e a GRD anterior se não for a primeira entrega
            mover_obsoletos_e_grd_anterior(obsoletos, diretorio, num_entrega_atual)
    if primeira_entrega:
        union_ = [] # Lista unificada de arquivos para a primeira entrega
        union_.extend(arquivos_novos)
        union_.extend(arquivos_revisados)
        union_.extend(arquivos_alterados)
        if not union_:
            return # Retorna se não houver arquivos na primeira entrega
    if primeira_entrega:
        lista_para_planilha = (
            arquivos_novos + arquivos_revisados + arquivos_alterados # Todos os arquivos relevantes para a primeira entrega
        )
        if not lista_para_planilha:
            return # Retorna se a lista para planilha estiver vazia
    else:
        lista_para_planilha = listar_arquivos_no_diretorio(diretorio) # Lista todos os arquivos no diretório para entregas subsequentes
        if not lista_para_planilha:
            return # Retorna se a lista para planilha estiver vazia

    criar_ou_atualizar_planilha(
        caminho_excel=caminho_excel_master, # Caminho da planilha mestra a ser atualizada
        tipo_entrega=tipo_entrega or "AP", # Tipo de entrega
        num_entrega=num_entrega_atual, # Número da entrega atual
        diretorio_base=diretorio, # Diretório base da entrega
        arquivos=lista_para_planilha, # Lista de arquivos para a planilha
        estado_anterior=dados_anteriores, # Dados da execução anterior
    )
    dados_anteriores["ultima_execucao"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") # Atualiza a data da última execução
    dados_anteriores["entregas_oficiais"] = num_entrega_atual # Atualiza o número de entregas oficiais
    grouping_final = {} # Agrupa todos os arquivos finais por nome base e extensão
    all_files_now = listar_arquivos_no_diretorio(diretorio) # Lista todos os arquivos no diretório novamente
    for rv, a, tam, cam, dmod in all_files_now:
        nb, rev, ex = identificar_nome_com_revisao(a) # Extrai informações do arquivo
        key = (nb.lower(), ex.lower()) # Chave de agrupamento
        grouping_final.setdefault(key, []).append((rv, a, tam, cam, dmod)) # Adiciona ao grupo
    for key, arr in grouping_final.items():
        arr.sort(key=lambda x: comparar_revisoes(x[0], 'R99')) # Ordena por revisão (crescente)
        revf = arr[-1][0] # Última revisão
        tamf = arr[-1][2] # Último tamanho
        camf = arr[-1][3] # Último caminho completo
        tsf = os.path.getmtime(camf) # Último timestamp de modificação
        dados_anteriores[f"{key[0]}|{key[1]}"] = {
            "revisao": revf if revf else '', # Armazena a revisão final
            "tamanho": tamf, # Armazena o tamanho final
            "timestamp": tsf, # Armazena o timestamp final
        }
    salvar_dados(diretorio, dados_anteriores) # Salva os dados atualizados da execução anterior