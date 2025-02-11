import sys
import os

# Adiciona o diretório raiz ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from utils.file_operations import listar_arquivos_no_diretorio, carregar_ultimo_diretorio
from utils.json_operations import carregar_nomenclatura_json
from ui.telas import janela_selecao_projeto

def main():
    print("Abrindo a interface de seleção de projeto...")
    numero_projeto, caminho_projeto = janela_selecao_projeto()

    if numero_projeto and caminho_projeto:
        print(f"Projeto selecionado: Número {numero_projeto}, Caminho {caminho_projeto}")
        # Aqui você pode continuar o fluxo de processamento
    else:
        print("Nenhum projeto foi selecionado.")

if __name__ == "__main__":
    main()

