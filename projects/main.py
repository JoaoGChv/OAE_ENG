import sys
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ui.telas import janela_selecao_projeto

def main():
    print("Abrindo a interface de seleção de projeto...")
    numero_projeto, caminho_projeto = janela_selecao_projeto()

    if numero_projeto and caminho_projeto:
        print(f"Projeto selecionado: Número {numero_projeto}, Caminho {caminho_projeto}")        
    else:
        print("Nenhum projeto foi selecionado.")

if __name__ == "__main__":
    main()

