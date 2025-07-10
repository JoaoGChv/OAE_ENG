"""Command line entry point for delivery manager."""
# Este é um docstring que descreve a finalidade do script.

from oae.ui import main # Importa a função 'main' do módulo 'ui' dentro do pacote 'oae'.
                        # A função 'main' é presumivelmente o ponto de entrada principal da interface gráfica do usuário.
import logging # Importa o módulo 'logging' para configurar e gerenciar logs.

if __name__ == "__main__":
    # Este bloco de código será executado apenas quando o script é executado diretamente (não importado como módulo).
    logging.basicConfig(
        level=logging.DEBUG, # Define o nível mínimo de log para DEBUG. Isso significa que todas as mensagens de log (DEBUG, INFO, WARNING, ERROR, CRITICAL) serão processadas.
        format="%(asctime)s %(levelname)s [%(name)s] %(message)s" # Define o formato das mensagens de log:
                                                                  # %(asctime)s: Tempo de registro da mensagem
                                                                  # %(levelname)s: Nível de log (DEBUG, INFO, etc.)
                                                                  # [%(name)s]: Nome do logger que emitiu a mensagem
                                                                  # %(message)s: A mensagem de log propriamente dita
    )
    main() # Chama a função principal da interface do usuário, iniciando a aplicação.