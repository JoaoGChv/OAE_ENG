---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Repositório da Aplicação TKinter para gerenciamento de arquivos da OAE - ENG.
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

*Está é uma interface Tkinter que faz o gerenciamento de arquivos para um escritório de projetos em Engenharia. O projeto tem especificações próprias mas é modular o suficiente para que possa ser aplicado em outros contextos. Ele permite:*

- Seleção de Projetos: Através de uma interface gráfica, o usuário pode selecionar um projeto a partir de uma lista carregada de um arquivo JSON.
- Seleção de Disciplinas: Dentro de um projeto, o usuário escolhe uma disciplina específica para a qual os arquivos serão entregues.
- Seleção de Arquivos: O usuário seleciona os arquivos relevantes para a entrega dentro da pasta de entregas da disciplina escolhida.
- Extração de Dados: O código extrai informações relevantes dos nomes dos arquivos, como status, fase, tipo e revisão.
- Validação de Nomenclatura: Uma interface permite ao usuário revisar e corrigir a nomenclatura dos arquivos com base em regras predefinidas.
- Verificação de Revisão: O sistema identifica e separa arquivos revisados e obsoletos com base na nomenclatura.
- Organização de Arquivos: Os arquivos são movidos para pastas "Revisados" e "Obsoletos" para melhor organização.
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Para consguir rodar a aplicação em qualquer máquina, faça o seguinte processo:

- Instale o Python (Qualquer versão que tenha suporte);
- Rode no seu terminal (Linux ou PowerShell) - python -m venv venv ( __No Windows: .\venv\Scripts\activate__ );
- Dentro do seu terminal execute - pip install -r requirements.txt (__Quando fizer o Git clone, ele vai ser puxada para sua máquina__)

