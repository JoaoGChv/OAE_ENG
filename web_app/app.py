from flask import Flask, render_template, redirect, url_for
from utils.planilha_gerador import criar_ou_atualizar_planilha
from pathlib import Path
import tempfile

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/start_delivery')
def start_delivery():
    """Example route that generates a dummy GRD spreadsheet."""
    temp_dir = tempfile.gettempdir()
    excel_path = Path(temp_dir) / 'grd_example.xlsx'
    criar_ou_atualizar_planilha(excel_path, 'AP', 1, temp_dir, [])
    return render_template('result.html', path=excel_path)

if __name__ == '__main__':
    app.run(debug=True)