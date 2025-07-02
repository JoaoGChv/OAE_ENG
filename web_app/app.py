from flask import (
    Flask,
    render_template,
    redirect,
    url_for,
    request,
    flash,
)
from utils.planilha_gerador import criar_ou_atualizar_planilha
from oae.file_ops import (
    PROJETOS_JSON,
    criar_pasta_entrega_ap_pe,
    listar_arquivos_no_diretorio,
)
from pathlib import Path
import os
import json
import tempfile
import re

app = Flask(__name__)
app.secret_key = "oae-secret-key"


@app.route('/')
def index():
    return render_template('index.html')

def _load_projects():
    if not os.path.exists(PROJETOS_JSON):
        return []
    try:
        with open(PROJETOS_JSON, "r", encoding="utf-8") as f:
            return list(json.load(f).items())
    except Exception:
        return []


@app.route('/select_project', methods=['GET', 'POST'])
def select_project():
    projects = _load_projects()
    if request.method == 'POST':
        project = request.form.get('project')
        tipo = request.form.get('tipo', 'AP')
        if not project:
            flash('Selecione um projeto')
        else:
            return redirect(url_for('upload_files', project=project, tipo=tipo))
    return render_template('select_project.html', projects=projects)


def _next_delivery_dir(base: str, tipo: str) -> str:
    prefixo = '1.AP - Entrega-' if tipo == 'AP' else '2.PE - Entrega-'
    subdir = 'AP' if tipo == 'AP' else 'PE'
    pasta_base = os.path.join(base, subdir)
    os.makedirs(pasta_base, exist_ok=True)
    entregas = [
        d for d in os.listdir(pasta_base)
        if d.startswith(prefixo) and not d.endswith('-OBSOLETO')
    ]
    if entregas:
        nums = [int(re.search(r"(\d+)$", d).group(1)) for d in entregas]
        n_prox = max(nums) + 1
    else:
        n_prox = 1
    return os.path.join(pasta_base, f"{prefixo}{n_prox}")


@app.route('/upload', methods=['GET', 'POST'])
def upload_files():
    project = request.args.get('project')
    tipo = request.args.get('tipo', 'AP')
    if request.method == 'POST' and project:
        files = request.files.getlist('files')
        temp_saved = []
        tmp_dir = tempfile.mkdtemp()
        for f in files:
            if not f.filename:
                continue
            dest = os.path.join(tmp_dir, f.filename)
            f.save(dest)
            size = os.path.getsize(dest)
            temp_saved.append(('', f.filename, size, dest, ''))
        if temp_saved:
            try:
                criar_pasta_entrega_ap_pe(project, tipo, temp_saved)
                nova_pasta = _next_delivery_dir(project, tipo)
                arquivos = listar_arquivos_no_diretorio(nova_pasta)
                excel_path = Path(tempfile.gettempdir()) / 'grd_web.xlsx'
                criar_ou_atualizar_planilha(
                    excel_path,
                    tipo,
                    1,
                    nova_pasta,
                    arquivos,
                )
                return render_template('result.html', path=excel_path)
            except Exception as exc:
                flash(str(exc))
    return render_template('upload_files.html', project=project, tipo=tipo)

@app.route('/start_delivery')
def start_delivery():
    """Example route that generates a dummy GRD spreadsheet."""
    temp_dir = tempfile.gettempdir()
    excel_path = Path(temp_dir) / 'grd_example.xlsx'
    criar_ou_atualizar_planilha(excel_path, 'AP', 1, temp_dir, [])
    return render_template('result.html', path=excel_path)

if __name__ == '__main__':
    app.run(debug=True)