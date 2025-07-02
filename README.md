# OAE Engineering Delivery Manager

This repository contains tools for managing deliveries and generating Excel spreadsheets.

## Installation

Install the required Python packages using pip:

```bash
pip install -r requirements.txt
```

## Usage

Run the main script to start the application:

```bash
python gerenciar_entregas.py
```

`gerenciar_entregas.py` is a Tkinter-based tool that helps organize and verify
project deliverables. It reads project and nomenclature information from JSON
files, assists in selecting files for each delivery and generates or updates
a GRD (`GRD_ENTREGAS.xlsx`).

## Requirements

- Python 3 with Tkinter available
- `openpyxl` (required)
- `send2trash` (optional, for safe deletion)

Install packages with:

```bash
pip install openpyxl send2trash
```

## Environment Variables

Default locations for JSON files can be overridden using the following
variables:

- `OAE_PROJETOS_JSON` – path to the JSON containing project directories.
- `OAE_NOMENCLATURAS_JSON` – path to the JSON with naming conventions.
- `OAE_ULTIMO_DIR_JSON` – path where the last used directory is stored.

The script defines these defaults internally:

```python
PROJETOS_JSON = _resolve_json_path(
    "OAE_PROJETOS_JSON",
    r"G:\Drives compartilhados\OAE-JSONS\diretorios_projetos.json",
)
NOMENCLATURAS_JSON = _resolve_json_path(
    "OAE_NOMENCLATURAS_JSON",
    r"G:\Drives compartilhados\OAE-JSONS\nomenclaturas.json",
)
ARQ_ULTIMO_DIR = _resolve_json_path(
    "OAE_ULTIMO_DIR_JSON",
    r"G:\Drives compartilhados\OAE-JSONS\ultimo_diretorio_arqs.json",
)
```

## Usage

Run the application with Python:

```bash
python gerenciar_entregas.py
```

A graphical interface will guide you through selecting the project, choosing
files and updating the GRD spreadsheet.


