# Project Name

## Description
Brief description of the project.

## Project Structure
```
├── data/
│   ├── raw/           # Unprocessed data
│   └── processed/     # Data after cleaning
├── notebooks/         # Jupyter Notebooks
├── src/               # Project source code
├── output/            # Reports and visualizations
├── tests/             # Unit tests
├── docs/              # Documentation
├── config/            # Configuration files
├── .gitignore
├── requirements.txt
└── README.md
```

## Pre-coding Checklist

Before starting to code, go through the following checks:

### 1. Connect to WSL in VS Code
- Open VS Code
- Press `Ctrl+Shift+P` and type **Remote-WSL: Connect to WSL**
- Alternatively, click the green `><` button in the bottom-left corner and select **Connect to WSL**
- The bottom-left corner should show `WSL: Ubuntu` (or your distro name) once connected

### 2. Select the `eval_env` environment in VS Code
- Press `Ctrl+Shift+P` and type **Python: Select Interpreter**
- Choose the interpreter that shows `eval_env` in its path (e.g. `~/eval_env/bin/python`)
- The selected environment will appear in the bottom status bar of VS Code

### 3. Make sure you are on the latest version of `main`
```bash
git checkout main
git fetch origin
git pull origin main
```
- Confirm you are up to date — the output should say `Already up to date.` or show the pulled changes
- You can verify the current branch with `git branch` (the active one is marked with `*`)

---

## Setup
```bash
pip install -r requirements.txt
```

## Usage

### Run a Jupyter Notebook
Notebooks live in the `notebooks/` folder. To open one in VS Code:
1. Open the notebook file (e.g. [notebooks/hello_world.ipynb](notebooks/hello_world.ipynb))
2. Click **Select Kernel** in the top-right corner
3. Choose **Jupyter Kernel... → eval_env**
4. Run cells with `Shift+Enter`

To launch Jupyter in the browser instead:
```bash
conda activate eval_env
jupyter notebook
```

---

### Run from the CLI
From the root of the project, run:
```bash
python3 main.py
```
Expected output:
```
Hello, World!
```
