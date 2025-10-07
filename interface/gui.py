# gui.py — interface simplifiée pour Word→Excel Updater
import os
import sys
import traceback
import importlib
import importlib.util
from pathlib import Path
from datetime import datetime
import shutil

# ---------- PySimpleGUI import (robuste selon version) ----------
try:
    import PySimpleGUI as sg
except Exception as e:
    raise RuntimeError(
        "PySimpleGUI n'est pas installé ou version incompatible. "
        "Installez-le avec :\n"
        "python -m pip install --upgrade --extra-index-url https://PySimpleGUI.net/install PySimpleGUI"
    ) from e

# safe theme call
if hasattr(sg, "theme"):
    try:
        sg.theme("SystemDefault")
    except Exception:
        sg.theme("DefaultNoMoreNagging")

# ---------- Répertoires du projet ----------
THIS_FILE = Path(__file__).resolve()
PROJECT_ROOT = THIS_FILE.parent.parent
SRC_DIR = PROJECT_ROOT / "src"
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"

# ajouter src dans sys.path
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

# chemins par défaut
DEFAULT_WORD = str(INPUT_DIR / "contrat_parametrage.docx") if (INPUT_DIR / "contrat_parametrage.docx").exists() else ""
DEFAULT_XLSX = str(OUTPUT_DIR / "parametres_mis_a_jour.xlsx") if (OUTPUT_DIR / "parametres_mis_a_jour.xlsx").exists() else ""

# ---------- Helpers ----------
_loaded_modules = {}

def load_module_from_src(module_name: str):
    """Charge ou recharge un module depuis src/."""
    module_path = SRC_DIR / f"{module_name}.py"
    if not module_path.exists():
        return None
    try:
        spec = importlib.util.spec_from_file_location(module_name, str(module_path))
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        _loaded_modules[module_name] = module
        return module
    except Exception as e:
        print(f"Erreur chargement module {module_name}: {e}")
        return None

def import_word_extractor():
    mod = load_module_from_src("word_reader")
    if not mod:
        return None, None
    for fname in ("extract_all_changes", "extract_cell_changes", "extract_rows", "extract_rows_with_actions"):
        if hasattr(mod, fname):
            return getattr(mod, fname), fname
    return None, None

def import_excel_updater():
    mod = load_module_from_src("excel_writer")
    if not mod:
        return None
    return getattr(mod, "update_excel", None)

# ---------- UI layout simplifié ----------
layout = [
    [sg.Text("Fichier Word (input) :")],
    [sg.Input(DEFAULT_WORD, key="-WORD-", expand_x=True), sg.FileBrowse(file_types=(("Word","*.docx"),))],
    [sg.Text("Fichier Excel (output) :")],
    [sg.Input(DEFAULT_XLSX, key="-XLSX-", expand_x=True), sg.FileBrowse(file_types=(("Excel","*.xlsx"),))],
    [sg.Checkbox("Faire une copie de sécurité avant écriture", default=True, key="-BACKUP-")],
    [sg.HorizontalSeparator()],
    [sg.Button("Analyser (preview)", key="-ANALYZE-"), sg.Button("Appliquer sur Excel", key="-APPLY-"), sg.Button("Quitter")],
    [sg.HorizontalSeparator()],
    [sg.Text("Aperçu des changements détectés :")],
    [sg.Table(values=[], headings=["Table/Mod","Key","Column","Action","New value"],
              key="-TABLE-", auto_size_columns=True, num_rows=15, expand_x=True, expand_y=True)],
    [sg.Text("Logs:")],
    [sg.Multiline(size=(80,10), key="-LOG-", autoscroll=True, disabled=False)]
]

window = sg.Window("Word→Excel Updater — Interface", layout, resizable=True, finalize=True)

# ---------- fonctions utilitaires ----------
current_changes = []
current_extractor_name = None

def log(msg, append=True):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    window["-LOG-"].update(f"[{ts}] {msg}\n", append=append)

def safe_backup(filepath):
    try:
        p = Path(filepath)
        if not p.exists():
            return None
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        bak = p.with_name(p.stem + "_backup_" + timestamp + p.suffix)
        shutil.copy2(filepath, bak)
        return str(bak)
    except Exception as e:
        log(f"Backup failed: {e}")
        return None

def format_changes_for_table(chgs):
    rows = []
    for ch in chgs:
        tbl = ch.get("table", "") or ch.get("table_name","")
        key = ch.get("key", "") or ch.get("Log_ID", "") or ch.get("Component_ID","")
        col = ch.get("column", "")
        act = ch.get("action", "")
        val = ch.get("new_value", "")
        rows.append([tbl, key, col, act, str(val)])
    return rows

# ---------- event loop ----------
while True:
    event, values = window.read(timeout=100)
    if event in (sg.WIN_CLOSED, "Quitter"):
        break

    if event == "-WORD-":
        extractor, name = import_word_extractor()
        if extractor:
            current_extractor_name = name
            window["-MODE-"].update(name)
            log(f"Importé word_reader.{name}")
        else:
            window["-MODE-"].update("(aucun)")
            log("Aucun extracteur trouvé.")

    if event == "-ANALYZE-":
        wordp = values["-WORD-"]
        if not wordp or not os.path.exists(wordp):
            sg.popup_error("Veuillez sélectionner un fichier Word valide.")
            continue

        extractor, name = import_word_extractor()
        if not extractor:
            sg.popup_error("Aucune fonction d'extraction trouvée dans src/word_reader.py.")
            continue

        try:
            log(f"Analyse du fichier : {wordp}")
            changes = extractor(wordp)
            if not isinstance(changes, list):
                sg.popup_error("Format inattendu : la fonction doit renvoyer une liste de changements.")
                continue
            current_changes = changes
            window["-TABLE-"].update(values=format_changes_for_table(changes))
            log(f"{len(changes)} changement(s) détecté(s).")
        except Exception as e:
            log(f"Erreur : {e}")
            log(traceback.format_exc())
            sg.popup_error("Erreur lors de l'analyse. Voir logs.")

    if event == "-APPLY-":
        wordp = values["-WORD-"]
        xlsxp = values["-XLSX-"]
        if not os.path.exists(xlsxp):
            sg.popup_error("Veuillez sélectionner un fichier Excel existant.")
            continue
        if not current_changes:
            sg.popup_error("Aucun changement détecté. Analysez le fichier Word d'abord.")
            continue

        if values["-BACKUP-"]:
            bak = safe_backup(xlsxp)
            if bak:
                log(f"Copie de sécurité créée : {bak}")

        updater = import_excel_updater()
        if not updater:
            sg.popup_error("Aucune fonction update_excel trouvée dans src/excel_writer.py.")
            continue

        try:
            log("Application des changements...")
            updater(current_changes, xlsxp)
            log("Mise à jour terminée. Vérifie le fichier Excel.")
            sg.popup_ok("Terminé — vérifie le fichier Excel.")
        except Exception as e:
            log(f"Erreur pendant l'application : {e}")
            log(traceback.format_exc())
            sg.popup_error("Erreur lors de la mise à jour. Voir logs.")

window.close()
