import os
import shutil
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

def get_next_version(filepath):
    """Détermine le nom de la prochaine version (V1, V2, ...)."""
    if not os.path.exists(filepath):
        return "V1"
    wb = load_workbook(filepath)
    versions = [int(s[1:]) for s in wb.sheetnames if s.startswith("V") and s[1:].isdigit()]
    wb.close()
    return f"V{max(versions)+1}" if versions else "V1"

def copy_sheet_format(source_sheet, target_sheet):
    """
    Copie la mise en forme (largeurs de colonnes, styles de cellules) de source_sheet
    vers target_sheet *sans* recopier les valeurs.
    """
    from openpyxl.utils import get_column_letter

    # Copier les largeurs de colonnes (avec vérification)
    max_col = source_sheet.max_column
    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        col_dim = source_sheet.column_dimensions.get(col_letter)
        if col_dim is not None and getattr(col_dim, "width", None):
            target_sheet.column_dimensions[col_letter].width = col_dim.width

    # Copier les styles cellule par cellule
    for row in source_sheet.iter_rows():
        for cell in row:
            tgt = target_sheet.cell(row=cell.row, column=cell.column)
            try:
                if cell.has_style:
                    tgt._style = cell._style
            except Exception:
                pass


def update_excel(rows, output_path):
    """
    rows : liste de dicts extraits du Word (chaque dict doit contenir 'Log_ID' et 'action' au moins)
    output_path : chemin vers le fichier Excel existant (sera remplacé par une nouvelle version)
    """
    version = get_next_version(output_path)
    temp_path = output_path.replace(".xlsx", "_temp.xlsx")

    if not os.path.exists(output_path):
        print("❌ Le fichier Excel n'existe pas. Veuillez le placer dans le dossier output.")
        return

    wb = load_workbook(output_path)
    # trouver la dernière version Vx (fallback à la dernière feuille si aucune Vx)
    candidate_sheets = [s for s in wb.sheetnames if s.startswith("V") and s[1:].isdigit()]
    if candidate_sheets:
        last_version = sorted(candidate_sheets, key=lambda x: int(x[1:]))[-1]
    else:
        last_version = wb.sheetnames[-1]

    source_sheet = wb[last_version]

    # créer la nouvelle feuille et copier le format (mais pas les valeurs)
    new_sheet = wb.create_sheet(title=version)
    copy_sheet_format(source_sheet, new_sheet)

    # Lire les lignes non vides (au moins une cellule non None)
    data = [
        list(row)
        for row in source_sheet.iter_rows(values_only=True)
        if any(cell is not None and str(cell).strip() != "" for cell in row)
    ]

    if not data:
        headers = []
        df = pd.DataFrame()
    else:
        headers = [str(h).strip() if h is not None else "" for h in data[0]]
        body = data[1:] if len(data) > 1 else []

        # Supprimer les lignes totalement vides du corps
        body = [
            row for row in body
            if any(str(cell).strip() != "" and cell is not None for cell in row)
        ]

        df = pd.DataFrame(body, columns=headers)

    # s'assurer que le dataframe a une colonne Log_ID (sinon la créer)
    if "Log_ID" not in df.columns:
        df["Log_ID"] = ""

    # Appliquer les actions
    for row in rows:
        action = row.pop("action", "NONE")
        log_id = str(row.get("Log_ID", "")).strip()

        if action == "ADD_UPDATE":
            # si la colonne n'existe pas dans df, on l'ajoute
            for col in row.keys():
                if col not in df.columns:
                    df[col] = ""

            # mise à jour ou ajout
            if log_id and (log_id in df["Log_ID"].astype(str).values):
                # mettre à jour les colonnes fournies
                for col, val in row.items():
                    df.loc[df["Log_ID"].astype(str) == log_id, col] = val
            else:
                # ajouter une nouvelle ligne en respectant l'ordre des colonnes
                new_row = {col: row.get(col, "") for col in df.columns}
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

        elif action == "DELETE":
            if "Log_ID" in df.columns and log_id:
                df = df[df["Log_ID"].astype(str) != log_id]

        # sinon action == "NONE" -> on ignore

    # écrire le dataframe dans la nouvelle feuille (en conservant le style déjà copié)
    # écrire l'entête
    for c_idx, h in enumerate(df.columns.tolist(), start=1):
        cell = new_sheet.cell(row=1, column=c_idx)
        cell.value = h

    # Supprimer les lignes vides éventuelles
    df = df.dropna(how="all").reset_index(drop=True)

    # écrire les lignes
    for r_idx, row_vals in enumerate(df.values.tolist(), start=2):
        for c_idx, value in enumerate(row_vals, start=1):
            cell = new_sheet.cell(row=r_idx, column=c_idx)
            cell.value = value

    # sauvegarder dans un fichier temporaire, puis remplacer l'ancien fichier
    wb.save(temp_path)
    wb.close()

    try:
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(temp_path, output_path)
        print(f"✅ Nouvelle version créée : {version}")
    except PermissionError:
        alt = output_path.replace(".xlsx", f"_{version}_new.xlsx")
        shutil.move(temp_path, alt)
        print(f"⚠️ Fichier verrouillé. Sauvegarde sous : {alt}")
