#excel_writer.py
import os
import shutil
import pandas as pd
from openpyxl import load_workbook, Workbook
from datetime import datetime

def get_next_version(filepath):
    if not os.path.exists(filepath):
        return "V1"
    wb = load_workbook(filepath)
    versions = [int(s[1:]) for s in wb.sheetnames if s.startswith("V") and s[1:].isdigit()]
    wb.close()
    return f"V{max(versions)+1}" if versions else "V1"

def update_excel(rows, output_path):
    version = get_next_version(output_path)
    temp_path = output_path.replace(".xlsx", "_temp.xlsx")

    # Charger le fichier existant ou Print une erreur si le fichier n'existe pas ou qu'on peut pas le lire
    if os.path.exists(output_path):
        wb = load_workbook(output_path)
        last_version = sorted([s for s in wb.sheetnames if s.startswith("V")], key=lambda x: int(x[1:]))[-1]
        sheet = wb[last_version]
        # Lecture propre des données
        data = list(sheet.values)
        data = [list(row) for row in data if any(cell is not None for cell in row)]  # ignore lignes vides
        headers = [str(h).strip() if h is not None else "" for h in data[0]]

        df = pd.DataFrame(data[1:], columns=headers)


    else:  #print une erreur si le fichier n'existe pas ou qu'on peut pas le lire
        print("❌ Le fichier Excel n'existe pas ou ne peut pas être lu. Veuillez vérifier le chemin ou les permissions.")
        return


    # Appliquer les actions
    for row in rows:
        action = row.pop("action")
        df["Log_ID"] = df["Log_ID"].astype(str).str.strip()
        log_id = str(row.get("Log_ID")).strip()


        if action == "ADD_UPDATE":
            # Nettoyage du Log_ID
            df["Log_ID"] = df["Log_ID"].astype(str).str.strip()
            log_id = str(row.get("Log_ID")).strip()

            # Vérifie si la ligne existe déjà
            if log_id in df["Log_ID"].values:
                # Mise à jour de la ligne existante
                for col, val in row.items():
                    df.loc[df["Log_ID"] == log_id, col] = val
            else:
                # Ajout d'une nouvelle ligne
                new_row = pd.DataFrame([row])
                new_row = new_row.reindex(columns=df.columns, fill_value="")
                df = pd.concat([df, new_row], ignore_index=True)

        elif action == "DELETE":
            df = df[df["Log_ID"] != log_id]

    # Créer la nouvelle feuille
    ws = wb.create_sheet(title=version)
    for r_idx, row in enumerate([df.columns.tolist()] + df.values.tolist(), start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    wb.save(temp_path)

    try:
        os.remove(output_path)
        shutil.move(temp_path, output_path)
        print(f"✅ Nouvelle version créée : {version}")
    except PermissionError:
        alt = output_path.replace(".xlsx", f"_{version}_new.xlsx")
        shutil.move(temp_path, alt)
        print(f"⚠️ Fichier verrouillé. Sauvegarde sous : {alt}")
