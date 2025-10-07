import os
from word_reader import extract_cell_changes
from excel_writer import update_excel

if __name__ == "__main__":
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    input_path = os.path.join(BASE_DIR, "input", "contrat_parametrage.docx")
    output_path = os.path.join(BASE_DIR, "output", "parametres_mis_a_jour.xlsx")

    changes = extract_cell_changes(input_path)
    for c in changes:
        print(f"{c['Log_ID']} - {c['column']} ({c['action']}) → {c['new_value']}")

    update_excel(changes, output_path)
