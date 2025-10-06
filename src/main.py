import os
from word_reader import extract_rows
from excel_writer import update_excel

if __name__ == "__main__":
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    input_path = os.path.join(BASE_DIR, "input", "contrat_parametrage.docx")
    output_path = os.path.join(BASE_DIR, "output", "parametres_mis_a_jour.xlsx")

    rows = extract_rows(input_path)
    for row in rows:
        print(f"{row['Log_ID']} → action: {row['action']}")

    update_excel(rows, output_path)
