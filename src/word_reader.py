from docx import Document
from rules import COLOR_MAP, RULES

def get_row_highlight(row):
    """Retourne la première couleur de surlignage trouvée dans une ligne."""
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                if run.font.highlight_color:
                    try:
                        return run.font.highlight_color.name
                    except:
                        return str(run.font.highlight_color)
    return None

def extract_rows(filepath):
    """Extrait les lignes du tableau et leur action correspondante."""
    doc = Document(filepath)
    table = doc.tables[0]
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    rows = []

    for row in table.rows[1:]:
        data = {headers[i]: row.cells[i].text.strip() for i in range(len(headers))}
        color = get_row_highlight(row)
        mapped = COLOR_MAP.get(color)
        action = RULES.get(mapped, "NONE")
        data["action"] = action
        rows.append(data)

    return rows


#test_word_reader.py
'''import os
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    # Chemins d'entrée et de sortie
input_path = os.path.join(BASE_DIR, "input", "contrat_parametrage.docx")

print(extract_rows(input_path))'''