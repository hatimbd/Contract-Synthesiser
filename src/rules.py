# Définition des couleurs et actions
RULES = {
    "RED": "ADD_UPDATE",     # Rouge → ajouter ou mettre à jour la ligne
    "PINK": "DELETE"     # Violet → supprimer la ligne
}

# Mapping openpyxl colors ou python-docx highlight_color
COLOR_MAP = {
    "RED": "RED",
    "PINK": "PINK",
    "FFFF0000": "RED",      # fallback hex
    "FFC0CB": "PINK"
}
