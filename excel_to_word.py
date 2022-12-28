from docx import Document
import pandas as pd
from pathlib import Path
from utils import modified_indentation

SOURCE_FOLDER = "document_source"
RESULT_FOLDER = "result"
EXCEL_FILE = "EmployersRequirementsExcelMaster.xlsx"

# Définie les chemins
PATH_RESULT = Path(__file__).resolve().parent.parent / RESULT_FOLDER
PATH_DOCUMENT_SOURCE = Path(__file__).resolve().parent.parent / SOURCE_FOLDER
PATH_EXCEL_FILE = Path(__file__).resolve().parent.parent / EXCEL_FILE


def get_df_from_excel():
    # Crée un dataframe à partir du fichier Excel
    df = pd.read_excel(PATH_EXCEL_FILE, sheet_name="database", header=1)

    # Crée les indentations pour les paragraphes
    df['Requirement'] = df['Requirement'].map(modified_indentation)
    df['Exigence'] = df['Exigence'].map(modified_indentation)

    # Crée des dictionnaires avec en clé la "Doc Reference" référence"" et en valeur "Requirement", "Exigence" et
    # "Level"
    zip_requirement = zip(df['Doc Reference'], df['Requirement'])
    zip_exigence = zip(df['Doc Reference'], df['Exigence'])
    zip_level = zip(df['Doc Reference'], df['Level'])
    requirement = dict(zip_requirement)
    exigence = dict(zip_exigence)
    level = dict(zip_level)
    return requirement, exigence, level


# Liste tous les fichiers présents dans le dossier en paramètre
def get_files_name(folder_path):
    files_name = []
    for f in folder_path.glob("*.docx"):
        files_name.append(f.name)
    return files_name


# Crée une instance de Document du word "template"
def create_word_instances(files):
    word_instances = []
    for w in files:
        word_instances.append(Document(PATH_DOCUMENT_SOURCE / w))
    return word_instances


#  Liste contenant tous les objets cells du document word
def get_all_cells(document: Document):
    cells = []
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                cells.append(cell)
    return cells


# remplit les objets "cells" avec le text extrait du excel
def create_modified_template(cells, requirement, exigence, level):
    for cell in cells:
        if cell.text in requirement.keys():
            ind = cells.index(cell) + 1
            cells[ind].text = requirement[cell.text]
        if cell.text in exigence.keys():
            ind = cells.index(cell) + 1
            cells[ind].text = requirement[cell.text]
        if cell.text in level.keys():
            ind = cells.index(cell) + 1
            cells[ind].text = requirement[cell.text]


def main():
    print("Lancement du programme")
    print("Chargement du fichier excel source ..")
    requirement, exigence, level = get_df_from_excel()
    print("Chargement terminé")
    files_name = get_files_name(PATH_DOCUMENT_SOURCE)
    print(f"Listes des fichiers word à modifier :")
    for file in files_name:
        print(f"{int(files_name.index(file)) + 1}: {file}")
    choix = input("Souhaitez vous lancer les modifications ? ( y ou n)")
    if choix == "y":
        words_instances = create_word_instances(files_name)
        for instance, file_name in zip(words_instances, files_name):
            cells = get_all_cells(instance)
            create_modified_template(cells, requirement, exigence, level)
            instance.save(PATH_RESULT / file_name)
    if choix == "n":
        exit()
