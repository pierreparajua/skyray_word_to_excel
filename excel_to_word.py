from docx import Document
import pandas as pd

from pathlib import Path

from utils import modified_indentation

# Définie les noms des dossiers et des fichiers comme des constantes.
SOURCE_FOLDER = "document_source"
RESULT_FOLDER = "result"
EXCEL_FILE = "EmployersRequirementsExcelMaster.xlsx"

# Définie les chemins des constantes précédentes.
PATH_RESULT = Path(__file__).resolve().parent.parent / RESULT_FOLDER
PATH_DOCUMENT_SOURCE = Path(__file__).resolve().parent.parent / SOURCE_FOLDER
PATH_EXCEL_FILE = Path(__file__).resolve().parent.parent / EXCEL_FILE


def get_df_from_excel() -> tuple:
    """
    Récupère les données du fichier excel source et retourne 3 dictionnaires.
    """
    # Crée un dataframe à partir du fichier Excel
    df = pd.read_excel(PATH_EXCEL_FILE, sheet_name="database", header=1)

    # Modifie les colonnes 'Requirements' et 'Exigences' pour rajouter les tabulations.
    df['Requirement'] = df['Requirement'].map(modified_indentation)
    df['Exigence'] = df['Exigence'].map(modified_indentation)

    # Crée 3 dictionnaires dont la clé est "Doc Reference" et les valeurs sont "Requirement", "Exigence" et
    # "Level".
    zip_requirement = zip(df['Doc Reference'], df['Requirement'])
    zip_exigence = zip(df['Doc Reference'], df['Exigence'])
    zip_level = zip(df['Doc Reference'], df['Level'])
    requirement = dict(zip_requirement)
    exigence = dict(zip_exigence)
    level = dict(zip_level)
    return requirement, exigence, level


def get_files_name(folder_path: Path) -> list:
    """
    Liste tous les fichiers en '.docx' présents dans le dossier en paramètre.
    :param folder_path: Chemin du dossier.
    :return : Liste contenant le nom de tous les fichiers words présents dans le dossier.
    """
    files_name = []
    for f in folder_path.glob("*.docx"):
        files_name.append(f.name)
    return files_name


def create_word_instances(files: list) -> list:
    """
    À partir de la liste des fichiers Words, retourne une liste d'instance 'Document'.
    :param files : Liste de chemin des documents Words.
    :return : Liste d'instance 'Document'.
    """
    word_instances = []
    for file in files:
        word_instances.append(Document(PATH_DOCUMENT_SOURCE / file))
    return word_instances


def get_all_cells(instance: list) -> list:
    """
    Retourne une liste contenant tous les objets 'cells' pour une instance d'un document word.
    :param instance : Liste d'instance 'Document'
    :return : Liste d'objet 'cells'.
    """
    cells = []
    for table in instance.tables:
        for row in table.rows:
            for cell in row.cells:
                cells.append(cell)
    return cells


def create_modified_template(cells: list, content: dict, level: dict):
    """
    Remplit les objets "cells" avec le text extrait de l'Excel.
    :param cells : Liste des objets 'cells' pour une instance.
    :param content : Dictionnaire contenant les 'requirements' ou les 'Exigences'.
    :param level: Dictionnaire contenant les 'level'.
    """
    for cell in cells:
        # Si le contenu de la cellule ('ID') est égal à la clé du dictionnaire,
        # alors la cellule suivante est remplie avec la valeur du dictionnaire 'requirements' ou 'Exigences'.
        if cell.text in content.keys():
            ind = cells.index(cell) + 1
            cells[ind].text = content[cell.text]
        # et la cellule suivante encore est remplie avec la valeur du dictionnaire 'level'.
        if cell.text in level.keys():
            ind = cells.index(cell) + 2
            cells[ind].text = "      " + str(level[cell.text])


def main():
    # Fonction principale du programme.
    # Exécute les fonctions précédentes chronologiquement.
    print("Lancement du programme")
    print("Chargement du fichier excel source ..")
    requirement, exigence, level = get_df_from_excel()
    print("Chargement terminé")
    files = get_files_name(PATH_DOCUMENT_SOURCE)
    print(f"Listes des fichiers word à modifier :")
    for file in files:
        print(f"\t{int(files.index(file)) + 1}: {file}")
    choix = input("Souhaitez vous lancer les modifications ? ( y ou n): ")
    if choix == "y":
        words_instances = create_word_instances(files)
        while True:
            choix_langue = input("Choisissez la anglais ou francais: \n\t1: Anglais \n\t2: Francais \nRéponse: ")
            if choix_langue == "1":
                for instance, file_name in zip(words_instances, files):
                    cells = get_all_cells(instance)
                    create_modified_template(cells, requirement, level)
                    instance.save(PATH_RESULT / file_name)
                break
            elif choix_langue == "2":
                for instance, file_name in zip(words_instances, files):
                    cells = get_all_cells(instance)
                    create_modified_template(cells, exigence, level)
                    instance.save(PATH_RESULT / file_name)
                break
            else:
                print("Vous devez taper '1' ou '2'")
    if choix == "n":
        exit()
