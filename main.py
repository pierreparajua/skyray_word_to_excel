from docx import Document
import pandas as pd
from pathlib import Path
from utils import modified_indentation

SOURCE_FOLDER = "document_source"
RESULT_FOLDER = "result"


# Définie les chemins
PATH_RESULT = Path(__file__).resolve().parent / RESULT_FOLDER
PATH_DOCUMENT_SOURCE = Path(__file__).resolve().parent / SOURCE_FOLDER

def get_df_from_excel():
    # Crée un dataframe à partir du fichier Excel
    df = pd.read_excel('EmployersRequirementsExcelMaster.xlsx', sheet_name="database", header=1)

    # recré les indentations pour les paragraphes
    df['Requirement']= df['Requirement'].map(modified_indentation)
    df['Exigence']= df['Exigence'].map(modified_indentation)


    # Crée des dictionnaires avec en clé la "Doc Reference" référence"" et en valeur "Requirement", "Exigence" et "Level"
    zip_requirement = zip(df['Doc Reference'], df['Requirement'])
    zip_exigence = zip(df['Doc Reference'], df['Exigence'])
    zip_level = zip(df['Doc Reference'], df['Level'])
    requirement = dict(zip_requirement)
    exigence = dict(zip_exigence)
    level = dict(zip_level)
    return requirement, exigence, level



# Liste tous les fichiers présent dans le dossier en paramètre
def get_files_name(folder_path):
    files_name = []
    for f in folder_path.glob("*.docx"):
        files_name.append(f.name)
    return files_name

# Crée une instance de Document du word template
def create_word_instances(files):
    word_instances = []
    for w in files:
        word_instances.append(Document(PATH_DOCUMENT_SOURCE / w))
    return word_instances

#  Liste contenant tout les objets cells du document word
def get_all_cells(document: Document):
    cells = []
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                cells.append(cell)
    return cells

# remplit les oblets cells avec le text extrait du excel
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
    



if __name__ == "__main__":
    requirement, exigence, level = get_df_from_excel()
    files_name = get_files_name(PATH_DOCUMENT_SOURCE)
    words_instances = create_word_instances(files_name)   
    for instance, file_name in zip(words_instances, files_name):
        cells = get_all_cells(instance)
        create_modified_template(cells, requirement, exigence, level)
        instance.save(PATH_RESULT / file_name)


