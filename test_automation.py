from docx import Document
import pandas as pd
from pathlib import Path

# Définie les chemins
PATH_RESULT = Path(__file__).resolve().parent / "result"
PATH_TEMPLATE = Path(__file__).resolve().parent / "template"
PATH_DOCUMENT_SOURCE = Path(__file__).resolve().parent / "document_source"

# Crée un dataframe à partir du fichier Excel
df = pd.read_excel(PATH_DOCUMENT_SOURCE / 'EmployersRequirementsExcelMaster.xlsx', sheet_name="database", header=1)

# Crée des dictionnaires avec en clé la "Doc Reference" référence"" et en valeur "Requirement", "Exigence" et "Level"
zip_requirement = zip(df['Doc Reference'], df['Requirement'])
zip_exigence = zip(df['Doc Reference'], df['Exigence'])
zip_level = zip(df['Doc Reference'], df['Level'])
requirement = dict(zip_requirement)
exigence = dict(zip_exigence)
level = dict(zip_level)

# Crée une instance de Document du word template
doc = Document(PATH_TEMPLATE / "Appendix 3.01 - PV modules_template.docx")


def get_all_cells(document: Document):
    """
    :param document: Instance de Document
    :return: Liste contenant tout les  d 'objet cells du document word
    """
    cells = []
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                cells.append(cell)
    return cells


def create_modified_template(datas):
    all_cells = get_all_cells(doc)
    for cell in all_cells:
        if cell.text in requirement.keys():
            ind = all_cells.index(cell) + 1
            all_cells[ind].text = requirement[cell.text]
    doc.save(PATH_RESULT / 'test_result.docx')


# code
create_modified_template(requirement)
