from docx import Document


def get_Vorlage():
    pathR =r"RechnungVorlage.docx"
    doc = Document(pathR)
    return doc


