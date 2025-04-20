#Instalar biblioteca pip install python-docx
#Documentação: https://python-docx.readthedocs.io/en/latest/

# Realiza a modificação manualmente, inserindo informações do aluno

from docx import Document
from docx.shared import Pt

word = Document("C:\\Users\\nicho\\PycharmProjects\\Enviando_Email\\Certificado1.docx")# modelo exemplo

estilo = word.styles["Normal"]

for paragrafo in word.paragraphs:
    if "@nome" in paragrafo.text:
        paragrafo.text = "Nichole Furtado"
        fonte = estilo.font
        fonte.name =  "Calibri (Corpo)"
        fonte.size = Pt(24)

word.save("Nichole.docx")

