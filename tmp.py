# -- coding: utf-8 --
from docx import Document

doc = Document("demo.docx")

for style in doc.styles:
    print(style.name)

