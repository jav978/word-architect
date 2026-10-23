from docx import Document

doc = Document()

doc.add_paragraph("¡Hola! Word-Architect funciona correctamente.")
doc.save("test.docx")