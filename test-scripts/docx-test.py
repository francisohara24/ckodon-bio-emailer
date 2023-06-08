from docx import Document

text = """Ah you think darkness is your ally.
You merely adopted it.
I was born in it.
Molded by it.
"""

doc = Document()
doc.add_paragraph(text)
doc.save("../data/document.docx")
