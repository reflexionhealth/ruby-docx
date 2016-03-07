require 'docx'

doc = Docx::Editor.new_document
pg = doc.add_paragraph
pg.add_text("Hello document!")
doc.save_as("hello.docx")