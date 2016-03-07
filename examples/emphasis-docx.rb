require 'docx'

doc = Docx::Editor.new_document
pg = doc.add_paragraph
pg.add_text("Hello document!")

pg = doc.add_paragraph
text = pg.add_text("TLA")
text.properties.bold.val = Docx::True
text.properties.color.val = "0000FF"
pg.add_run
pg.write_tab  # tabs in .docx files are not plain text
pg.write_text("Three letter acronym")

doc.save_as("emphasis.docx")
