require 'docx'

Numeric.extend(Docx::NumericExt)
doc = Docx::Editor.new_document
doc.define_list_style(:toc, Docx::Numbering.table_of_contents)
doc.set_page_size(width: 4.inches, height: 3.5.inches) # or Docx::Inches * 3.5 (w/o NumericExt)
doc.set_page_margins(top: 0.5.inches, right: 0.5.inches, bottom: 0.5.inches, left: 0.5.inches)
doc.set_page_numbering(start: 0) # Let table of contents be page 0

# Add entries to the table of contents
pg = doc.add_paragraph
pg.set_style(:h2)
pg.add_text('Table of Contents')
def entry(doc, name, depth = 0)
  pg = doc.add_paragraph
  pg.set_font_size(8.pt)
  pg.set_list_style(:toc, indent: depth)
  pg.add_text(name)
end
entry(doc, "Foreword")
entry(doc, "Sections")
entry(doc, "All the stuff", 1)
entry(doc, "What will be told", 2)
entry(doc, "What is being told", 2)
entry(doc, "What was told", 2)
entry(doc, "Afterword")
doc.add_paragraph # an empty line

# Add a publisher note
doc.add_horizontal_rule
text = doc.add_paragraph.add_text('Published by "So and so and Company".')
text.properties.italic.val = Docx::True
doc.add_page_break

# Some example text
pg = doc.add_paragraph
pg.set_style(:h2) # title, subtitle, and h1 .. h6 are in a blank document
pg.add_text("Foreward")

doc.add_paragraph.add_text("Blah blah blah ... ")
doc.add_paragraph # an empty line
doc.add_paragraph.add_text("Ja Ja Ja ... ")
doc.save_as("draft.docx")
