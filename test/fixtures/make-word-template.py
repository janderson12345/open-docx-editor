# Fixture from python-docx, whose default template comes from Word itself.
import docx, sys
from docx.shared import Pt, RGBColor
d = docx.Document()
d.add_heading('Word-template document', 0)
d.add_heading('Heading level 1', 1)
p = d.add_paragraph('A paragraph with ')
p.add_run('bold').bold = True
p.add_run(' and ')
r = p.add_run('green italic'); r.italic = True; r.font.color.rgb = RGBColor(0, 0x80, 0)
d.add_paragraph('First bullet', style='List Bullet')
d.add_paragraph('Second bullet', style='List Bullet')
d.add_paragraph('Numbered', style='List Number')
d.add_paragraph('Quoted text', style='Intense Quote')
t = d.add_table(rows=2, cols=3); t.style = 'Light Grid Accent 1'
for i, c in enumerate(t.rows[0].cells): c.text = 'H%d' % i
for i, c in enumerate(t.rows[1].cells): c.text = 'v%d' % i
d.add_page_break()
d.add_paragraph('After break')
d.save(sys.argv[1])
