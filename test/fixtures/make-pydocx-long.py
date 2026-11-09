# Multi-page fixture with a header and a "Page X of Y" footer (complex fields), made with python-docx.
import docx, sys, random
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
random.seed(7)
d = docx.Document()
sec = d.sections[0]
sec.header.paragraphs[0].text = 'Acme Corp - Annual Review'

def field(par, instr):
    for kind, extra in (('begin', None), ('instr', instr), ('separate', None), ('text', '1'), ('end', None)):
        r = par.add_run()
        if kind == 'instr':
            t = OxmlElement('w:instrText'); t.set(qn('xml:space'), 'preserve'); t.text = ' %s ' % instr; r._r.append(t)
        elif kind == 'text':
            r.text = extra
        else:
            f = OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'), kind); r._r.append(f)

fp = sec.footer.paragraphs[0]
fp.add_run('Page ')
field(fp, 'PAGE')
fp.add_run(' of ')
field(fp, 'NUMPAGES')
words = 'revenue grew in every region while operating costs stayed flat and the team shipped four major releases'.split()
d.add_heading('Annual Review', 0)
for s in range(6):
    d.add_heading('Chapter %d' % (s + 1), 1)
    for k in range(5):
        d.add_paragraph(' '.join(random.choice(words) for _ in range(random.randint(40, 90))).capitalize() + '.')
    if s == 2:
        d.add_page_break()
d.save(sys.argv[1])
