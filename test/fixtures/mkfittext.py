# Builds test/fixtures/fittext.docx: runs fitted to a width (w:fitText): spread ("Name" to 1in), squeezed, two runs as
# one group, and the text after each, which Word starts exactly that width on (its PDF of this file).
# usage: python test/fixtures/mkfittext.py test/fixtures/fittext.docx
import sys, zipfile, re
src = sys.argv[2] if len(sys.argv) > 2 else 'test/fixtures/rich.docx'
def p(inner): return f'<w:p>{inner}</w:p>'
def r(t, rpr=''): return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{t}</w:t></w:r>'
body = ''.join([
    p(r('Name', '<w:fitText w:val="1440" w:id="1"/>') + r('|after')),
    p(r('Phone number', '<w:fitText w:val="2880" w:id="2"/>') + r('|after')),
    p(r('Ab', '<w:fitText w:val="1440" w:id="3"/>') + r('cd', '<w:b/><w:fitText w:val="1440" w:id="3"/>') + r('|after')),
    p(r('A long label squeezed', '<w:fitText w:val="1080" w:id="4"/>') + r('|after')),
    p(r('Two words', '<w:fitText w:val="2160" w:id="5"/>') + r('|after')),
])
z = zipfile.ZipFile(src)
with zipfile.ZipFile(sys.argv[1], 'w', zipfile.ZIP_DEFLATED) as zo:
    for n in z.namelist():
        d = z.read(n)
        if n == 'word/document.xml':
            s = re.sub(r'<w:body>[\s\S]*?(<w:sectPr)', lambda m: '<w:body>' + body + m.group(1), d.decode(), count=1)
            d = s.encode()
        zo.writestr(n, d)
