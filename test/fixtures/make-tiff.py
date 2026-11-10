# Builds test/fixtures/tiff.docx: three TIFF pictures as Word embeds them, in the forms they come in: uncompressed
# colour, LZW-compressed colour, and black-and-white with CCITT Group 4 (fax) compression, as scanners write. Each is a
# picture of its own with a caption. Made with Pillow; Word prints it as the reference.
# usage: python test/fixtures/make-tiff.py
import io, os, zipfile
from PIL import Image, ImageDraw

here = os.path.dirname(os.path.abspath(__file__))

def picture(mode, size, compression):
    im = Image.new('RGB', size, (255, 255, 255))
    d = ImageDraw.Draw(im)
    w, h = size
    d.rectangle([4, 4, w - 5, h - 5], outline=(0, 0, 0), width=3)
    d.ellipse([w // 4, h // 4, 3 * w // 4, 3 * h // 4], fill=(31, 95, 168))
    d.rectangle([w // 8, h // 8, w // 3, h // 3], fill=(200, 40, 30))
    d.line([0, h - 1, w - 1, 0], fill=(0, 0, 0), width=4)
    if mode == '1':
        im = im.convert('L').point(lambda v: 255 if v > 160 else 0).convert('1')
    out = io.BytesIO()
    im.save(out, format='TIFF', compression=compression)
    return out.getvalue()

pics = [
    ('Uncompressed colour TIFF', picture('RGB', (240, 160), None)),
    ('LZW colour TIFF', picture('RGB', (240, 160), 'tiff_lzw')),
    ('Black-and-white TIFF, CCITT Group 4 (fax)', picture('1', (300, 200), 'group4')),
]

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
EMU = 9525  # per CSS pixel

def drawing(n, rid, w, h):
    cx, cy = w * EMU, h * EMU
    return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:docPr id="{n}" name="Picture {n}" descr="A test picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="pic{n}.tif"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            f'</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')

body, rels = [], []
for i, (caption, data) in enumerate(pics, 1):
    im = Image.open(io.BytesIO(data))
    body.append(f'<w:p><w:r><w:t>{caption}:</w:t></w:r></w:p>')
    body.append(f'<w:p>{drawing(i, f"rIdT{i}", im.width, im.height)}</w:p>')
    rels.append(f'<Relationship Id="rIdT{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/picture{i}.tif"/>')
body.append('<w:p><w:r><w:t>The end of the document.</w:t></w:r></w:p>')
sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'

# (A styles part of its own, Calibri 11pt with Word's usual spacing: without one, Word sets the text in its current
# default font, which isn't what this document tests.)
styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>')
rels.append('<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')

files = {
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="tif" ContentType="image/tiff"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/_rels/document.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(rels) + '</Relationships>',
    'word/styles.xml': styles,
    'word/document.xml': f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}{sect}</w:body></w:document>',
}
with zipfile.ZipFile(os.path.join(here, 'tiff.docx'), 'w', zipfile.ZIP_DEFLATED) as z:
    for name, text in files.items():
        z.writestr(name, text)
    for i, (_, data) in enumerate(pics, 1):
        z.writestr(f'word/media/picture{i}.tif', data)
print('tiff.docx:', ', '.join(f'{c} {len(d)} bytes' for c, d in pics))
