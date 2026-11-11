"""Builds the Aptos stand-in fonts from Arimo (Arial's widths and design; SIL Open Font License 1.1, src/fonts/LICENSE.md:
the fonts made are modified versions of it, under the same license and with names of their own).

  python aptos_standin.py cmap <arimo.woff2>            -> prints {code point: advance in em}, as JSON
  python aptos_standin.py build <spec.json>             -> writes each face's .woff2

spec.json: {"faces": [{"src": arimo.woff2, "out": out.woff2, "name": "...", "scale": s,
                       "widths": {"<code point>": advance in em}}]}

Every letter is drawn at one size (the face's scale, applied by the CSS size-adjust) and stretched or narrowed
sideways only, to the advance Aptos gives it: the letters keep Aptos's widths, so lines break where Word breaks them,
and keep one height. Each code point gets its own glyph (two code points that share one in Arimo, a space and a
no-break space, can take different widths); composite glyphs are flattened first. No hinting, kerning or other
layout tables: the editor turns kerning and ligatures off.
"""
import json
import sys

from fontTools.fontBuilder import FontBuilder
from fontTools.pens.recordingPen import DecomposingRecordingPen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont


def cmap(path):
    font = TTFont(path)
    upm = font['head'].unitsPerEm
    hmtx = font['hmtx']
    print(json.dumps({cp: hmtx[g][0] / upm for cp, g in sorted(font.getBestCmap().items())}))


def build_face(face):
    src = TTFont(face['src'])
    upm = src['head'].unitsPerEm
    best = src.getBestCmap()
    glyphs = src.getGlyphSet()
    hmtx = src['hmtx']
    scale = face['scale']
    order = ['.notdef']
    outlines = {}
    advances = {}
    cmap_out = {}

    def draw(name, xscale):
        rec = DecomposingRecordingPen(glyphs)
        glyphs[name].draw(rec)
        pen = TTGlyphPen(None)
        rec.replay(TransformPen(pen, (xscale, 0, 0, 1, 0, 0)))
        return pen.glyph()

    outlines['.notdef'] = draw('.notdef', 1) if '.notdef' in glyphs else TTGlyphPen(None).glyph()
    advances['.notdef'] = (hmtx['.notdef'][0] if '.notdef' in glyphs else upm // 2, 0)
    for key, width in sorted(face['widths'].items(), key=lambda kv: int(kv[0])):
        cp = int(key)
        if cp not in best:
            continue
        gname = best[cp]
        old = hmtx[gname][0]
        new = round(width / scale * upm)
        xscale = new / old if old > 0 and new > 0 else 1
        name = 'u%04X' % cp
        order.append(name)
        outlines[name] = draw(gname, xscale)
        cmap_out[cp] = name
        advances[name] = (new if old > 0 else old, 0)
    fb =FontBuilder(upm, isTTF=True)
    fb.setupGlyphOrder(order)
    fb.setupCharacterMap(cmap_out)
    fb.setupGlyf(outlines)
    glyf = fb.font['glyf']
    metrics = {}
    for name in order:
        g = glyf[name]
        lsb = 0
        if g.numberOfContours != 0:
            g.recalcBounds(glyf)
            lsb = g.xMin
        metrics[name] = (advances[name][0], lsb)
    fb.setupHorizontalMetrics(metrics)
    hhea = src['hhea']
    fb.setupHorizontalHeader(ascent=hhea.ascent, descent=hhea.descent)
    fb.setupNameTable({'familyName': face['name'], 'styleName': face.get('style', 'Regular'),
                       'psName': face['name'].replace(' ', '') + '-' + face.get('style', 'Regular').replace(' ', ''),
                       'copyright': 'Copyright 2020 The Arimo Project Authors (https://github.com/googlefonts/arimo). '
                                    'Modified (letter widths) by open-docx-editor. Licensed under the SIL Open Font '
                                    'License, Version 1.1 (https://openfontlicense.org).'})
    os2 = src['OS/2']
    fb.setupOS2(sTypoAscender=os2.sTypoAscender, sTypoDescender=os2.sTypoDescender, sTypoLineGap=os2.sTypoLineGap,
                usWinAscent=os2.usWinAscent, usWinDescent=os2.usWinDescent, usWeightClass=os2.usWeightClass,
                fsSelection=os2.fsSelection, version=4)
    fb.font['head'].macStyle = src['head'].macStyle
    fb.setupPost()
    fb.font.flavor = 'woff2'
    fb.save(face['out'])


def main():
    if sys.argv[1] == 'cmap':
        cmap(sys.argv[2])
    else:
        with open(sys.argv[2], encoding='utf-8') as f:
            spec = json.load(f)
        for face in spec['faces']:
            build_face(face)


main()
