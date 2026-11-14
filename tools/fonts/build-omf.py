# Writes src/fonts/omf.js: the Microsoft font families the open-metric-fonts packages (OFL, metric-compatible: the
# originals' character widths and line spacing) stand in for, and where each face's web fonts are, so the editor can
# draw a document's Grandview, Bahnschrift or Sitka in them when the computer lacks the original (fontload.js).
#
# usage: python tools/fonts/build-omf.py <open-metric-fonts checkout, with dist/npm built by its packaging>
#
# Read from the checkout: each package's fontconfig aliases (the Microsoft family a family of ours stands in for), our
# fonts' name tables (which file is a family's Regular, Italic, Bold and Bold Italic, as Word links them), their
# vertical metrics (the Microsoft font's line height, which ours copies), and the npm package's index.css (each face's
# WOFF2 files, one per alphabet, and the characters each holds).
import glob
import json
import os
import re
import sys
from xml.etree import ElementTree

from fontTools.ttLib import TTFont

root = sys.argv[1]
here = os.path.dirname(os.path.abspath(__file__))
out_path = os.path.join(here, '..', '..', 'src', 'fonts', 'omf.js')


def slug(stem):
    return re.sub(r'[^a-z0-9]+', '-', stem.lower()).strip('-')


def word_line(font):
    """Word's single line height per em: the typo metrics where the font asks for them (USE_TYPO_METRICS), else the
    Windows ascent and descent plus what the hhea line gap adds beyond them (GDI's external leading). The editor's own
    table (props.js LINE) has Calibri 1.2207, Arial 1.1499, Corbel 1.2207 and Aptos 1.2207 by this rule."""
    os2, hhea, upm = font['OS/2'], font['hhea'], font['head'].unitsPerEm
    if os2.fsSelection & (1 << 7):
        return (os2.sTypoAscender - os2.sTypoDescender + os2.sTypoLineGap) / upm
    win = os2.usWinAscent + os2.usWinDescent
    gap = max(0, hhea.lineGap - (win - (hhea.ascent - hhea.descent)))
    return (win + gap) / upm


version = None
ranges = {}  # alphabet -> unicode-range
families = {}  # Microsoft family -> [package, line, faces]
for pkg_dir in sorted(glob.glob(os.path.join(root, 'fonts', '*'))):
    pkg = os.path.basename(pkg_dir)
    npm = os.path.join(root, 'dist', 'npm', pkg)
    meta = json.load(open(os.path.join(npm, 'package.json'), encoding='utf8'))
    if version and meta['version'] != version:
        raise SystemExit(f'{pkg} is {meta["version"]}, others {version}')
    version = meta['version']
    # Each WOFF2 file of the package: its face's file stem, its alphabet and the characters it holds.
    css = open(os.path.join(npm, 'index.css'), encoding='utf8').read()
    subsets = {}  # face file stem -> [alphabets]
    for name, body in re.findall(r'/\* ([\w-]+) \*/\s*@font-face \{(.*?)\}', css, re.S):
        src = re.search(r'url\(\./files/([\w-]+)\.woff2\)', body).group(1)
        rng = re.search(r'unicode-range: ([^;]+);', body).group(1).replace(' ', '')
        if not src.endswith('-' + name):
            raise SystemExit(f'{pkg}: {src} is not its alphabet {name}')
        # (An alphabet covering the same characters in every package is listed once; one that doesn't, as each
        # package's "other": its remaining characters, under "<package>:<alphabet>".)
        if ranges.setdefault(name, rng) != rng:
            ranges[f'{pkg}:{name}'] = rng
        subsets.setdefault(src[: -len(name) - 1], []).append(name)
    # Our families' faces as Windows links them (name ID 1 and 2): Regular and Italic at 400, Bold and Bold Italic at 700.
    linked = {}  # our family -> [(weight, italic, stem)]
    lines = {}
    regulars = set()  # families whose line is their Regular's
    for ttf in sorted(glob.glob(os.path.join(pkg_dir, '*.ttf'))):
        font = TTFont(ttf, lazy=True)
        fam = font['name'].getDebugName(1)
        stem = slug(os.path.splitext(os.path.basename(ttf))[0])
        if stem not in subsets:
            raise SystemExit(f'{pkg}: no web fonts for {stem}')
        # (Its style bits, as Windows links a family's four faces: Lucida Bright's bold is "Demibold", Lucida Sans
        # Typewriter's italic "Oblique".)
        sel = font['OS/2'].fsSelection
        weight = 700 if sel & 0x20 else 400
        linked.setdefault(fam, []).append((weight, bool(sel & 1), stem))
        # (The line height from the family's Regular face, as Word's paragraph lines take it: a Bold can have taller
        # metrics, Book Antiqua's and Century Gothic's do. Another face only where there's no Regular.)
        regular = not (sel & 0x21)
        if regular or fam not in lines:
            if regular or fam not in regulars:
                lines[fam] = round(word_line(font), 4)
            if regular:
                regulars.add(fam)
    # The Microsoft families ours stand in for (the package's fontconfig aliases).
    for conf in glob.glob(os.path.join(pkg_dir, 'fontconfig', '*.conf')):
        for alias in ElementTree.parse(conf).getroot().iter('alias'):
            ms, ours = alias.find('family').text, alias.find('accept/family').text
            if ours not in linked:
                # (A name for the whole typographic family, Sitka or Walbaum, whose fonts are "<name> Text", "<name>
                # Display"...: its text size, the one for body text.)
                sizes = sorted(f for f in linked if f.startswith(ours + ' ') and len(linked[f]) > 0)
                pick = next((f for f in sizes if f == ours + ' Text'), sizes[0] if sizes else None)
                if not pick:
                    raise SystemExit(f'{pkg}: {ms} -> {ours}, which none of its fonts is')
                print(f'  {ms} -> {ours}: its fonts are {sizes}; using {pick}')
                ours = pick
            faces = sorted(linked[ours])
            families[ms] = [pkg, lines[ours], [[w, int(i), s, ','.join(subsets[s])] for w, i, s in faces]]

js = [
    '// Generated by tools/fonts/build-omf.py from open-metric-fonts ' + version + ' (https://github.com/open-metric-fonts/open-metric-fonts):',
    "// the Microsoft families its packages stand in for, with the originals' character widths and line spacing. Don't edit:",
    '// run the script again. Loaded by fontload.js only for fonts a document uses that the computer lacks.',
    '',
    f"export const OMF_VERSION = '{version}';",
    '',
    '/** The characters each alphabet\'s WOFF2 file holds (its @font-face unicode-range); "<package>:<alphabet>" where a',
    ' * package\'s differs (its "other": the characters no alphabet above has). */',
    'export const OMF_RANGES = ' + json.dumps(ranges, indent=2) + ';',
    '',
    "/** Microsoft family -> [package (@open-metric-fonts/<package>), Word's line height per em, faces: [weight, italic,",
    ' * file stem, alphabets]], the faces as Word links them (Regular and Italic 400, Bold and Bold Italic 700). */',
    'export const OMF = {',
]
for ms in sorted(families, key=str.lower):
    js.append(f'  {json.dumps(ms)}: {json.dumps(families[ms], separators=(",", ":"))},')
js.append('};')
open(out_path, 'w', encoding='utf8', newline='\n').write('\n'.join(js) + '\n')
print(f'{len(families)} Microsoft families from {len(glob.glob(os.path.join(root, "fonts", "*")))} packages, version {version}, alphabets {list(ranges)}')
