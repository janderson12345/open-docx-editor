# Compares our lines with Word's per document (the JSON lines.cjs writes for each PDF) and prints a row per document
# and the totals: page counts, lines on the same page, lines within 1px, line steps within 1px.
# usage: python compare.py <wordJsonDir> <oursJsonDir>
import re, sys, os, json, difflib
wd, od = sys.argv[1:3]
T = dict(wordLines=0, matched=0, samePage=0, within1px=0, steps=0, steps1px=0); pc = n = 0
rows = []
for f in sorted(os.listdir(wd)):
    if not f.endswith('.json') or not os.path.exists(os.path.join(od, f)): continue
    w = json.load(open(os.path.join(wd, f), encoding='utf8')); o = json.load(open(os.path.join(od, f), encoding='utf8'))
    # Hyphen variants (U+2010 from a Cambria Math hyphen, U+2011 non-breaking) read as '-': the same line either way.
    norm = lambda t: re.sub(r'([.·-])\1{2,}', r'\1', t.replace('‐', '-').replace('‑', '-'))  # a tab leader: one mark however many characters
    sm = difflib.SequenceMatcher(None, [norm(a[0]) for a in w], [norm(a[0]) for a in o], autojunk=False)
    pairs = [(w[b.a + i], o[b.b + i]) for b in sm.get_matching_blocks() for i in range(b.size)]
    d = [b[2] - a[2] for a, b in pairs if a[1] == b[1]]
    st = [(b1[2] - b0[2]) - (a1[2] - a0[2]) for (a0, b0), (a1, b1) in zip(pairs, pairs[1:]) if a0[1] == b0[1] == a1[1] == b1[1] and a1[2] > a0[2] + 0.5]
    r = dict(wordLines=len(w), matched=len(pairs), samePage=len(d), within1px=sum(abs(v) <= 1 for v in d), steps=len(st), steps1px=sum(abs(v) <= 1 for v in st))
    wp, op = max([a[1] for a in w] or [0]), max([a[1] for a in o] or [0])
    n += 1; pc += wp == op
    for k in T: T[k] += r[k]
    mean = sum(abs(v) for v in d) / max(1, len(d))
    rows.append(f"{f[:-5][:40]:40} pages {wp:>3}/{op:<3} samePage {r['samePage']}/{r['matched']} within1px {r['within1px']} meanAbs {mean:.1f} steps1px {r['steps1px']}/{r['steps']}")
print('\n'.join(rows))
print(f"TOTAL docs {n} page counts {pc}/{n} lines on same page {T['samePage']}/{T['matched']} ({100*T['samePage']/max(1,T['matched']):.1f}%) within 1px {T['within1px']}/{T['samePage']} ({100*T['within1px']/max(1,T['samePage']):.1f}%) steps within 1px {T['steps1px']}/{T['steps']} ({100*T['steps1px']/max(1,T['steps']):.1f}%)")
