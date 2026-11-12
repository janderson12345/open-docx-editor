# Scores each generated document: how far our lines are from Word's. Lines are matched as tools/word-compare/compare.py
# matches them (in order, by their first 40 characters); a document's score is the share of Word's lines we don't
# reproduce on the same page within 1px, plus our lines Word doesn't have, plus 20 for each page of difference.
# 0 is the same as Word.
# usage: python score.py <wordJsonDir> <oursJsonDir> <out.json>
import re, sys, os, json, difflib

wd, od, outp = sys.argv[1:4]
# Beyond compare.py's: case is ignored (small caps come out of Word's PDF in capitals) and a repeated character counts
# once (Word's PDF draws shadowed text twice, letter by letter), and only the first 16 characters are compared (a
# line's key holds 40 characters as drawn, fewer real ones when they're doubled) -- on both sides, so it only hides
# how text is drawn.
norm = lambda t: re.sub(r'(.)\1+', r'\1', t.replace('‐', '-').replace('‑', '-').casefold())[:16]

def dedupe(lines):
    # Text drawn again within 3px of a line that holds it (emboss and engrave draw an offset copy of their run, which
    # reads as a line of its own): part of that line.
    full = lambda t: re.sub(r'(.)\1+', r'\1', t.casefold())
    return [l for i, l in enumerate(lines) if not any(j != i and k[1] == l[1] and 0 < abs(k[2] - l[2]) <= 3 and full(l[0]) in full(k[0])
                                                    for j, k in enumerate(lines[max(0, i - 4):i + 5], max(0, i - 4)))]

res = {}
for f in sorted(os.listdir(wd)):
    if not f.endswith('.json') or not os.path.exists(os.path.join(od, f)): continue
    w = dedupe(json.load(open(os.path.join(wd, f), encoding='utf8'))); o = dedupe(json.load(open(os.path.join(od, f), encoding='utf8')))
    sm = difflib.SequenceMatcher(None, [norm(a[0]) for a in w], [norm(a[0]) for a in o], autojunk=False)
    pairs = [(w[b.a + i], o[b.b + i]) for b in sm.get_matching_blocks() for i in range(b.size)]
    d = [b[2] - a[2] for a, b in pairs if a[1] == b[1]]
    wp, op = max([a[1] for a in w] or [0]), max([a[1] for a in o] or [0])
    within = sum(abs(v) <= 1 for v in d)
    missing, extra = len(w) - len(pairs), len(o) - len(pairs)
    other_page = len(pairs) - len(d)
    drift = len(d) - within
    score = 100 * (missing + extra + other_page + drift) / max(1, len(w)) + 20 * abs(wp - op)
    # The first of Word's lines we don't reproduce: where to look first.
    first = next((a for a, b in pairs if a[1] != b[1] or abs(b[2] - a[2]) > 1), None)
    hit = {i for b in sm.get_matching_blocks() for i in range(b.a, b.a + b.size)}
    first_missing = next((w[i] for i in range(len(w)) if i not in hit), None)
    res[f[:-5]] = dict(score=round(score, 1), wordPages=wp, ourPages=op, wordLines=len(w), ourLines=len(o), matched=len(pairs),
                       otherPage=other_page, within1px=within, drift=drift, missing=missing, extra=extra,
                       meanAbs=round(sum(abs(v) for v in d) / max(1, len(d)), 1), maxAbs=round(max([abs(v) for v in d] or [0]), 1),
                       firstOff=first and [first[0], first[1], first[2]], firstMissing=first_missing)
json.dump(res, open(outp, 'w', encoding='utf8'), indent=0)
print(f'scored {len(res)} documents')
