# Scores each generated document: how far our lines are from Word's. Lines are matched as tools/word-compare/compare.py
# matches them (in order, by their first 40 characters); a document's score is the share of Word's lines we don't
# reproduce on the same page within 1px, plus our lines Word doesn't have, plus 20 for each page of difference.
# 0 is the same as Word.
# usage: python score.py <wordJsonDir> <oursJsonDir> <out.json>
import re, sys, os, json, difflib

SPAN_TOL = float(os.environ.get('SPAN_TOL', 3))  # px; see same_span

wd, od, outp = sys.argv[1:4]
# Beyond compare.py's: case is ignored (small caps come out of Word's PDF in capitals) and a repeated character counts
# once (Word's PDF draws shadowed text twice, letter by letter), and only the first 16 characters are compared (a
# line's key holds 40 characters as drawn, fewer real ones when they're doubled) -- on both sides, so it only hides
# how text is drawn.
norm = lambda t: re.sub(r'(.)\1+', r'\1', t.replace('‐', '-').replace('‑', '-').casefold())[:16]

def boxless(lines):
    # Check box characters left out, on both sides: Word draws a legacy check box (FORMCHECKBOX) as lines, with no text
    # in its PDF, where our print writes the box as text (U+2610, U+2612): "☐Yes" against "Yes", 567 lines in place but
    # not matched (..\why-tests\symtext\boxes.pdf). A check box content control is the character in both. A line left
    # under 3 characters goes, as lines.cjs drops one on either side.
    out = []
    for l in lines:
        k = re.sub('[\u2610-\u2612]', '', l[0])
        if len(k) >= 3: out.append([k] + l[1:])
    return out

def dedupe(lines):
    # Text drawn again within 3px of a line that holds it (emboss and engrave draw an offset copy of their run, which
    # reads as a line of its own): part of that line.
    full = lambda t: re.sub(r'(.)\1+', r'\1', t.casefold())
    return [l for i, l in enumerate(lines) if not any(j != i and k[1] == l[1] and 0 < abs(k[2] - l[2]) <= 3 and full(l[0]) in full(k[0])
                                                    for j, k in enumerate(lines[max(0, i - 4):i + 5], max(0, i - 4)))]

def same_span(w, wi, o, oi, tol=SPAN_TOL):
    # Both rows start and end at the same x (lines.cjs's 5th and 6th): text moved along the line (a tab stop in the wrong
    # place, 1889's "SECTION 1.0<tab>SUBCONTRACT TYPE" 19px short) moves its end, and is not taken for another cut. Line
    # files written before lines.cjs gave the x's are not regrouped at all.
    if not all(len(L[i]) > 5 for L, I in ((w, wi), (o, oi)) for i in I): return False
    span = lambda L, I: (min(L[i][4] for i in I), max(L[i][5] for i in I))
    (a0, a1), (b0, b1) = span(w, wi), span(o, oi)
    return abs(a0 - b0) <= tol and abs(a1 - b1) <= tol

def regroup(w, o):
    # One line cut into other pieces on the two sides (lines.cjs splits a baseline where over 12pt separates its text,
    # and a tab stop or cell edge a few points apart puts a gap on one side over 12pt, on the other under): where a row
    # of ours (one page, a baseline within 1px of Word's) holds the same text as Word's row in another number of
    # pieces, ours is cut as Word's, each piece at our row's place. Only the cut changes: the text must be the same
    # (to the 40 characters a piece keeps), so a piece missing or extra on one side still counts, and the place is
    # still scored, within 1px.
    full = lambda t: re.sub(r'(.)\1+', r'\1', t.replace('‐', '-').replace('‑', '-').casefold())
    def rows(L):
        r = {}
        for i, l in enumerate(L): r.setdefault((l[1], l[2]), []).append(i)
        return r
    byPage = {}
    for k, oi in rows(o).items(): byPage.setdefault(k[0], []).append((k, oi))
    used, repl = set(), {}
    for (p, y), wi in rows(w).items():
        for k, oi in byPage.get(p, []):
            if abs(k[1] - y) > 1 or k in used: continue
            a, b = ''.join(full(w[i][0]) for i in wi), ''.join(full(o[i][0]) for i in oi)
            if len(wi) == len(oi):
                if a == b: used.add(k); break
                continue
            # (A piece keeps 40 characters, so the side with fewer may hold less of the line: then it must be cut short.)
            short = (wi if len(a) < len(b) else oi, w if len(a) < len(b) else o)
            m = min(len(a), len(b))
            if m >= 3 and a[:m] == b[:m] and (len(a) == len(b) or any(len(short[1][i][0]) >= 40 for i in short[0])) and same_span(w, wi, o, oi):
                used.add(k)
                repl[oi[0]] = [[w[i][0], p, o[oi[0]][2]] + o[oi[0]][3:] for i in wi]
                for i in oi[1:]: repl[i] = []
                break
    return [x for i, l in enumerate(o) for x in repl.get(i, [l])]

res = {}
for f in sorted(os.listdir(wd)):
    if not f.endswith('.json') or not os.path.exists(os.path.join(od, f)): continue
    w = dedupe(boxless(json.load(open(os.path.join(wd, f), encoding='utf8')))); o = dedupe(boxless(json.load(open(os.path.join(od, f), encoding='utf8'))))
    o = regroup(w, o)
    sm = difflib.SequenceMatcher(None, [norm(a[0]) for a in w], [norm(a[0]) for a in o], autojunk=False)
    pairs = [(w[b.a + i], o[b.b + i]) for b in sm.get_matching_blocks() for i in range(b.size)]
    d = [b[2] - a[2] for a, b in pairs if a[1] == b[1]]
    wp, op = max([a[1] for a in w] or [0]), max([a[1] for a in o] or [0])
    within = sum(abs(v) <= 1 for v in d)
    # The same, by each line's exact position where both files hold it (lines.cjs): the measure above compares baselines
    # rounded to half a point (0.67px), so two lines 0.7px apart can read 1.33px apart, or 1.3px apart read 0.67.
    exact = [b[3] - a[3] for a, b in pairs if a[1] == b[1] and len(a) > 3 and len(b) > 3]
    within_exact = sum(abs(v) <= 1 for v in exact) if len(exact) == len(d) else None
    missing, extra = len(w) - len(pairs), len(o) - len(pairs)
    other_page = len(pairs) - len(d)
    drift = len(d) - within
    score = 100 * (missing + extra + other_page + drift) / max(1, len(w)) + 20 * abs(wp - op)
    # The first of Word's lines we don't reproduce: where to look first.
    first = next((a for a, b in pairs if a[1] != b[1] or abs(b[2] - a[2]) > 1), None)
    hit = {i for b in sm.get_matching_blocks() for i in range(b.a, b.a + b.size)}
    first_missing = next((w[i][:3] for i in range(len(w)) if i not in hit), None)
    res[f[:-5]] = dict(score=round(score, 1), wordPages=wp, ourPages=op, wordLines=len(w), ourLines=len(o), matched=len(pairs),
                       otherPage=other_page, within1px=within, within1pxExact=within_exact, drift=drift, missing=missing, extra=extra,
                       meanAbs=round(sum(abs(v) for v in d) / max(1, len(d)), 1), maxAbs=round(max([abs(v) for v in d] or [0]), 1),
                       firstOff=first and [first[0], first[1], first[2]], firstMissing=first_missing)
json.dump(res, open(outp, 'w', encoding='utf8'), indent=0)
print(f'scored {len(res)} documents')
