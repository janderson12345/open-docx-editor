# Word's lines and ours for one document, side by side: text, page, baseline, and the difference.
# usage: python sidebyside.py <runDir> <run> <document name>
import json, os, sys, difflib, re
run_dir, run, name = sys.argv[1:4]
w = json.load(open(os.path.join(run_dir, 'lines-word', name + '.json'), encoding='utf8'))
o = json.load(open(os.path.join(run_dir, f'lines-{run}', name + '.json'), encoding='utf8'))
norm = lambda t: re.sub(r'(.)\1+', r'\1', t.casefold())[:16]
sm = difflib.SequenceMatcher(None, [norm(a[0]) for a in w], [norm(a[0]) for a in o], autojunk=False)
for tag, i1, i2, j1, j2 in sm.get_opcodes():
    for k in range(max(i2 - i1, j2 - j1)):
        a = w[i1 + k] if i1 + k < i2 else None
        b = o[j1 + k] if j1 + k < j2 else None
        d = f'{b[2] - a[2]:+.1f}' if a and b and a[1] == b[1] else ('page' if a and b else '')
        fa = f'{a[0][:34]:34} p{a[1]} {a[2]:6.1f}' if a else ' ' * 47
        fb = f'{b[0][:34]:34} p{b[1]} {b[2]:6.1f}' if b else ''
        print(f'{"=" if tag == "equal" else "!"} {fa} | {fb} {d}')
