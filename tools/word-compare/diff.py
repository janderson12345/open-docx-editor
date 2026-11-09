# Per-document changes between two compare.py reports: page-count match, same-page lines, within-1px lines.
# usage: python diff.py <old report> <new report>
import sys, re
def load(f):
    r = {}
    for line in open(f, encoding='utf8', errors='replace'):
        m = re.match(r'(.{40}) pages +(\d+)/(\d+) +samePage (\d+)/(\d+) within1px (\d+)', line)
        if m: r[m[1].strip()] = (m[2] == m[3], int(m[4]), int(m[6]), f'{m[2]}/{m[3]}')
    return r
a, b = load(sys.argv[1]), load(sys.argv[2])
for k in sorted(a):
    if k not in b or a[k] == b[k]: continue
    (pa, sa, wa, ta), (pb, sb, wb, tb) = a[k], b[k]
    worse = pb < pa or sb < sa or wb < wa
    print(f"{'WORSE' if worse else 'better' if (pb > pa or sb > sa or wb > wa) else 'same':6} {k:40} pages {ta}->{tb} samePage {sa}->{sb} ({sb - sa:+}) within1px {wa}->{wb} ({wb - wa:+})")
