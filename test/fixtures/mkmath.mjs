// Builds test/fixtures/math.docx: Word equations (Office Math) inline and on lines of their own: a power, the
// quadratic formula (fraction, root), a sum and an integral with limits, a matrix in brackets, a function, a limit,
// an accent and an equation array.
import '../setup.js';
import fs from 'fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/index.js';
const MNS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const r = (t, sty) => `<m:r>${sty ? `<m:rPr><m:sty m:val="${sty}"/></m:rPr>` : ''}<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t xml:space="preserve">${t}</m:t></m:r>`;
const e = (...x) => `<m:e>${x.join('')}</m:e>`;
const sup = (b, s) => `<m:sSup>${e(b)}<m:sup>${s}</m:sup></m:sSup>`;
const sub = (b, s) => `<m:sSub>${e(b)}<m:sub>${s}</m:sub></m:sSub>`;
const frac = (n, d) => `<m:f><m:num>${n}</m:num><m:den>${d}</m:den></m:f>`;
const sqrt = (x) => `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${e(x)}</m:rad>`;
const nary = (chr, loc, lo, hi, body) => `<m:nary><m:naryPr><m:chr m:val="${chr}"/><m:limLoc m:val="${loc}"/></m:naryPr><m:sub>${lo}</m:sub><m:sup>${hi}</m:sup>${e(body)}</m:nary>`;
const delim = (beg, end, ...parts) => `<m:d><m:dPr><m:begChr m:val="${beg}"/><m:endChr m:val="${end}"/></m:dPr>${parts.map((p) => e(p)).join('')}</m:d>`;
const matrix = (rows) => `<m:m>${rows.map((rw) => `<m:mr>${rw.map((c) => e(c)).join('')}</m:mr>`).join('')}</m:m>`;
const func = (name, arg) => `<m:func><m:fName>${r(name, 'p')}</m:fName>${e(arg)}</m:func>`;
const limLow = (base, lim) => `<m:limLow>${e(base)}<m:lim>${lim}</m:lim></m:limLow>`;
const acc = (chr, x) => `<m:acc><m:accPr><m:chr m:val="${chr}"/></m:accPr>${e(x)}</m:acc>`;
const eqArr = (...rows) => `<m:eqArr>${rows.map((x) => e(x)).join('')}</m:eqArr>`;
const omath = (...x) => `<m:oMath>${x.join('')}</m:oMath>`;
const para = (...x) => `<m:oMathPara>${omath(...x)}</m:oMathPara>`;
const P = (...x) => `<w:p>${x.join('')}</w:p>`;
const R = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const body = [
  P(R('Equations')),
  P(R('The area of a circle is '), omath(r('A=π'), sup(r('r'), r('2'))), R(', inline with the text.')),
  P(R('The quadratic formula:')),
  P(para(r('x='), frac(r('-b±') + sqrt(sup(r('b'), r('2')) + r('-4ac')), r('2a')))),
  P(R('A sum and an integral:')),
  P(para(nary('∑', 'undOvr', r('i=1'), r('n'), r('i')), r('='), frac(r('n') + delim('(', ')', r('n+1')), r('2')))),
  P(para(nary('∫', 'subSup', r('0'), r('1'), sup(r('x'), r('2')) + r('dx')), r('='), frac(r('1'), r('3')))),
  P(R('A matrix, a function, a limit and an accent:')),
  P(para(r('M='), delim('[', ']', matrix([[r('1'), r('0')], [r('0'), r('1')]])))),
  P(omath(func('sin', r('θ')), r('+'), limLow(r('lim', 'p'), r('x→0')), frac(func('sin', r('x')), r('x')), r('=1')), R(' and '), omath(acc('̂', r('x')), r('+'), sub(r('a'), r('1')))),
  P(para(eqArr(r('x+y=2'), r('x-y=0')))),
  P(R('End of equations.')),
].join('');
const files = unzipSync(blankDocxBytes());
let doc = strFromU8(files['word/document.xml']);
if (!/xmlns:m=/.test(doc)) doc = doc.replace(/<w:document\b/, `<w:document ${MNS}`);
files['word/document.xml'] = strToU8(doc.replace(/<w:body>/, `<w:body>${body}`));
fs.writeFileSync(new URL('./math.docx', import.meta.url), zipSync(files));
