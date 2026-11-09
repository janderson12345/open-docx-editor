// Legacy form fields (FORMTEXT, FORMCHECKBOX, FORMDROPDOWN): what a begin fldChar's w:ffData says.
import { unescapeXml } from './xml.js';

/** { name, kind: 'text' | 'checkbox' | 'dropdown', checked?, items?, index?, text? }, or null when it isn't a form field. */
export function legacyInfo(xml) {
  if (!/<w:ffData\b/.test(xml)) return null;
  const name = unescapeXml((/<w:name\b[^>]*w:val="([^"]*)"/.exec(xml) || [])[1] || '');
  if (/<w:checkBox\b/.test(xml)) {
    const v = (/<w:checked\b([^>]*)\/>/.exec(xml) || /<w:default\b([^>]*)\/>/.exec(xml.slice(xml.indexOf('<w:checkBox'))) || [])[1];
    const checked = v != null && !/w:val="(0|false)"/.test(v);
    return { name, kind: 'checkbox', checked };
  }
  if (/<w:ddList\b/.test(xml)) {
    const items = [...xml.matchAll(/<w:listEntry\b[^>]*w:val="([^"]*)"/g)].map((m) => unescapeXml(m[1]));
    const i = +((/<w:result\b[^>]*w:val="(\d+)"/.exec(xml) || /<w:ddList>[\s\S]*?<w:default\b[^>]*w:val="(\d+)"/.exec(xml) || [])[1] || 0);
    return { name, kind: 'dropdown', items, index: i, text: items[i] ?? '' };
  }
  return { name, kind: 'text' };
}
