// Builds test/fixtures/binding.docx: content controls bound to custom XML (w:dataBinding), in the body and the header,
// showing their placeholders (the values are in customXml/item1.xml, as a document filled by another program leaves
// them). `bindingDocx` makes others like it for the unit tests (test/databind.test.js).
import '../setup.js';
import fs from 'node:fs';
import { pathToFileURL } from 'node:url';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { blankDocxBytes } from '../../src/ooxml/template.js';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const PH = '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>.</w:t></w:r>';

/** A bound control: { xpath, store?, prefixes?, kind (w:sdtPr's kind element), content, plc, rPr, tag, docPart }. */
export function bound({ xpath, store = '', prefixes = '', kind = '<w:text/>', content = PH, plc = true, rPr = '<w:rPr><w:b/></w:rPr>', tag = '', docPart = '' }) {
  const db = `<w:dataBinding${prefixes ? ` w:prefixMappings="${prefixes}"` : ''} w:xpath="${xpath}"${store ? ` w:storeItemID="${store}"` : ''}/>`;
  return `<w:sdt><w:sdtPr>${rPr}${tag ? `<w:tag w:val="${tag}"/>` : ''}<w:id w:val="${xpath.length * 7919}"/>${docPart ? `<w:placeholder><w:docPart w:val="${docPart}"/></w:placeholder>` : ''}${plc ? '<w:showingPlcHdr/>' : ''}${kind}${db}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
}

/**
 * A .docx with `body`, an optional default `header`, custom XML parts [{ xml, id? }] (id: its itemProps' itemID), a
 * glossary of placeholders { name: text } and pictures { 'media/x.png': bytes } (the document's rIdP1, rIdP2...).
 */
export function bindingDocx({ body, header = null, parts = [], glossary = null, media = {} }) {
  const f = unzipSync(blankDocxBytes());
  let ct = strFromU8(f['[Content_Types].xml']);
  let rels = strFromU8(f['word/_rels/document.xml.rels']);
  const rel = (id, type, target) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
  parts.forEach((p, i) => {
    const n = i + 1;
    f[`customXml/item${n}.xml`] = strToU8(p.xml);
    rels = rels.replace('</Relationships>', `${rel(`rIdX${n}`, 'customXml', `../customXml/item${n}.xml`)}</Relationships>`);
    if (p.id) {
      f[`customXml/itemProps${n}.xml`] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="${p.id}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>`);
      f[`customXml/_rels/item${n}.xml.rels`] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', 'customXmlProps', `itemProps${n}.xml`)}</Relationships>`);
      ct = ct.replace('</Types>', `<Override PartName="/customXml/itemProps${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>`);
    }
  });
  let doc = strFromU8(f['word/document.xml']).replace(/<w:document\b/, `<w:document xmlns:w14="${W14}"`).replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`);
  if (header != null) {
    f['word/header1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="${W14}">${header}</w:hdr>`);
    rels = rels.replace('</Relationships>', `${rel('rIdH1', 'header', 'header1.xml')}</Relationships>`);
    ct = ct.replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>');
    doc = doc.replace(/<w:sectPr>/, '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>');
  }
  if (glossary) {
    const dp = Object.entries(glossary).map(([name, text]) => `<w:docPart><w:docPartPr><w:name w:val="${name}"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category><w:types><w:type w:val="bbPlcHdr"/></w:types><w:behaviors><w:behavior w:val="content"/></w:behaviors></w:docPartPr><w:docPartBody><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/><w:i/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:docPartBody></w:docPart>`).join('');
    f['word/glossary/document.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts>${dp}</w:docParts></w:glossaryDocument>`);
    rels = rels.replace('</Relationships>', `${rel('rIdG1', 'glossaryDocument', 'glossary/document.xml')}</Relationships>`);
    ct = ct.replace('</Types>', '<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/></Types>');
  }
  Object.entries(media).forEach(([name, bytes], i) => {
    f[`word/${name}`] = bytes;
    rels = rels.replace('</Relationships>', `${rel(`rIdP${i + 1}`, 'image', name)}</Relationships>`);
  });
  f['word/document.xml'] = strToU8(doc);
  f['[Content_Types].xml'] = strToU8(ct);
  f['word/_rels/document.xml.rels'] = strToU8(rels);
  return zipSync(f);
}

const P = (label, x) => `<w:p><w:r><w:t xml:space="preserve">${label}: </w:t></w:r>${x}</w:p>`;
const cellCtl = (xpath) => `<w:sdt><w:sdtPr><w:id w:val="77"/><w:showingPlcHdr/><w:text/><w:dataBinding w:xpath="${xpath}"/></w:sdtPr><w:sdtContent><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p>${PH}</w:p></w:tc></w:sdtContent></w:sdt>`;
export const FIXTURE = {
  body: '<w:p><w:r><w:t>Bound controls</w:t></w:r></w:p>' + P('Contract', bound({ xpath: '/root/number', tag: 'Number' }))
    + P('Order no', bound({ xpath: '/root/order', tag: 'Order' }))
    + P('Order no again', bound({ xpath: '/root[1]/order[1]', tag: 'OrderAgain' }))
    + P('Reference', bound({ xpath: '/root/ref', tag: 'Ref' }))
    + '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
    + `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Title</w:t></w:r></w:p></w:tc>${cellCtl('/root/title')}</w:tr></w:tbl>`
    + '<w:p><w:r><w:t>End.</w:t></w:r></w:p>',
  header: P('Header order', bound({ xpath: '/root/order', tag: 'HeaderOrder' })),
  parts: [{ xml: '<?xml version="1.0" encoding="utf-8"?><root><number>NNC-0042</number><order>first order</order><ref> </ref><title>Bound title</title></root>', id: '{5F1B0E7A-3C2D-4E8F-9A10-B2C3D4E5F607}' }],
};

if (import.meta.url === pathToFileURL(process.argv[1] || '').href) {
  fs.writeFileSync(new URL('./binding.docx', import.meta.url), bindingDocx(FIXTURE));
}
