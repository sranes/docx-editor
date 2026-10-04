/**
 * Builds the sample mail merge template as DOCX bytes.
 *
 * The template uses every merge construct the sample verifies:
 *
 * - `{{Name}}` text placeholders, including one that Word-style run splitting cuts in two;
 * - plain text content controls tagged `var:<Name>`;
 * - block content controls tagged `if:<flag>` or `if:!<flag>` (conditional rules);
 * - a table whose template row holds `{{items.<column>}}` placeholders (repeating rows);
 * - a repeating section tagged `repeat:payments` whose item holds `{{payments.<column>}}`;
 * - a Word `MERGEFIELD` complex field;
 * - a placeholder in the default header.
 */

import { strToU8, zipSync } from 'fflate';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const HEADER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';

const run = (text: string, bold = false): string =>
  `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

const paragraph = (...runs: string[]): string => `<w:p>${runs.join('')}</w:p>`;

const inlineVariable = (name: string): string =>
  `<w:sdt><w:sdtPr><w:alias w:val="${name}"/><w:tag w:val="var:${name}"/><w:text/></w:sdtPr>` +
  `<w:sdtContent>${run(`[${name}]`)}</w:sdtContent></w:sdt>`;

const ruleBlock = (rule: string, ...paragraphs: string[]): string =>
  `<w:sdt><w:sdtPr><w:alias w:val="Rule: ${rule}"/><w:tag w:val="if:${rule}"/></w:sdtPr>` +
  `<w:sdtContent>${paragraphs.join('')}</w:sdtContent></w:sdt>`;

/** A repeating section tagged `repeat:<list>`, holding one item with the given paragraphs. */
const repeatingSection = (list: string, ...paragraphs: string[]): string =>
  `<w:sdt><w:sdtPr><w:alias w:val="Repeat: ${list}"/><w:tag w:val="repeat:${list}"/><w15:repeatingSection/></w:sdtPr>` +
  `<w:sdtContent><w:sdt><w:sdtPr><w15:repeatingSectionItem/></w:sdtPr>` +
  `<w:sdtContent>${paragraphs.join('')}</w:sdtContent></w:sdt></w:sdtContent></w:sdt>`;

const mergeField = (name: string): string =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> MERGEFIELD ${name} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  run(`«${name}»`) +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

const cell = (text: string, bold = false): string =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${paragraph(run(text, bold))}</w:tc>`;

const border = (side: string): string =>
  `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;

const itemsTable =
  `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>` +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('') +
  `</w:tblBorders></w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
  `<w:tr>${cell('Item', true)}${cell('Qty', true)}${cell('Price', true)}</w:tr>` +
  `<w:tr>${cell('{{items.name}}')}${cell('{{items.qty}}')}${cell('{{items.price}}')}</w:tr>` +
  `</w:tbl>`;

const body = [
  paragraph(run('Dear {{FirstName}} {{LastName}},')),
  // Word often splits a placeholder over runs with different formatting.
  paragraph(run('Customer ID: {{Custo'), run('merId}}', true)),
  paragraph(run('City: '), inlineVariable('City')),
  ruleBlock('isVip', paragraph(run('As a VIP member, you get 20% off your next order.', true))),
  ruleBlock('!isVip', paragraph(run('Join our VIP program to get 20% off.'))),
  ruleBlock('hasBalance', paragraph(run('Your outstanding balance is {{Balance}}.'))),
  paragraph(run('Account number: '), mergeField('AccountNo')),
  paragraph(run('Your order:')),
  itemsTable,
  paragraph(run('Payments received:')),
  repeatingSection('payments', paragraph(run('{{payments.date}}: {{payments.amount}}'))),
  paragraph(run('Regards,')),
  paragraph(run('{{Sender}}', true)),
].join('');

const documentXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w15="${W15}"><w:body>${body}` +
  `<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/>` +
  `<w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
  `</w:sectPr></w:body></w:document>`;

const headerXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:hdr xmlns:w="${W}" xmlns:r="${R}">${paragraph(run('{{Company}} | Confidential'))}</w:hdr>`;

/** Returns a new copy of the sample template on each call. */
export function sampleTemplateBytes(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CT}">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${MAIN}"/>` +
        `<Override PartName="/word/header1.xml" ContentType="${HEADER}"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL}">` +
        `<Relationship Id="rId1" Type="${OD}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL}">` +
        `<Relationship Id="rIdHeader1" Type="${OD}/header" Target="header1.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(documentXml),
    'word/header1.xml': strToU8(headerXml),
  });
}
