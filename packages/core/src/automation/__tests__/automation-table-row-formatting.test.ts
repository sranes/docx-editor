// Rows added with `updateTable` take the formatting of the row they copy.
//
// A merge fills a template row once per item. Without the source row's paragraph and run
// formatting, every added row came out in the document default font, unlike the row above it.

import { describe, expect, test } from 'bun:test';
import { docx, handlesAt, open, roots, savedMainXml } from './support/protocol.ts';
import type { AutomationHandle } from '../protocol.ts';

const formattedCell = (text: string) =>
  `<w:tc><w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="9" w:author="a"><w:pPr/></w:pPrChange></w:pPr>` +
  `<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:tc>`;

const template = docx(
  `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
    `<w:tr>${formattedCell('{{a}}')}${formattedCell('{{b}}')}</w:tr></w:tbl><w:p/>`
);

function tableOf(host: ReturnType<typeof open>): AutomationHandle {
  const { body } = roots(host);
  const [table] = handlesAt(
    host.execute({ operations: [{ op: 'getTables', scope: { body } }] }),
    0
  ) as [AutomationHandle];
  return table;
}

describe('addRows formatting', () => {
  test('added cells keep the paragraph and run formatting of the source row', () => {
    const host = open(template);
    const table = tableOf(host);
    const response = host.execute({
      operations: [
        {
          op: 'updateTable',
          table,
          mutation: { kind: 'addRows', location: 'end', count: 1, values: [['One', 'Two']] },
        },
      ],
    });
    expect(response.ok).toBe(true);
    const xml = savedMainXml(host);
    const rows = xml.match(/<w:tr[ >][^]*?<\/w:tr>/g) ?? [];
    expect(rows).toHaveLength(2);
    const added = rows[1]!;
    expect(added).toContain('<w:jc w:val="center"/>');
    expect(added).toMatch(/<w:r><w:rPr><w:b\/><w:color w:val="FF0000"\/><\/w:rPr><w:t>One<\/w:t>/);
    // Revision history is not copied: the new row did not exist when that change was made.
    expect(added).not.toContain('pPrChange');
  });

  test('added cells without values keep the formatting for text typed later', () => {
    const host = open(template);
    const table = tableOf(host);
    host.execute({
      operations: [
        { op: 'updateTable', table, mutation: { kind: 'addRows', location: 'end', count: 1 } },
      ],
    });
    const added = (savedMainXml(host).match(/<w:tr[ >][^]*?<\/w:tr>/g) ?? [])[1]!;
    expect(added).toMatch(/<w:r><w:rPr><w:b\/>/);
    expect(added).not.toContain('{{a}}');
  });
});
