// Repeating-section items through the automation protocol: add by copy, remove by index.

import { describe, expect, test } from 'bun:test';
import { docx, errorAt, handlesAt, open, roots, savedMainXml } from './support/protocol.ts';
import type { AutomationHandle, AutomationHost } from '../protocol.ts';

const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

const item = (id: number, body: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>${body}</w:sdtContent></w:sdt>`;
const line = (paraId: string, text: string) =>
  `<w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const field = (id: number, text: string) =>
  `<w:p><w:sdt><w:sdtPr><w:id w:val="${id}"/><w:tag w:val="var:Name"/><w:text/></w:sdtPr>` +
  `<w:sdtContent><w:r><w:t>${text}</w:t></w:r></w:sdtContent></w:sdt></w:p>`;

function section(...items: string[]): string {
  return (
    `<w:sdt xmlns:w15="${W15}" xmlns:w14="${W14}"><w:sdtPr><w:id w:val="1"/><w:tag w:val="items"/>` +
    `<w15:repeatingSection/></w:sdtPr><w:sdtContent>${items.join('')}</w:sdtContent></w:sdt>` +
    '<w:p><w:r><w:t>after</w:t></w:r></w:p>'
  );
}

function sectionOf(host: AutomationHost): AutomationHandle {
  const { body } = roots(host);
  const [handle] = handlesAt(
    host.execute({
      operations: [{ op: 'getContentControlsByTag', scope: { body }, tag: 'items' }],
    }),
    0
  ) as [AutomationHandle];
  return handle;
}

const values = (xml: string, pattern: RegExp) => [...xml.matchAll(pattern)].map((m) => m[1]!);

describe('repeating sections', () => {
  test('adding an item copies it after the source with fresh ids and no bookmarks', () => {
    const host = open(
      docx(
        section(
          item(2, line('00000A01', 'Row one') + field(3, 'Ada')),
          item(
            4,
            `<w:bookmarkStart w:id="0" w:name="mark"/>${line('00000A02', 'Row two')}<w:bookmarkEnd w:id="0"/>`
          )
        )
      )
    );
    const contentControl = sectionOf(host);
    const subtype = host.execute({
      operations: [{ op: 'getContentControlSubtype', contentControl }],
    });
    expect(subtype.results[0]).toEqual({
      status: 'ok',
      value: { kind: 'text', text: 'repeatingSection' },
    });

    expect(
      host.execute({ operations: [{ op: 'addRepeatingSectionItem', contentControl, index: 0 }] }).ok
    ).toBe(true);
    expect(
      host.execute({ operations: [{ op: 'addRepeatingSectionItem', contentControl }] }).ok
    ).toBe(true);
    const xml = savedMainXml(host);
    expect(values(xml, /<w:t>([^<]*)<\/w:t>/g)).toEqual([
      'Row one',
      'Ada',
      'Row one',
      'Ada',
      'Row two',
      'Row two',
      'after',
    ]);
    const ids = values(xml, /<w:id w:val="(\d+)"\/>/g);
    expect(new Set(ids).size).toBe(ids.length);
    const paraIds = values(xml, /w14:paraId="([0-9A-F]+)"/g);
    expect(new Set(paraIds).size).toBe(paraIds.length);
    expect(values(xml, /w:name="(mark)"/g)).toEqual(['mark']);
  });

  test('removing an item by index, and refusing to remove the last one', () => {
    const host = open(
      docx(section(item(2, line('00000A01', 'one')), item(3, line('00000A02', 'two'))))
    );
    const contentControl = sectionOf(host);
    expect(
      host.execute({ operations: [{ op: 'removeRepeatingSectionItem', contentControl, index: 0 }] })
        .ok
    ).toBe(true);
    expect(values(savedMainXml(host), /<w:t>([^<]*)<\/w:t>/g)).toEqual(['two', 'after']);
    const last = host.execute({
      operations: [{ op: 'removeRepeatingSectionItem', contentControl, index: 0 }],
    });
    expect(errorAt(last, 0)).toBe('unsupported-content');
    const missing = host.execute({
      operations: [{ op: 'removeRepeatingSectionItem', contentControl, index: 5 }],
    });
    expect(errorAt(missing, 0)).toBe('invalid-offset');
  });

  test('refuses to copy an item holding a comment, and a control that is not a section', () => {
    const host = open(
      docx(
        section(
          item(
            2,
            `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>noted</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p>`
          )
        ) + field(9, 'plain')
      )
    );
    const contentControl = sectionOf(host);
    expect(
      errorAt(host.execute({ operations: [{ op: 'addRepeatingSectionItem', contentControl }] }), 0)
    ).toBe('unsupported-content');
    const { body } = roots(host);
    const [plain] = handlesAt(
      host.execute({
        operations: [{ op: 'getContentControlsByTag', scope: { body }, tag: 'var:Name' }],
      }),
      0
    ) as [AutomationHandle];
    expect(
      errorAt(
        host.execute({ operations: [{ op: 'addRepeatingSectionItem', contentControl: plain }] }),
        0
      )
    ).toBe('unsupported-content');
  });

  test('is solitary in its batch', () => {
    const host = open(docx(section(item(2, line('00000A01', 'one')))));
    const contentControl = sectionOf(host);
    const response = host.execute({
      operations: [
        { op: 'addRepeatingSectionItem', contentControl },
        { op: 'addRepeatingSectionItem', contentControl },
      ],
    });
    expect(response.ok).toBe(false);
  });
});
