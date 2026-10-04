// Block-level content controls over whole paragraphs, and removing them with their content.
//
// A rule in a template wraps whole paragraphs. It must be a BLOCK control, so that removing it
// with its content leaves no empty paragraph behind, which an inline control cannot promise.

import { describe, expect, test } from 'bun:test';
import {
  cell,
  docx,
  errorAt,
  open,
  p,
  paragraphTexts,
  roots,
  row,
  savedMainXml,
  table,
} from './support/protocol.ts';
import type { AutomationHandle, AutomationHost } from '../protocol.ts';

function paragraphs(host: AutomationHost, body: AutomationHandle): readonly AutomationHandle[] {
  const result = host.execute({ operations: [{ op: 'getParagraphs', body }] }).results[0];
  if (result?.status !== 'ok' || result.value.kind !== 'handles') throw new Error('no paragraphs');
  return result.value.handles;
}

function insert(
  host: AutomationHost,
  start: [AutomationHandle, number],
  end: [AutomationHandle, number],
  subtype: 'richText' | 'plainText' = 'richText'
) {
  return host.execute({
    operations: [
      {
        op: 'insertContentControl',
        span: {
          start: { paragraph: start[0], offset: start[1] },
          end: { paragraph: end[0], offset: end[1] },
        },
        subtype,
        tag: 'if:rule',
        returnHandle: true,
      },
    ],
  });
}

describe('insertContentControl over whole paragraphs', () => {
  test('wraps the paragraphs and a table between them in one block control', () => {
    const host = open(
      docx(p('One') + p('Two') + table(row(cell(p('Cell')))) + p('Three') + p('Four'))
    );
    const { body } = roots(host);
    const [one, , , three] = paragraphs(host, body) as AutomationHandle[];
    // From the start of "One" to the end of "Three", across the table.
    const response = insert(host, [one!, 0], [three!, 5]);
    expect(response.ok).toBe(true);
    const xml = savedMainXml(host);
    expect(xml).toMatch(
      /<w:sdt><w:sdtPr>.*if:rule.*<w:sdtContent><w:p[ >].*One.*Two.*<w:tbl>.*Cell.*<\/w:tbl>.*Three.*<\/w:p><\/w:sdtContent><\/w:sdt><w:p[ >][^]*Four/
    );
    expect(paragraphTexts(host, body)).toEqual(['One', 'Two', 'Cell', 'Three', 'Four']);
  });

  test('an end at offset 0 of the next paragraph wraps up to the paragraph before it', () => {
    const host = open(docx(p('One') + p('Two')));
    const { body } = roots(host);
    const [one, two] = paragraphs(host, body) as [AutomationHandle, AutomationHandle];
    expect(insert(host, [one, 0], [two, 0]).ok).toBe(true);
    expect(savedMainXml(host)).toMatch(
      /<w:sdtContent><w:p[ >][^<]*(?:<(?!\/w:p>)[^<]*)*One.*<\/w:sdtContent><\/w:sdt><w:p[ >].*Two/
    );
  });

  test('refuses a span that cuts a paragraph mid-text, and plain text over paragraphs', () => {
    const host = open(docx(p('One') + p('Two')));
    const { body } = roots(host);
    const [one, two] = paragraphs(host, body) as [AutomationHandle, AutomationHandle];
    const before = savedMainXml(host);
    expect(errorAt(insert(host, [one, 1], [two, 3]), 0)).toBe('unsupported-content');
    expect(errorAt(insert(host, [one, 0], [two, 2]), 0)).toBe('unsupported-content');
    expect(errorAt(insert(host, [one, 0], [two, 3], 'plainText'), 0)).toBe('unsupported-content');
    expect(savedMainXml(host)).toBe(before);
  });

  test('refuses blocks that are not siblings, such as a body paragraph and a cell paragraph', () => {
    const host = open(docx(p('One') + table(row(cell(p('Cell'))))));
    const { body } = roots(host);
    const [one, inCell] = paragraphs(host, body) as [AutomationHandle, AutomationHandle];
    const response = insert(host, [one, 0], [inCell, 4]);
    expect(response.ok).toBe(false);
  });

  test('deleting the block control with its content removes the paragraphs entirely', () => {
    const host = open(docx(p('Keep') + p('Drop one') + p('Drop two') + p('End')));
    const { body } = roots(host);
    const [, first, second] = paragraphs(host, body) as AutomationHandle[];
    const created = insert(host, [first!, 0], [second!, 8]).results[0];
    if (created?.status !== 'ok' || created.value.kind !== 'handle') throw new Error('not created');
    const removed = host.execute({
      operations: [
        { op: 'deleteContentControl', contentControl: created.value.handle, keepContent: false },
      ],
    });
    expect(removed.ok).toBe(true);
    expect(paragraphTexts(host, body)).toEqual(['Keep', 'End']);
  });

  test('a span inside one paragraph still makes an inline control', () => {
    const host = open(docx(p('One')));
    const { body } = roots(host);
    const [one] = paragraphs(host, body) as [AutomationHandle];
    expect(insert(host, [one, 0], [one, 3]).ok).toBe(true);
    expect(savedMainXml(host)).toMatch(/<w:p[ >][^]*<w:sdt>[^]*One[^]*<\/w:sdt><\/w:p>/);
  });
});
