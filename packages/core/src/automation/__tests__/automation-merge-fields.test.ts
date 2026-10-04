// MERGEFIELD results and field spans.
//
// A mail merge needs two things from a field: a value written into its cached result, and the
// span it occupies, so a caller can replace the whole field with plain text. Both are checked
// against the saved package, because a result that reads back in the session but does not
// survive the serializer has not been merged.

import { describe, expect, test } from 'bun:test';
import { parseMergeField, mergeFieldResult } from '../merge-field.ts';
import {
  docx,
  errorAt,
  handlesAt,
  open,
  roots,
  savedMainXml,
  spanAt,
  storyText,
} from './support/protocol.ts';
import type { AutomationHandle, AutomationHost } from '../protocol.ts';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A complex field whose cached result is one run. */
const complexField = (instruction: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r><w:rPr><w:b/></w:rPr><w:t>${result}</w:t></w:r>` +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

const simpleField = (instruction: string, result: string) =>
  `<w:fldSimple w:instr=" ${instruction.replaceAll('"', '&quot;')} "><w:r><w:t>${result}</w:t></w:r></w:fldSimple>`;

function hostWith(...fields: string[]): AutomationHost {
  return open(docx(`<w:p>${run('Dear ')}${fields.join(run(' '))}${run('.')}</w:p>`));
}

function fieldsOf(host: AutomationHost): readonly AutomationHandle[] {
  const { body } = roots(host);
  return handlesAt(host.execute({ operations: [{ op: 'getFields', span: { body } }] }), 0);
}

describe('updateFieldResult on MERGEFIELD', () => {
  test('writes the value into a complex field result and keeps the field', () => {
    const host = hostWith(complexField('MERGEFIELD FirstName', '«FirstName»'));
    const [field] = fieldsOf(host) as [AutomationHandle];
    const response = host.execute({
      operations: [{ op: 'updateFieldResult', field, values: { FirstName: 'Ada' } }],
    });
    expect(response.ok).toBe(true);
    expect(storyText(host, roots(host).body)).toBe('Dear Ada.');
    const xml = savedMainXml(host);
    expect(xml).toContain('MERGEFIELD FirstName');
    expect(xml).not.toContain('«FirstName»');
  });

  test('writes a simple field and applies \\b, \\f and a case switch', () => {
    const host = hostWith(simpleField('MERGEFIELD City \\b "in " \\f "!" \\* Upper', '«City»'));
    const [field] = fieldsOf(host) as [AutomationHandle];
    const response = host.execute({
      operations: [{ op: 'updateFieldResult', field, values: { City: 'London' } }],
    });
    expect(response.ok).toBe(true);
    expect(storyText(host, roots(host).body)).toBe('Dear in LONDON!.');
  });

  test('refuses a missing value, an unsupported switch, and values on a non-merge field', () => {
    const host = hostWith(
      complexField('MERGEFIELD FirstName', '«FirstName»'),
      complexField('MERGEFIELD Due \\@ "d MMMM"', '«Due»'),
      simpleField('AUTHOR', 'someone')
    );
    const [name, due, author] = fieldsOf(host) as [
      AutomationHandle,
      AutomationHandle,
      AutomationHandle,
    ];
    const before = savedMainXml(host);
    expect(
      errorAt(
        host.execute({ operations: [{ op: 'updateFieldResult', field: name, values: {} }] }),
        0
      )
    ).toBe('unsupported-content');
    // An inherited member is not a value.
    expect(
      errorAt(
        host.execute({
          operations: [
            {
              op: 'updateFieldResult',
              field: name,
              values: JSON.parse('{"__proto__":{"FirstName":"x"}}'),
            },
          ],
        }),
        0
      )
    ).toBe('unsupported-content');
    expect(
      errorAt(
        host.execute({
          operations: [{ op: 'updateFieldResult', field: due, values: { Due: 'x' } }],
        }),
        0
      )
    ).toBe('unsupported-capability');
    expect(
      errorAt(
        host.execute({
          operations: [{ op: 'updateFieldResult', field: author, values: { a: 'b' } }],
        }),
        0
      )
    ).toBe('unsupported-capability');
    expect(savedMainXml(host)).toBe(before);
  });
});

describe('getFieldRange', () => {
  test('answers the span a replaceSpan uses to unlink the field into plain text', () => {
    const host = hostWith(complexField('MERGEFIELD FirstName', 'cached text that differs'));
    const [field] = fieldsOf(host) as [AutomationHandle];
    const span = spanAt(host.execute({ operations: [{ op: 'getFieldRange', field }] }), 0);
    expect(span.start.offset).toBe(5);
    expect(span.end.offset).toBeGreaterThan(5);
    const response = host.execute({ operations: [{ op: 'replaceSpan', span, text: 'Ada' }] });
    expect(response.ok).toBe(true);
    expect(storyText(host, roots(host).body)).toBe('Dear Ada.');
    expect(savedMainXml(host)).not.toContain('MERGEFIELD');
  });

  test('refuses a handle that is not a field', () => {
    const host = hostWith(complexField('MERGEFIELD FirstName', '«FirstName»'));
    const { body } = roots(host);
    expect(errorAt(host.execute({ operations: [{ op: 'getFieldRange', field: body }] }), 0)).toBe(
      'invalid-handle'
    );
  });
});

describe('parseMergeField', () => {
  test('reads quoted names and escaped quotes', () => {
    const parsed = parseMergeField('MERGEFIELD "First Name" \\b "say \\"hi\\" "');
    expect(parsed).toEqual({
      ok: true,
      field: { name: 'First Name', before: 'say "hi" ', after: '', textCase: null },
    });
  });

  test('refuses unclosed quotes, oversized instructions, and other fields', () => {
    expect(parseMergeField('MERGEFIELD "Open').ok).toBe(false);
    expect(parseMergeField(`MERGEFIELD ${'x'.repeat(5000)}`).ok).toBe(false);
    expect(parseMergeField('DATE \\@ "d"').ok).toBe(false);
  });

  test('an empty value shows nothing, without \\b or \\f text', () => {
    const parsed = parseMergeField('MERGEFIELD A \\b "x" \\f "y" \\* Caps');
    if (!parsed.ok) throw new Error(parsed.reason);
    expect(mergeFieldResult(parsed.field, '')).toBe('');
    expect(mergeFieldResult(parsed.field, 'ada lovelace')).toBe('xAda Lovelacey');
  });
});
