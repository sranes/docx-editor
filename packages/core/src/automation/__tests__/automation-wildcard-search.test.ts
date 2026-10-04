// `search` with `matchWildcards`: Word's wildcard subset, mapped back to model offsets.

import { describe, expect, test } from 'bun:test';
import { compileWordWildcards } from '../wildcards.ts';
import { docx, errorAt, open, p, roots, spansAt } from './support/protocol.ts';
import type { AutomationHost, AutomationSpan } from '../protocol.ts';

const run = (text: string, bold = false) =>
  `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

function texts(host: AutomationHost, pattern: string, options = {}): string[] {
  const { body } = roots(host);
  const spans = spansAt(
    host.execute({
      operations: [
        {
          op: 'search',
          scope: { body },
          text: pattern,
          options: { matchWildcards: true, ...options },
        },
      ],
    }),
    0
  );
  return spans.map((span: AutomationSpan) => {
    const result = host.execute({ operations: [{ op: 'getSpanText', span }] }).results[0];
    return result?.status === 'ok' && result.value.kind === 'text' ? result.value.text : '';
  });
}

describe('matchWildcards search', () => {
  test('finds every {{placeholder}}, also one split across runs', () => {
    const host = open(
      docx(
        `<w:p>${run('Dear {{First')}${run('Name}} {{LastName}}', true)}</w:p>${p('Total {{Amount}}.')}`
      )
    );
    expect(texts(host, '\\{\\{*\\}\\}')).toEqual(['{{FirstName}}', '{{LastName}}', '{{Amount}}']);
  });

  test('supports ?, sets, ranges, repeats, @ and word edges', () => {
    const host = open(docx(p('cat cot cut c4t coat caaat') + p('Item A-12, item B-7')));
    expect(texts(host, 'c?t')).toEqual(['cat', 'cot', 'cut', 'c4t']);
    expect(texts(host, 'c[ao]t')).toEqual(['cat', 'cot']);
    expect(texts(host, 'c[!a-z]t')).toEqual(['c4t']);
    expect(texts(host, 'ca{2,}t')).toEqual(['caaat']);
    expect(texts(host, '[A-Z]-[0-9]@')).toEqual(['A-12', 'B-7']);
    expect(texts(host, '<item')).toEqual(['item']);
  });

  test('is case-sensitive, as in Word', () => {
    const host = open(docx(p('Alpha alpha')));
    expect(texts(host, 'a?pha')).toEqual(['alpha']);
  });

  test('refuses what the subset does not cover, and whole-word together with wildcards', () => {
    const host = open(docx(p('alpha')));
    const { body } = roots(host);
    for (const [pattern, options] of [
      ['(alpha)', {}],
      ['^p', {}],
      ['*a*b*c*d*', {}],
      ['alpha', { matchWholeWord: true }],
    ] as const) {
      const response = host.execute({
        operations: [
          {
            op: 'search',
            scope: { body },
            text: pattern,
            options: { matchWildcards: true, ...options },
          },
        ],
      });
      expect(errorAt(response, 0)).toBe('unsupported-capability');
    }
  });
});

describe('compileWordWildcards', () => {
  test('escapes regular expression syntax that is literal in Word', () => {
    const compiled = compileWordWildcards('a.b+c$-');
    expect(compiled.ok && compiled.pattern.test('a.b+c$-')).toBe(true);
    expect(compiled.ok && new RegExp(compiled.pattern.source, 'u').test('aXb+c$-')).toBe(false);
  });

  test('refuses malformed sets and repeats', () => {
    for (const pattern of ['[abc', '[]', 'a{3', '{2}', 'a{5,2}', '@', '\\', '[z-a]'])
      expect(compileWordWildcards(pattern).ok).toBe(false);
  });
});
