// Word "Use wildcards" patterns, compiled to a regular expression for `search`.
//
// A SUBSET, and every construct outside it refuses rather than matching something else:
//
//   ?          one character             [abc]  one of the listed characters
//   *          any run of characters     [a-z]  one character in the range
//   {n} {n,} {n,m}  repeat the previous   [!a-z] one character not in the set
//   @          one or more of the previous
//   <  >       start and end of a word    \x     the character x itself
//
// Grouping `( )`, back references and `^` special codes refuse. Like Word, a wildcard search is
// always case-sensitive. The text searched comes from the document, so the pattern is bounded:
// no nested quantifiers can arise, and at most MAX_UNBOUNDED unbounded repeats keep backtracking
// polynomial in a small degree.

const MAX_PATTERN_LENGTH = 255;
const MAX_REPEAT = 255;
const MAX_UNBOUNDED = 4;
const WORD = '[\\p{L}\\p{N}_]';

export type WildcardCompile =
  | { readonly ok: true; readonly pattern: RegExp }
  | { readonly ok: false; readonly reason: string };

// Unicode-mode regular expressions accept an escape only for syntax characters, and the set
// differs inside and outside a [ ] class.
const escapeRegExp = (char: string): string => char.replace(/[\\^$.*+?()[\]{}|/]/g, '\\$&');
const escapeClass = (char: string): string => char.replace(/[\\\]\[^-]/g, '\\$&');

/** `[...]` starting at `index` (the `[`). Answers the class and the index after `]`. */
function readClass(
  pattern: string,
  index: number
): { readonly source: string; readonly next: number } | null {
  let cursor = index + 1;
  let negate = false;
  if (pattern[cursor] === '!') {
    negate = true;
    cursor += 1;
  }
  let body = '';
  while (cursor < pattern.length && pattern[cursor] !== ']') {
    let char = pattern[cursor]!;
    if (char === '\\') {
      cursor += 1;
      if (cursor >= pattern.length) return null;
      char = pattern[cursor]!;
    }
    if (pattern[cursor + 1] === '-' && cursor + 2 < pattern.length && pattern[cursor + 2] !== ']') {
      const to = pattern[cursor + 2]!;
      if (to < char) return null;
      body += `${escapeClass(char)}-${escapeClass(to)}`;
      cursor += 3;
      continue;
    }
    body += escapeClass(char);
    cursor += 1;
  }
  if (cursor >= pattern.length || body.length === 0) return null;
  return { source: `[${negate ? '^' : ''}${body}]`, next: cursor + 1 };
}

/** Compiles a Word wildcard pattern, or answers why it is outside the supported subset. */
export function compileWordWildcards(pattern: string): WildcardCompile {
  if (pattern.length === 0 || pattern.length > MAX_PATTERN_LENGTH)
    return { ok: false, reason: `a wildcard pattern has 1 to ${MAX_PATTERN_LENGTH} characters` };
  let source = '';
  // Whether the last piece is an atom a repeat may apply to: a character, `?` or a class.
  let repeatable = false;
  let unbounded = 0;
  let index = 0;
  while (index < pattern.length) {
    const char = pattern[index]!;
    if (char === '\\') {
      const next = pattern[index + 1];
      if (next === undefined) return { ok: false, reason: 'a wildcard pattern ends with \\' };
      source += escapeRegExp(next);
      repeatable = true;
      index += 2;
      continue;
    }
    if (char === '?') {
      source += '[^]';
      repeatable = true;
    } else if (char === '*') {
      source += '[^]*?';
      unbounded += 1;
      repeatable = false;
    } else if (char === '[') {
      const read = readClass(pattern, index);
      if (!read) return { ok: false, reason: 'a [ ] set in the wildcard pattern is not valid' };
      source += read.source;
      repeatable = true;
      index = read.next;
      continue;
    } else if (char === '{') {
      const close = pattern.indexOf('}', index);
      const counts =
        close < 0 ? null : /^(\d{1,3})(,(\d{0,3}))?$/.exec(pattern.slice(index + 1, close));
      if (!repeatable || !counts)
        return { ok: false, reason: 'a { } repeat must follow a character' };
      const min = Number(counts[1]);
      const max = counts[2] === undefined ? min : counts[3] === '' ? null : Number(counts[3]);
      if (min > MAX_REPEAT || (max !== null && (max > MAX_REPEAT || max < min)))
        return { ok: false, reason: `a { } repeat counts from 0 to ${MAX_REPEAT}` };
      if (max === null) unbounded += 1;
      source += `{${min},${max ?? ''}}`.replace(`{${min},${min}}`, `{${min}}`);
      repeatable = false;
      index = close + 1;
      continue;
    } else if (char === '@') {
      if (!repeatable) return { ok: false, reason: '@ must follow a character' };
      source += '+';
      unbounded += 1;
      repeatable = false;
    } else if (char === '<') {
      source += `(?<!${WORD})(?=${WORD})`;
      repeatable = false;
    } else if (char === '>') {
      source += `(?<=${WORD})(?!${WORD})`;
      repeatable = false;
    } else if (char === '(' || char === ')' || char === '^') {
      return { ok: false, reason: `the wildcard ${char} is not supported` };
    } else {
      source += escapeRegExp(char);
      repeatable = true;
    }
    index += 1;
  }
  if (unbounded > MAX_UNBOUNDED)
    return {
      ok: false,
      reason: `a wildcard pattern has at most ${MAX_UNBOUNDED} unbounded repeats`,
    };
  try {
    return { ok: true, pattern: new RegExp(source, 'gu') };
  } catch {
    return { ok: false, reason: 'that wildcard pattern is not valid' };
  }
}

/** `matchWildcards` with the other search options it can be combined with. */
export function compileWildcardSearch(
  pattern: string,
  options: { readonly matchWholeWord?: boolean }
): WildcardCompile {
  if (options.matchWholeWord === true)
    return {
      ok: false,
      reason: 'matchWholeWord cannot be combined with matchWildcards; use < and >',
    };
  return compileWordWildcards(pattern);
}
