// MERGEFIELD instruction parsing and evaluation for `updateFieldResult`.
//
// The instruction comes from the document, so it is untrusted: the tokenizer is a bounded
// linear scan, never a regular expression, and every switch it does not implement refuses
// rather than being ignored. A merge result that silently dropped a date picture would be a
// wrong letter, not an approximate one.

/** Longest instruction this parser reads. Word's own limit is far lower. */
const MAX_INSTRUCTION_LENGTH = 2048;

export type MergeFieldTextCase = 'Upper' | 'Lower' | 'FirstCap' | 'Caps';

export interface MergeFieldInstruction {
  readonly name: string;
  /** `\b`: text placed before a non-empty result. */
  readonly before: string;
  /** `\f`: text placed after a non-empty result. */
  readonly after: string;
  /** `\*` case switch. `MERGEFORMAT` and `CHARFORMAT` keep result formatting and set none. */
  readonly textCase: MergeFieldTextCase | null;
}

export type MergeFieldParse =
  | { readonly ok: true; readonly field: MergeFieldInstruction }
  | { readonly ok: false; readonly reason: string };

/** Splits an instruction into words and quoted strings. Answers null on an unclosed quote. */
function tokenize(instruction: string): string[] | null {
  const tokens: string[] = [];
  let index = 0;
  while (index < instruction.length) {
    const char = instruction[index]!;
    if (char === ' ' || char === '\t' || char === '\r' || char === '\n') {
      index += 1;
      continue;
    }
    if (char === '"') {
      let value = '';
      index += 1;
      let closed = false;
      while (index < instruction.length) {
        const next = instruction[index]!;
        if (next === '\\' && index + 1 < instruction.length) {
          value += instruction[index + 1];
          index += 2;
          continue;
        }
        if (next === '"') {
          closed = true;
          index += 1;
          break;
        }
        value += next;
        index += 1;
      }
      if (!closed) return null;
      tokens.push(value);
      continue;
    }
    let word = '';
    while (index < instruction.length && !/\s|"/.test(instruction[index]!)) {
      word += instruction[index];
      index += 1;
    }
    tokens.push(word);
  }
  return tokens;
}

const CASE_SWITCHES: Readonly<Record<string, MergeFieldTextCase | null>> = {
  upper: 'Upper',
  lower: 'Lower',
  firstcap: 'FirstCap',
  caps: 'Caps',
  mergeformat: null,
  charformat: null,
};

/** Reads a `MERGEFIELD` instruction. Any other field, or any unsupported switch, refuses. */
export function parseMergeField(instruction: string): MergeFieldParse {
  if (instruction.length > MAX_INSTRUCTION_LENGTH)
    return { ok: false, reason: 'field instruction is too long' };
  const tokens = tokenize(instruction);
  if (!tokens) return { ok: false, reason: 'field instruction has an unclosed quote' };
  if (tokens[0]?.toUpperCase() !== 'MERGEFIELD') return { ok: false, reason: 'not a MERGEFIELD' };
  const name = tokens[1];
  if (!name || name.startsWith('\\')) return { ok: false, reason: 'MERGEFIELD names no field' };
  let before = '';
  let after = '';
  let textCase: MergeFieldTextCase | null = null;
  for (let index = 2; index < tokens.length; index += 1) {
    const token = tokens[index]!;
    const flag = token.toLowerCase();
    if (flag === '\\b' || flag === '\\f') {
      const text = tokens[index + 1];
      if (text === undefined) return { ok: false, reason: `${token} needs text` };
      if (flag === '\\b') before = text;
      else after = text;
      index += 1;
      continue;
    }
    if (flag === '\\*') {
      const format = tokens[index + 1]?.toLowerCase();
      if (format === undefined || !(format in CASE_SWITCHES))
        return {
          ok: false,
          reason: `format switch \\* ${tokens[index + 1] ?? ''} is not supported`,
        };
      textCase = CASE_SWITCHES[format] ?? textCase;
      index += 1;
      continue;
    }
    // `\@` dates, `\#` numbers, `\m` mapped fields and `\v` vertical text are not evaluated.
    return { ok: false, reason: `MERGEFIELD switch ${token} is not supported` };
  }
  return { ok: true, field: { name, before, after, textCase } };
}

function applyCase(text: string, textCase: MergeFieldTextCase | null): string {
  switch (textCase) {
    case 'Upper':
      return text.toUpperCase();
    case 'Lower':
      return text.toLowerCase();
    case 'FirstCap':
      return text.charAt(0).toUpperCase() + text.slice(1);
    case 'Caps':
      return text.replace(
        /(^|\s)(\S)/g,
        (_, space: string, letter: string) => space + letter.toUpperCase()
      );
    default:
      return text;
  }
}

/** The result Word shows for this field and value. An empty value shows nothing, no `\b`/`\f`. */
export function mergeFieldResult(field: MergeFieldInstruction, value: string): string {
  if (value === '') return '';
  return field.before + applyCase(value, field.textCase) + field.after;
}
