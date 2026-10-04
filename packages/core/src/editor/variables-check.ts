// Validation for the `applyVariables` command, shared by `classifyCommand` and exec.

/** Most names one command may carry. A template with more is a data error, not a merge. */
const MAX_VARIABLES = 1000;
const MAX_NAME_LENGTH = 128;
const MAX_VALUE_LENGTH = 32_768;
// A name a `{{...}}` placeholder can spell: no braces, no control characters or line breaks.
const SEPARATORS = String.fromCharCode(0x2028, 0x2029);
const NAME_REFUSED = new RegExp('[{}\u0000-\u001f\u007f-\u009f' + SEPARATORS + ']');
const LINE_BREAK = new RegExp('[\r\n' + SEPARATORS + ']');

export type VariablesCheck =
  | { readonly ok: true; readonly entries: readonly (readonly [string, string])[] }
  | { readonly ok: false; readonly reason: string };

/** Validates caller-supplied `values`. Reads own keys only, so no inherited member is a name. */
export function checkVariables(values: unknown): VariablesCheck {
  if (typeof values !== 'object' || values === null || Array.isArray(values))
    return { ok: false, reason: 'applyVariables requires a values object' };
  const names = Object.keys(values);
  if (names.length > MAX_VARIABLES)
    return { ok: false, reason: `applyVariables accepts at most ${MAX_VARIABLES} values` };
  const entries: [string, string][] = [];
  for (const name of names) {
    const value: unknown = (values as Record<string, unknown>)[name];
    if (name.length === 0 || name.length > MAX_NAME_LENGTH || NAME_REFUSED.test(name))
      return { ok: false, reason: `'${name.slice(0, 40)}' is not a variable name` };
    if (typeof value !== 'string' || value.length > MAX_VALUE_LENGTH)
      return {
        ok: false,
        reason: `the value of '${name}' must be a string of at most ${MAX_VALUE_LENGTH} characters`,
      };
    if (LINE_BREAK.test(value))
      return { ok: false, reason: `the value of '${name}' must not contain a line break` };
    entries.push([name, value]);
  }
  return { ok: true, entries };
}
