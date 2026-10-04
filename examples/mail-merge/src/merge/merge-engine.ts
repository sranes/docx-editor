/**
 * A mail merge over the core automation protocol.
 *
 * The engine needs only an `AutomationHost`, so the same code runs in two places:
 *
 * - on a server, over `createServerAutomationHost(bytes)` from `@docx-editor.dev/core/automation`;
 * - in the browser, over `createBrowserAutomationHost(editor)` from `@docx-editor.dev/core/editor`.
 *
 * Both packages use the Apache-2.0 license. The engine imports types only, plus one constant.
 *
 * Template conventions:
 *
 * - `{{Name}}` text in any story (body, headers, footers) takes `record.values.Name`.
 * - A content control tagged `var:Name` takes `record.values.Name`.
 * - A content control tagged `if:flag` keeps its content when `record.flags.flag` is true,
 *   and `if:!flag` when it is false. The engine always removes the control wrapper.
 * - A table row with `{{list.column}}` cells repeats once for each item in `record.lists.list`.
 * - A Word field `MERGEFIELD Name` is replaced with plain text `record.values.Name`.
 */

import {
  HEADER_FOOTER_VARIANTS,
  type AutomationBatchResponse,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationSpan,
  type AutomationValue,
} from '@docx-editor.dev/core/automation';

export interface MergeRecord {
  readonly values: Readonly<Record<string, string>>;
  readonly flags?: Readonly<Record<string, boolean>>;
  readonly lists?: Readonly<Record<string, readonly Readonly<Record<string, string>>[]>>;
}

export interface MergeStep {
  readonly step: string;
  readonly ok: boolean;
  readonly detail: string;
}

export interface MergeReport {
  readonly steps: readonly MergeStep[];
  /** Placeholder names that stayed in the document after the merge. */
  readonly unresolved: readonly string[];
}

export class MergeError extends Error {}

const PLACEHOLDER = /\{\{\s*([\w.]+)\s*\}\}/g;
const LIST_PLACEHOLDER = /\{\{\s*(\w+)\.(\w+)\s*\}\}/;

function execute(host: AutomationHost, operations: readonly AutomationOperation[]) {
  const response: AutomationBatchResponse = host.execute({ operations });
  return response;
}

/** Runs one query and returns its value, or throws with the host's error code. */
function query(host: AutomationHost, operation: AutomationOperation): AutomationValue {
  const result = execute(host, [operation]).results[0];
  if (result?.status !== 'ok') {
    const error =
      result?.status === 'error' ? `${result.error.code}: ${result.error.message}` : 'skipped';
    throw new MergeError(`${operation.op} failed (${error})`);
  }
  return result.value;
}

function handleOf(value: AutomationValue): AutomationHandle {
  if (value.kind !== 'handle') throw new MergeError(`expected a handle, got ${value.kind}`);
  return value.handle;
}

function handlesOf(value: AutomationValue): readonly AutomationHandle[] {
  if (value.kind !== 'handles') throw new MergeError(`expected handles, got ${value.kind}`);
  return value.handles;
}

function textOf(value: AutomationValue): string {
  if (value.kind !== 'text') throw new MergeError(`expected text, got ${value.kind}`);
  return value.text;
}

function spansOf(value: AutomationValue): readonly AutomationSpan[] {
  if (value.kind !== 'spans') throw new MergeError(`expected spans, got ${value.kind}`);
  return value.spans;
}

/** Runs a write batch. Returns a step that reports success or the first refusal. */
function write(
  host: AutomationHost,
  step: string,
  operations: readonly AutomationOperation[]
): MergeStep {
  if (operations.length === 0) return { step, ok: true, detail: 'nothing to do' };
  const response = execute(host, operations);
  if (response.ok) return { step, ok: true, detail: `${operations.length} operation(s) applied` };
  const failed = response.results.find((result) => result.status === 'error');
  const detail =
    failed?.status === 'error' ? `${failed.error.code}: ${failed.error.message}` : 'refused';
  return { step, ok: false, detail };
}

/** The main body plus every header and footer story. */
export function storiesOf(
  host: AutomationHost
): readonly { name: string; body: AutomationHandle }[] {
  const document = handleOf(query(host, { op: 'getDocument' }));
  const stories = [{ name: 'body', body: handleOf(query(host, { op: 'getBody', document })) }];
  const sections = handlesOf(query(host, { op: 'getSections', document }));
  sections.forEach((section, index) => {
    for (const kind of ['header', 'footer'] as const) {
      for (const variant of HEADER_FOOTER_VARIANTS) {
        // A variant the section does not declare is an empty story: every read answers empty.
        const result = execute(host, [{ op: 'getFurniture', kind, section, variant }]).results[0];
        if (result?.status !== 'ok' || result.value.kind !== 'handle') continue;
        stories.push({
          name: `section ${index + 1} ${variant} ${kind}`,
          body: result.value.handle,
        });
      }
    }
  });
  return stories;
}

function contentControlsByTag(host: AutomationHost, body: AutomationHandle) {
  const controls = handlesOf(query(host, { op: 'getContentControls', scope: { body } }));
  return controls.map((contentControl) => ({
    contentControl,
    tag: textOf(query(host, { op: 'getContentControlTag', contentControl })),
  }));
}

/** Step 1: apply the `if:` rules. Removes each rule wrapper, and its content when the rule fails. */
function applyRules(host: AutomationHost, body: AutomationHandle, record: MergeRecord): MergeStep {
  const operations: AutomationOperation[] = [];
  for (const { contentControl, tag } of contentControlsByTag(host, body)) {
    if (!tag.startsWith('if:')) continue;
    const negated = tag.startsWith('if:!');
    const flag = tag.slice(negated ? 4 : 3);
    const keepContent = (record.flags?.[flag] ?? false) !== negated;
    operations.push({ op: 'deleteContentControl', contentControl, keepContent });
  }
  return write(host, 'rules (if: content controls)', operations);
}

/** Step 2: fill the `var:` content controls. */
function fillVariableControls(
  host: AutomationHost,
  body: AutomationHandle,
  record: MergeRecord
): MergeStep {
  const operations: AutomationOperation[] = [];
  for (const { contentControl, tag } of contentControlsByTag(host, body)) {
    if (!tag.startsWith('var:')) continue;
    const value = record.values[tag.slice(4)];
    if (value === undefined) continue;
    operations.push({
      op: 'setContentControlValue',
      contentControl,
      value: { kind: 'text', text: value },
    });
  }
  return write(host, 'variables (var: content controls)', operations);
}

/** Step 3: repeat table rows that hold `{{list.column}}` placeholders. */
function fillRepeatingRows(
  host: AutomationHost,
  body: AutomationHandle,
  record: MergeRecord
): MergeStep[] {
  const steps: MergeStep[] = [];
  for (const table of handlesOf(query(host, { op: 'getTables', scope: { body } }))) {
    const read = query(host, { op: 'getTable', table });
    if (read.kind !== 'table') continue;
    const rowIndex = read.table.values.findIndex((row) =>
      row.some((cell) => LIST_PLACEHOLDER.test(cell))
    );
    if (rowIndex < 0) continue;
    const templateRow = read.table.values[rowIndex]!;
    const listName = LIST_PLACEHOLDER.exec(
      templateRow.find((cell) => LIST_PLACEHOLDER.test(cell))!
    )![1]!;
    const items = record.lists?.[listName] ?? [];
    const values = items.map((item) =>
      templateRow.map((cell) =>
        cell.replace(
          new RegExp(LIST_PLACEHOLDER, 'g'),
          (_, __, column: string) => item[column] ?? ''
        )
      )
    );
    // A structural table mutation must be the only command in its batch.
    if (values.length > 0) {
      steps.push(
        write(host, `table rows for "${listName}" (add ${values.length})`, [
          {
            op: 'updateTable',
            table,
            mutation: { kind: 'addRows', location: 'end', count: values.length, values },
          },
        ])
      );
    }
    steps.push(
      write(host, `table rows for "${listName}" (remove template row)`, [
        { op: 'updateTable', table, mutation: { kind: 'deleteRows', index: rowIndex, count: 1 } },
      ])
    );
  }
  return steps;
}

/**
 * Step 3b: a repeating section tagged `repeat:list` gets one item per list entry. Its
 * `{{list.column}}` placeholders are filled per item, and the section wrappers are removed.
 */
function fillRepeatingSections(
  host: AutomationHost,
  body: AutomationHandle,
  record: MergeRecord
): MergeStep[] {
  const steps: MergeStep[] = [];
  for (const { contentControl: section, tag } of contentControlsByTag(host, body)) {
    if (!tag.startsWith('repeat:')) continue;
    const list = tag.slice('repeat:'.length);
    const entries = record.lists?.[list] ?? [];
    if (entries.length === 0) {
      steps.push(
        write(host, `repeat:${list} (no entries)`, [
          { op: 'deleteContentControl', contentControl: section, keepContent: false },
        ])
      );
      continue;
    }
    // Each add is solitary in its batch. Copies of item 0 are identical, so order is free.
    for (let index = 1; index < entries.length; index += 1) {
      steps.push(
        write(host, `repeat:${list} (add item ${index + 1})`, [
          { op: 'addRepeatingSectionItem', contentControl: section, index: 0 },
        ])
      );
    }
    const items = handlesOf(
      query(host, { op: 'getContentControls', scope: { contentControl: section } })
    ).filter(
      (contentControl) =>
        textOf(query(host, { op: 'getContentControlSubtype', contentControl })) ===
        'repeatingSectionItem'
    );
    const replacements: { span: AutomationSpan; text: string }[] = [];
    items.forEach((item, position) => {
      const entry = entries[position] ?? {};
      const scope = query(host, { op: 'getContentControlRange', contentControl: item });
      if (scope.kind !== 'span') return;
      for (const [column, text] of Object.entries(entry)) {
        const found = spansOf(
          query(host, { op: 'search', scope: scope.span, text: `{{${list}.${column}}}` })
        );
        for (const span of found) replacements.push({ span, text });
      }
    });
    steps.push(
      write(
        host,
        `repeat:${list} (fill ${items.length} items)`,
        sortForReplacement(replacements).map(({ span, text }) => ({
          op: 'replaceSpan',
          span,
          text,
        }))
      ),
      write(host, `repeat:${list} (unwrap)`, [
        ...items.map(
          (contentControl): AutomationOperation => ({
            op: 'deleteContentControl',
            contentControl,
            keepContent: true,
          })
        ),
        { op: 'deleteContentControl', contentControl: section, keepContent: true },
      ])
    );
  }
  return steps;
}

/** Step 4: replace each `MERGEFIELD Name` field with plain text. */
function replaceMergeFields(
  host: AutomationHost,
  body: AutomationHandle,
  record: MergeRecord
): MergeStep {
  const fields = handlesOf(query(host, { op: 'getFields', span: { body } }));
  const operations: AutomationOperation[] = [];
  const unresolvedFields: string[] = [];
  const spans: { span: AutomationSpan; text: string }[] = [];
  for (const field of fields) {
    const read = query(host, { op: 'getField', field });
    if (read.kind !== 'field') continue;
    const name = /^\s*MERGEFIELD\s+"?([^"\s\\]+)/i.exec(read.field.code)?.[1];
    if (!name) continue;
    const value = record.values[name];
    if (value === undefined) {
      unresolvedFields.push(name);
      continue;
    }
    // Replacing the whole field span unlinks it into plain text, as Word's "merge to document".
    const range = query(host, { op: 'getFieldRange', field });
    if (range.kind === 'span') spans.push({ span: range.span, text: value });
  }
  for (const { span, text } of sortForReplacement(spans))
    operations.push({ op: 'replaceSpan', span, text });
  const step = write(host, 'MERGEFIELD fields', operations);
  return unresolvedFields.length === 0
    ? step
    : { ...step, detail: `${step.detail}; no value for ${unresolvedFields.join(', ')}` };
}

/** Later spans first in each paragraph, so earlier offsets stay valid inside one batch. */
function sortForReplacement<T extends { span: AutomationSpan }>(items: readonly T[]): T[] {
  return [...items].sort((a, b) => b.span.start.offset - a.span.start.offset);
}

/** Step 5: replace `{{Name}}` text in one story. */
function replacePlaceholders(
  host: AutomationHost,
  story: { name: string; body: AutomationHandle },
  record: MergeRecord
): MergeStep {
  const replacements: { span: AutomationSpan; text: string }[] = [];
  for (const [name, value] of Object.entries(record.values)) {
    const found = spansOf(
      query(host, {
        op: 'search',
        scope: { body: story.body },
        text: `{{${name}}}`,
        options: { matchCase: true },
      })
    );
    for (const span of found) replacements.push({ span, text: value });
  }
  const operations = sortForReplacement(replacements).map(
    ({ span, text }): AutomationOperation => ({ op: 'replaceSpan', span, text })
  );
  return write(host, `placeholders in ${story.name}`, operations);
}

function unresolvedIn(
  host: AutomationHost,
  stories: readonly { body: AutomationHandle }[]
): string[] {
  const names = new Set<string>();
  for (const { body } of stories) {
    const text = textOf(query(host, { op: 'getText', target: body }));
    for (const match of text.matchAll(PLACEHOLDER)) names.add(match[1]!);
  }
  return [...names];
}

/** Last step: remove the `var:` control wrappers and keep the filled text. */
function unwrapVariableControls(host: AutomationHost, body: AutomationHandle): MergeStep {
  const operations: AutomationOperation[] = contentControlsByTag(host, body)
    .filter(({ tag }) => tag.startsWith('var:'))
    .map(({ contentControl }) => ({
      op: 'deleteContentControl',
      contentControl,
      keepContent: true,
    }));
  return write(host, 'unwrap var: content controls', operations);
}

export interface MergeOptions {
  /** Keep the `var:` content controls in the output, for example for a preview. Default `false`. */
  readonly keepVariableControls?: boolean;
}

/** Merges one record into the document that `host` holds. */
export function mergeRecord(
  host: AutomationHost,
  record: MergeRecord,
  options: MergeOptions = {}
): MergeReport {
  const stories = storiesOf(host);
  const body = stories[0]!.body;
  const steps: MergeStep[] = [
    applyRules(host, body, record),
    fillVariableControls(host, body, record),
    ...fillRepeatingRows(host, body, record),
    ...fillRepeatingSections(host, body, record),
    replaceMergeFields(host, body, record),
    ...stories.map((story) => replacePlaceholders(host, story, record)),
  ];
  if (!options.keepVariableControls) steps.push(unwrapVariableControls(host, body));
  return { steps, unresolved: unresolvedIn(host, stories) };
}

/** Lists the merge variables a template uses: `{{Name}}` text, `var:` and `if:` tags, and MERGEFIELD names. */
export function templateVariables(host: AutomationHost): {
  values: string[];
  flags: string[];
  lists: string[];
} {
  const values = new Set<string>();
  const flags = new Set<string>();
  const lists = new Set<string>();
  const stories = storiesOf(host);
  for (const { body } of stories) {
    const text = textOf(query(host, { op: 'getText', target: body }));
    for (const match of text.matchAll(PLACEHOLDER)) {
      const [list] = match[1]!.split('.');
      (match[1]!.includes('.') ? lists : values).add(match[1]!.includes('.') ? list! : match[1]!);
    }
  }
  const body = stories[0]!.body;
  for (const { tag } of contentControlsByTag(host, body)) {
    if (tag.startsWith('var:')) values.add(tag.slice(4));
    if (tag.startsWith('if:')) flags.add(tag.replace(/^if:!?/, ''));
  }
  for (const field of handlesOf(query(host, { op: 'getFields', span: { body } }))) {
    const read = query(host, { op: 'getField', field });
    const name =
      read.kind === 'field'
        ? /^\s*MERGEFIELD\s+"?([^"\s\\]+)/i.exec(read.field.code)?.[1]
        : undefined;
    if (name) values.add(name);
  }
  return { values: [...values].sort(), flags: [...flags].sort(), lists: [...lists].sort() };
}
