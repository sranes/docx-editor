// A browser automation host created with `historyGroup`: several batches, one undo step.
//
// A batch writes into one story, so a merge across paragraphs or stories is several batches.
// Without a group each one is its own undo step and the user presses undo once per batch.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { createBrowserAutomationHost } from '../automation-host.ts';
import type { AutomationHandle, AutomationHost } from '../../automation/index.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

function mount(): DocxEditorInstance {
  const container = document.createElement('div');
  const editor = createDocxEditor({ container, document: docx(`${p('alpha')}${p('beta')}`) });
  if (!editor.surface) throw new Error('surface failed to mount');
  return editor;
}

function value(
  host: AutomationHost,
  operation: Parameters<AutomationHost['execute']>[0]['operations'][number]
) {
  const result = host.execute({ operations: [operation] }).results[0];
  if (result?.status !== 'ok') throw new Error(`${operation.op} failed`);
  return result.value;
}

function paragraphs(host: AutomationHost): readonly AutomationHandle[] {
  const documentValue = value(host, { op: 'getDocument' });
  if (documentValue.kind !== 'handle') throw new Error('no document');
  const bodyValue = value(host, { op: 'getBody', document: documentValue.handle });
  if (bodyValue.kind !== 'handle') throw new Error('no body');
  const listed = value(host, { op: 'getParagraphs', body: bodyValue.handle });
  if (listed.kind !== 'handles') throw new Error('no paragraphs');
  return listed.handles;
}

function texts(host: AutomationHost): string[] {
  return paragraphs(host).map((target) => {
    const read = value(host, { op: 'getText', target });
    return read.kind === 'text' ? read.text : '';
  });
}

/** Two batches, one per paragraph: the shape of a merge that touches more than one place. */
function twoBatches(host: AutomationHost): void {
  const [first, second] = paragraphs(host) as [AutomationHandle, AutomationHandle];
  expect(
    host.execute({
      operations: [{ op: 'insertText', at: { paragraph: first, offset: 0 }, text: 'A-' }],
    }).ok
  ).toBe(true);
  expect(
    host.execute({
      operations: [{ op: 'insertText', at: { paragraph: second, offset: 0 }, text: 'B-' }],
    }).ok
  ).toBe(true);
}

describe('createBrowserAutomationHost with historyGroup', () => {
  test('batches run while the group is open undo as one step', () => {
    const editor = mount();
    const group = editor.beginHistoryGroup();
    const host = createBrowserAutomationHost(editor, { historyGroup: group });
    twoBatches(host);
    group.end();
    expect(texts(createBrowserAutomationHost(editor))).toEqual(['A-alpha', 'B-beta']);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(createBrowserAutomationHost(editor))).toEqual(['alpha', 'beta']);
    host.dispose();
  });

  test('without a group each batch is its own undo step', () => {
    const editor = mount();
    const host = createBrowserAutomationHost(editor);
    twoBatches(host);
    editor.exec({ type: 'undo' });
    expect(texts(host)).toEqual(['A-alpha', 'beta']);
    host.dispose();
  });

  test('after the group ends, the host refuses every batch and writes nothing', () => {
    const editor = mount();
    const group = editor.beginHistoryGroup();
    const host = createBrowserAutomationHost(editor, { historyGroup: group });
    group.end();
    const response = host.execute({ operations: [{ op: 'getDocument' }] });
    expect(response.ok).toBe(false);
    expect(response.changed).toBe(false);
    const result = response.results[0];
    expect(result?.status === 'error' && result.error.code).toBe('unsupported-capability');
    host.dispose();
  });
});

describe('getSelection', () => {
  test('answers the body selection as a span that getSpanText can read', () => {
    const editor = mount();
    const host = createBrowserAutomationHost(editor);
    const [first, second] = paragraphs(host) as [AutomationHandle, AutomationHandle];
    // Select from offset 1 of the first paragraph to offset 2 of the second.
    expect(
      host.execute({
        operations: [
          {
            op: 'selectSpan',
            span: { start: { paragraph: first, offset: 1 }, end: { paragraph: second, offset: 2 } },
            mode: 'select',
          },
        ],
      }).ok
    ).toBe(true);
    const span = value(host, { op: 'getSelection' });
    expect(span.kind).toBe('span');
    if (span.kind !== 'span') return;
    expect(span.span.start.offset).toBe(1);
    expect(span.span.end.offset).toBe(2);
    const text = value(host, { op: 'getSpanText', span: span.span });
    expect(text).toEqual({ kind: 'text', text: 'lpha\rbe' });
    host.dispose();
  });

  test('a headless host refuses: it has no reader', async () => {
    const { createServerAutomationHost } = await import('../../automation/server-host.ts');
    const opened = createServerAutomationHost(docx(p('alpha')));
    if (!opened.ok) throw new Error('did not open');
    const result = opened.host.execute({ operations: [{ op: 'getSelection' }] }).results[0];
    expect(result?.status === 'error' && result.error.code).toBe('unsupported-capability');
  });
});
