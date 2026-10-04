// `applyVariables`: `{{name}}` placeholders in every story, replaced as one undo step.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { checkVariables } from '../variables-check.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const run = (text: string, bold = false) =>
  `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

function docx(): Uint8Array {
  const body =
    `<w:p>${run('Dear {{First')}${run('Name}} {{LastName}},')}</w:p>` +
    `<w:p>${run('{{Sender}}', true)}</w:p>` +
    `<w:p>${run('Keep {{Unknown}}.')}</w:p>` +
    `<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>`;
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/>` +
        `<Override PartName="/word/header1.xml" ContentType="${MAIN}.header+xml"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rIdH" Type="${R}/header" Target="header1.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`
    ),
    'word/header1.xml': strToU8(
      `<w:hdr xmlns:w="${W}"><w:p>${run('{{Company}} memo')}</w:p></w:hdr>`
    ),
  });
}

function mount(): DocxEditorInstance {
  const container = document.createElement('div');
  const editor = createDocxEditor({ container, document: docx() });
  if (!editor.surface) throw new Error('surface failed to mount');
  return editor;
}

async function parts(editor: DocxEditorInstance): Promise<{ body: string; header: string }> {
  const files = unzipSync(new Uint8Array(await editor.save()));
  const text = (xml: string) =>
    [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');
  return {
    body: text(strFromU8(files['word/document.xml']!)),
    header: text(strFromU8(files['word/header1.xml']!)),
  };
}

const VALUES = {
  FirstName: 'Ada',
  LastName: 'Lovelace',
  Sender: 'The Team',
  Company: 'Engines Ltd',
};

describe('applyVariables', () => {
  test('replaces placeholders in the body and the header, and leaves unknown names', async () => {
    const editor = mount();
    const result = editor.exec({ type: 'applyVariables', values: VALUES });
    expect(result).toMatchObject({ ok: true, changed: true });
    const { body, header } = await parts(editor);
    expect(body).toBe('Dear Ada Lovelace,The TeamKeep {{Unknown}}.');
    expect(header).toBe('Engines Ltd memo');
  });

  test('keeps the run formatting of the placeholder', async () => {
    const editor = mount();
    editor.exec({ type: 'applyVariables', values: VALUES });
    const xml = strFromU8(unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!);
    expect(xml).toMatch(/<w:b\/><\/w:rPr><w:t[^>]*>The Team<\/w:t>/);
  });

  test('one undo restores every story', async () => {
    const editor = mount();
    editor.exec({ type: 'applyVariables', values: VALUES });
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const { body, header } = await parts(editor);
    expect(body).toBe('Dear {{FirstName}} {{LastName}},{{Sender}}Keep {{Unknown}}.');
    expect(header).toBe('{{Company}} memo');
  });

  test('refuses invalid values before it writes', () => {
    const editor = mount();
    for (const values of [
      { 'a}b': 'x' },
      { Name: 'two\nlines' },
      { Name: 5 as unknown as string },
    ]) {
      expect(editor.can({ type: 'applyVariables', values })).toMatchObject({
        ok: false,
        code: 'invalidArgs',
      });
      expect(editor.exec({ type: 'applyVariables', values })).toMatchObject({ ok: false });
    }
  });

  test('reads own keys only', () => {
    const parsed = Object.assign(Object.create({ Sender: 'x' }), { A: 'b' }) as Record<
      string,
      string
    >;
    expect(checkVariables(parsed)).toEqual({ ok: true, entries: [['A', 'b']] });
  });
});
