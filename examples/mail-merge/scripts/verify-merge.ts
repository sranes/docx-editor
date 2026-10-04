/**
 * Verifies the mail merge features of `@docx-editor.dev/core` on a server, without a browser.
 *
 *   bun scripts/verify-merge.ts
 *
 * The script:
 *   1. builds the sample template and lists its merge variables;
 *   2. edits the template through the API (adds a rule and a variable control) and saves it;
 *   3. merges two records into separate copies of the saved template;
 *   4. reopens each saved result and checks its text and its XML.
 *
 * Output files go to `examples/mail-merge/out/`.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { createServerAutomationHost, type AutomationHost } from '@docx-editor.dev/core/automation';
import {
  mergeRecord,
  storiesOf,
  templateVariables,
  type MergeRecord,
} from '../src/merge/merge-engine.ts';
import { sampleTemplateBytes } from '../src/merge/sample-template.ts';
import { RECORDS } from '../src/merge/sample-records.ts';

const OUT = join(import.meta.dir, '..', 'out');
mkdirSync(OUT, { recursive: true });

let failures = 0;
function check(label: string, pass: boolean, detail = ''): void {
  if (!pass) failures += 1;
  console.log(`  ${pass ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
}

function open(bytes: Uint8Array): AutomationHost {
  const opened = createServerAutomationHost(bytes);
  if (!opened.ok) throw new Error(`cannot open document: ${opened.reason}`);
  return opened.host;
}

function save(host: AutomationHost): Uint8Array {
  const saved = host.save();
  if (!saved.ok) throw new Error(`save failed: ${saved.error.code}`);
  return saved.bytes;
}

function storyText(host: AutomationHost, index: number): string {
  const story = storiesOf(host)[index]!;
  const result = host.execute({ operations: [{ op: 'getText', target: story.body }] }).results[0];
  return result?.status === 'ok' && result.value.kind === 'text' ? result.value.text : '';
}

function part(bytes: Uint8Array, name: string): string {
  return strFromU8(unzipSync(bytes)[name]!);
}

// 1. Template variables -----------------------------------------------------------------
console.log('\n1. Read the template variables');
const template = open(sampleTemplateBytes());
const variables = templateVariables(template);
console.log(`  ${JSON.stringify(variables)}`);
check(
  'finds {{text}}, var:, and MERGEFIELD names',
  [
    'FirstName',
    'LastName',
    'CustomerId',
    'City',
    'Balance',
    'AccountNo',
    'Sender',
    'Company',
  ].every((name) => variables.values.includes(name))
);
check(
  'finds if: rule flags',
  ['isVip', 'hasBalance'].every((flag) => variables.flags.includes(flag))
);
check('finds repeating table lists', variables.lists.includes('items'));

// 2. Author the template through the API ------------------------------------------------
console.log('\n2. Edit the template through the API');
{
  const body = storiesOf(template)[0]!.body;
  const search = (text: string) => {
    const result = template.execute({ operations: [{ op: 'search', scope: { body }, text }] })
      .results[0];
    return result?.status === 'ok' && result.value.kind === 'spans' ? result.value.spans : [];
  };
  // Wrap the "Your order:" text in a rule. Content control creation accepts one paragraph.
  const [orderSpan] = search('Your order:');
  const rule = template.execute({
    operations: [
      { op: 'insertContentControl', span: orderSpan!, subtype: 'richText', tag: 'if:hasItems' },
    ],
  });
  check(
    'insertContentControl adds an if: rule around text',
    rule.ok,
    JSON.stringify(rule.results[0])
  );
  // Turn the text "Regards," into a variable control with a new tag.
  const [regardsSpan] = search('Regards,');
  const variable = template.execute({
    operations: [
      { op: 'insertContentControl', span: regardsSpan!, subtype: 'plainText', tag: 'var:Closing' },
    ],
  });
  check(
    'insertContentControl adds a var: control',
    variable.ok,
    JSON.stringify(variable.results[0])
  );
}
const authoredTemplate = save(template);
template.dispose();
writeFileSync(join(OUT, 'template.docx'), authoredTemplate);
check(
  'template saves after API edits',
  authoredTemplate.byteLength > 0,
  `${authoredTemplate.byteLength} bytes`
);

// 3 and 4. Merge each record, save, reopen, and check ------------------------------------
for (const [index, record] of RECORDS.entries()) {
  const name = `${record.values.FirstName}-${record.values.LastName}`.toLowerCase();
  console.log(`\n3. Merge record ${index + 1} (${name})`);
  const host = open(authoredTemplate);
  const report = mergeRecord(host, record);
  for (const step of report.steps) check(step.step, step.ok, step.detail);
  check('no unresolved placeholders', report.unresolved.length === 0, report.unresolved.join(', '));
  const merged = save(host);
  host.dispose();
  writeFileSync(join(OUT, `merged-${name}.docx`), merged);

  console.log(`4. Reopen merged-${name}.docx and check it`);
  const reopened = open(merged);
  const body = storyText(reopened, 0);
  const header = storiesOf(reopened)
    .slice(1)
    .map((_, i) => storyText(reopened, i + 1))
    .join('|');
  reopened.dispose();
  verifyRecord(record, body, header, part(merged, 'word/document.xml'));
}

function verifyRecord(record: MergeRecord, body: string, header: string, xml: string): void {
  const v = record.values;
  const flags = record.flags ?? {};
  check('text placeholders filled', body.includes(`Dear ${v.FirstName} ${v.LastName},`));
  check('placeholder split across runs filled', body.includes(`Customer ID: ${v.CustomerId}`));
  check('var: content control filled', body.includes(`City: ${v.City}`));
  check('var: control added by the API filled', body.includes(v.Closing!));
  check('header placeholder filled', header.includes(`${v.Company} | Confidential`), header);
  check('MERGEFIELD replaced by its value', body.includes(`Account number: ${v.AccountNo}`));
  check('MERGEFIELD code removed from XML', !xml.includes('MERGEFIELD'));
  check('if:isVip rule', body.includes('As a VIP member') === Boolean(flags.isVip));
  check('if:!isVip rule', body.includes('Join our VIP program') === !flags.isVip);
  check('if:hasBalance rule', body.includes('outstanding balance') === Boolean(flags.hasBalance));
  check(
    'if:hasItems rule added by the API',
    body.includes('Your order:') === Boolean(flags.hasItems)
  );
  check('rule wrappers removed from XML', !/w:val="if:/.test(xml));
  const payments = record.lists?.payments ?? [];
  const paymentLines = body.split('\r').filter((line) => /^\d{4}-\d{2}-\d{2}: /.test(line));
  check(
    `repeating section has one line per payment (${payments.length})`,
    paymentLines.length === payments.length &&
      payments.every((p) => paymentLines.includes(`${p.date}: ${p.amount}`)) &&
      !body.includes('{{payments.'),
    paymentLines.join(' | ')
  );
  check('repeating section wrappers removed from XML', !/repeatingSection/.test(xml));
  check('variable wrappers removed from XML', !/w:val="var:/.test(xml));
  const items = record.lists?.items ?? [];
  check(
    `table has one row per item (${items.length})`,
    items.every((item) => body.includes(`${item.name}\r${item.qty}\r${item.price}`)) &&
      !body.includes('{{items.')
  );
  check(
    'table XML has a header row plus one row per item',
    (xml.match(/<w:tr[ >]/g) ?? []).length === items.length + 1
  );
  // `{{Sender}}` sits in a bold run in the template. The value must keep that run formatting.
  const senderRun =
    xml.match(/<w:r>(?:(?!<\/w:r>).)*<\/w:r>/g)?.find((r) => r.includes(v.Sender!)) ?? '';
  check('bold formatting kept on a replaced placeholder', senderRun.includes('<w:b/>'), senderRun);
}

console.log(
  `\n${failures === 0 ? 'All checks passed.' : `${failures} check(s) failed.`} Files are in ${OUT}`
);
process.exit(failures === 0 ? 0 : 1);
