// The mail merge side panel. It renders in a portal, so its place in the editor tree only
// gives it access to the editor instance through `useDocxEditor()`.

import { useState, type RefObject } from 'react';
import { createPortal } from 'react-dom';
import { useDocxEditor, type DocxEditorRef } from '@docx-editor.dev/react';
import { createBrowserAutomationHost } from '@docx-editor.dev/core/editor';
import { createServerAutomationHost, type AutomationHost } from '@docx-editor.dev/core/automation';
import { mergeRecord, templateVariables, type MergeReport } from './merge/merge-engine';
import { RECORDS } from './merge/sample-records';

const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

interface MergePanelProps {
  readonly editorRef: RefObject<DocxEditorRef | null>;
  readonly onLoad: (bytes: ArrayBuffer, title: string) => void;
}

function download(bytes: Uint8Array, name: string): void {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: DOCX }));
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

function openServerHost(bytes: ArrayBuffer): AutomationHost {
  const opened = createServerAutomationHost(new Uint8Array(bytes));
  if (!opened.ok) throw new Error(`cannot open document: ${opened.reason}`);
  return opened.host;
}

function recordName(index: number): string {
  const values = RECORDS[index]!.values;
  return `${values.FirstName} ${values.LastName}`;
}

export function MergePanel({ editorRef, onLoad }: MergePanelProps) {
  const editor = useDocxEditor();
  const [name, setName] = useState('FirstName');
  const [flag, setFlag] = useState('isVip');
  const [recordIndex, setRecordIndex] = useState(0);
  const [template, setTemplate] = useState<ArrayBuffer | null>(null);
  const [log, setLog] = useState('Ready.');

  const report = (label: string, value: unknown) =>
    setLog(`${label}\n${typeof value === 'string' ? value : JSON.stringify(value, null, 2)}`);

  const guard = (label: string, action: () => unknown | Promise<unknown>) => async () => {
    try {
      const value = await action();
      if (value !== undefined) report(label, value);
    } catch (error) {
      report(`${label} failed`, error instanceof Error ? error.message : String(error));
    }
  };

  /** Runs an editor command at the current selection and answers its result. */
  const atSelection = (
    command: (target: NonNullable<ReturnType<DocxEditorRef['snapshot']>['selection']>) => unknown
  ) => {
    const ref = editorRef.current;
    const selection = ref?.snapshot().selection;
    if (!ref || !selection) return 'Click in the document first.';
    return command(selection);
  };

  const listVariables = guard('Template variables', () => {
    if (!editor) return 'The editor is not ready.';
    const host = createBrowserAutomationHost(editor);
    try {
      return templateVariables(host);
    } finally {
      host.dispose();
    }
  });

  // `insertText` acts at the selection. The editor refuses a command that sets `target`.
  const insertPlaceholder = guard(`Insert {{${name}}}`, () =>
    atSelection(() => editorRef.current!.exec({ type: 'insertText', text: `{{${name}}}` }))
  );

  const makeVariableControl = guard(`Make var:${name} control`, () =>
    atSelection((target) =>
      editorRef.current!.exec({
        type: 'insertContentControl',
        subtype: 'plainText',
        tag: `var:${name}`,
        title: name,
        target,
      })
    )
  );

  const makeRule = (negated: boolean) =>
    guard(`Wrap selection in if:${negated ? '!' : ''}${flag}`, () =>
      atSelection((target) =>
        editorRef.current!.exec({
          type: 'insertContentControl',
          subtype: 'richText',
          tag: `if:${negated ? '!' : ''}${flag}`,
          title: `Rule: ${negated ? 'not ' : ''}${flag}`,
          target,
        })
      )
    );

  const tryApplyVariables = guard('applyVariables command', () =>
    atSelection(() =>
      editorRef.current!.exec({ type: 'applyVariables', values: RECORDS[recordIndex]!.values })
    )
  );

  /** Saves the template, merges a copy on a headless host, and shows the copy in the editor. */
  const preview = guard('Preview merge', async () => {
    const bytes = template ?? (await editorRef.current?.save());
    if (!bytes) return 'The editor has no document.';
    setTemplate(bytes);
    const host = openServerHost(bytes);
    let result: MergeReport;
    try {
      result = mergeRecord(host, RECORDS[recordIndex]!);
      const saved = host.save();
      if (!saved.ok) throw new Error(saved.error.message);
      onLoad(saved.bytes.slice().buffer as ArrayBuffer, `Preview: ${recordName(recordIndex)}`);
    } finally {
      host.dispose();
    }
    return result;
  });

  const backToTemplate = guard('Back to template', () => {
    if (!template) return 'No preview is open.';
    onLoad(template, 'Template');
    setTemplate(null);
    return 'Template restored.';
  });

  /** Merges into the open editor through the browser host. Undo reverts it. */
  const mergeLive = guard('Merge in the live editor', () => {
    if (!editor) return 'The editor is not ready.';
    // Batches into the same story join one undo step; the header is a separate step.
    const group = editor.beginHistoryGroup();
    const host = createBrowserAutomationHost(editor, { historyGroup: group });
    try {
      return mergeRecord(host, RECORDS[recordIndex]!, { keepVariableControls: true });
    } finally {
      group.end();
      host.dispose();
    }
  });

  const downloadAll = guard('Download all merged documents', async () => {
    const bytes = template ?? (await editorRef.current?.save());
    if (!bytes) return 'The editor has no document.';
    const summary = RECORDS.map((record, index) => {
      const host = openServerHost(bytes);
      try {
        const result = mergeRecord(host, record);
        const saved = host.save();
        if (!saved.ok) throw new Error(saved.error.message);
        download(saved.bytes, `merged-${index + 1}.docx`);
        return {
          record: recordName(index),
          failedSteps: result.steps.filter((s) => !s.ok),
          unresolved: result.unresolved,
        };
      } finally {
        host.dispose();
      }
    });
    return summary;
  });

  return createPortal(
    <aside className="mm-panel" data-testid="merge-panel">
      <h2>Mail merge</h2>
      <section>
        <h3>Template</h3>
        <button type="button" onClick={listVariables} data-testid="list-variables">
          List variables
        </button>
        <label>
          Variable name
          <input
            value={name}
            onChange={(e) => setName(e.target.value)}
            data-testid="variable-name"
          />
        </label>
        <div className="mm-row">
          <button type="button" onClick={insertPlaceholder} data-testid="insert-placeholder">
            Insert {'{{'}…{'}}'}
          </button>
          <button type="button" onClick={makeVariableControl} data-testid="make-variable">
            Selection → var:
          </button>
        </div>
        <label>
          Rule flag
          <input value={flag} onChange={(e) => setFlag(e.target.value)} data-testid="rule-flag" />
        </label>
        <div className="mm-row">
          <button type="button" onClick={makeRule(false)} data-testid="make-rule">
            Selection → if:
          </button>
          <button type="button" onClick={makeRule(true)} data-testid="make-rule-negated">
            Selection → if:!
          </button>
        </div>
      </section>
      <section>
        <h3>Merge</h3>
        <label>
          Record
          <select
            value={recordIndex}
            onChange={(e) => setRecordIndex(Number(e.target.value))}
            data-testid="record"
          >
            {RECORDS.map((_, index) => (
              <option key={index} value={index}>
                {recordName(index)}
              </option>
            ))}
          </select>
        </label>
        <div className="mm-row">
          <button type="button" onClick={preview} data-testid="preview">
            Preview
          </button>
          <button type="button" onClick={backToTemplate} data-testid="back">
            Back to template
          </button>
        </div>
        <div className="mm-row">
          <button type="button" onClick={mergeLive} data-testid="merge-live">
            Merge in editor
          </button>
          <button type="button" onClick={downloadAll} data-testid="download-all">
            Download all
          </button>
        </div>
        <button type="button" onClick={tryApplyVariables} data-testid="apply-variables">
          Try applyVariables
        </button>
      </section>
      <pre className="mm-log" data-testid="merge-log">
        {log}
      </pre>
    </aside>,
    document.body
  );
}
