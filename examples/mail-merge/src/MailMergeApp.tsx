// A mail merge client over @docx-editor.dev/core and @docx-editor.dev/react only.
//
// The packaged <DocxEditor> edits the template. The side panel uses the core automation
// protocol for the merge, so the same merge engine also runs on a server.

import { useCallback, useRef, useState } from 'react';
import { DocxEditor, type DocxEditorRef } from '@docx-editor.dev/react';
import type { DocumentSource } from '@docx-editor.dev/core/contracts/editor';
import { MergePanel } from './MergePanel';
import { sampleTemplateBytes } from './merge/sample-template';
import './styles.css';

export function MailMergeApp() {
  const editor = useRef<DocxEditorRef>(null);
  const [doc, setDoc] = useState<DocumentSource>(() => sampleTemplateBytes().buffer as ArrayBuffer);
  const [title, setTitle] = useState('Template');

  const load = useCallback((bytes: ArrayBuffer, name: string) => {
    setDoc(bytes);
    setTitle(name);
  }, []);

  return (
    <div className="mm-app">
      <DocxEditor
        ref={editor}
        document={doc}
        author="Template author"
        title={title}
        onTitleChange={setTitle}
      >
        <MergePanel editorRef={editor} onLoad={load} />
      </DocxEditor>
    </div>
  );
}
