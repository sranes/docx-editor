// `getSelection`: the reader's current selection in the body, as a span.
//
// The inverse of `selectSpan`, and limited the same way: a host with no reader answers
// `unsupported-capability`, and a selection outside the body (a header being edited, say)
// refuses rather than being reported as a body position it is not.

import type { AutomationHandleTable } from './handles.ts';
import type { PlannedOperation } from './plan-types.ts';
import type { AutomationPackageReads } from './reads.ts';
import { BODY_STORY } from './stories.ts';

/** A position as the port reports it: a canonical paragraph id and a model offset. */
export interface AutomationPortPosition {
  readonly paragraphId: string;
  readonly offset: number;
}

/** The reader's selection: where it started and where it ends now, in the body. */
export interface AutomationPortSelection {
  readonly anchor: AutomationPortPosition;
  readonly head: AutomationPortPosition;
}

const refuse = (
  code: 'unsupported-capability' | 'invalid-offset',
  message: string,
  detail: string
): PlannedOperation => ({ ok: false, error: { code, message, detail } });

export function planGetSelection(
  selection: (() => AutomationPortSelection | null) | undefined,
  hasSelection: boolean,
  handles: AutomationHandleTable,
  reads: AutomationPackageReads
): PlannedOperation {
  if (!hasSelection || !selection)
    return refuse('unsupported-capability', 'this host has no reader selection', 'selection');
  const current = selection();
  if (!current)
    return refuse('invalid-offset', 'the reader has no selection in the body', 'no-body-selection');
  const story = reads.story(BODY_STORY);
  const place = (position: AutomationPortPosition) => {
    const index = story?.indexOf(position.paragraphId) ?? -1;
    const length = story?.rawText(position.paragraphId)?.length ?? -1;
    return index >= 0 &&
      Number.isInteger(position.offset) &&
      position.offset >= 0 &&
      position.offset <= length
      ? { index, ...position }
      : null;
  };
  const anchor = place(current.anchor);
  const head = place(current.head);
  if (!anchor || !head)
    return refuse('invalid-offset', 'the selection is not in this document', 'stale-selection');
  const forward =
    anchor.index < head.index || (anchor.index === head.index && anchor.offset <= head.offset);
  const [start, end] = forward ? [anchor, head] : [head, anchor];
  return {
    ok: true,
    kind: 'query',
    value: {
      kind: 'span',
      span: {
        start: {
          paragraph: handles.paragraph(start.paragraphId, BODY_STORY),
          offset: start.offset,
        },
        end: { paragraph: handles.paragraph(end.paragraphId, BODY_STORY), offset: end.offset },
      },
    },
  };
}
