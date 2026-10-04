// Which blocks an `insertContentControl` span covers, when it covers whole paragraphs.
//
// Word authors a BLOCK control when a selection holds whole paragraphs and an inline one when
// it holds characters. A span from offset 0 of one paragraph to the end of a later one (or to
// offset 0 of the paragraph after it, which is the same text plus the paragraph mark) is whole
// paragraphs. A span that starts or ends mid-paragraph across a paragraph mark is neither, and
// refuses rather than being rounded to one or the other.

import type { InsertableContentControlKind } from '../store/store/tree-op-content-controls.ts';
import type { TreeDocOp } from '../store/store/tree-ops.ts';
import type { AutomationStoryReads } from './reads.ts';
import { spanParagraphIds, type ResolvedRange } from './spans.ts';

export type ContentControlSpanShape =
  | { readonly kind: 'inline' }
  | { readonly kind: 'blocks'; readonly firstBlockId: string; readonly lastBlockId: string }
  | { readonly kind: 'partial' };

export function contentControlSpanShape(
  range: ResolvedRange,
  story: AutomationStoryReads
): ContentControlSpanShape {
  if (range.start.paragraphId === range.end.paragraphId) return { kind: 'inline' };
  if (range.start.offset !== 0) return { kind: 'partial' };
  const ids = spanParagraphIds(range, story);
  const endLength = (story.rawText(range.end.paragraphId) ?? '').length;
  const last =
    range.end.offset === 0
      ? ids[ids.length - 2]
      : range.end.offset === endLength
        ? ids[ids.length - 1]
        : undefined;
  if (last === undefined) return { kind: 'partial' };
  return { kind: 'blocks', firstBlockId: range.start.paragraphId, lastBlockId: last };
}

/** The store op that authors the control `shape` describes. */
export function contentControlInsertOp(
  shape: Exclude<ContentControlSpanShape, { kind: 'partial' }>,
  range: NonNullable<ResolvedRange>,
  operation: {
    readonly subtype: InsertableContentControlKind;
    readonly tag?: string;
    readonly title?: string;
  }
): TreeDocOp {
  const meta = {
    ...(operation.tag === undefined ? {} : { tag: operation.tag }),
    ...(operation.title === undefined ? {} : { alias: operation.title }),
  };
  return shape.kind === 'blocks'
    ? {
        op: 'wrapBlocksInContentControl',
        firstBlockId: shape.firstBlockId,
        lastBlockId: shape.lastBlockId,
        ...meta,
      }
    : {
        op: 'insertContentControl',
        paragraphId: range.start.paragraphId,
        start: range.start.offset,
        end: range.end.offset,
        type: operation.subtype,
        ...meta,
      };
}
