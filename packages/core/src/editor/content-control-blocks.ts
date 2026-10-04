// The live selection as whole body paragraphs, for a block-level content control.
//
// A selection from offset 0 of one paragraph to the end of a later one, or to offset 0 of the
// paragraph after it (the same text plus the paragraph mark), holds whole paragraphs. Body
// only: a header or footer selection orders against its own story, which this does not read.

import { storyParagraphs } from '../store/package/story-blocks.ts';
import { paragraphTextOf } from '../store/store/tree-ops.ts';
import type { PaginatedSurface } from './paginated-surface-contract.ts';

export function wholeParagraphSelection(
  surface: PaginatedSurface
): { readonly firstBlockId: string; readonly lastBlockId: string } | null {
  if (surface.activeScope().kind !== 'body') return null;
  const part = surface.session.part();
  const body = part.root.children.find((node) => node.kind === 'body') ?? part.root;
  const ids = storyParagraphs(body).map((node) => node.id);
  const { anchor, head } = surface.state().selection;
  const anchorIndex = ids.indexOf(anchor.paragraphId);
  const headIndex = ids.indexOf(head.paragraphId);
  if (anchorIndex < 0 || headIndex < 0 || anchorIndex === headIndex) return null;
  const [from, to, fromIndex, toIndex] =
    anchorIndex < headIndex
      ? [anchor, head, anchorIndex, headIndex]
      : [head, anchor, headIndex, anchorIndex];
  if (from.offset !== 0) return null;
  const toLength = (paragraphTextOf(part, to.paragraphId) ?? '').length;
  const lastIndex = to.offset === 0 ? toIndex - 1 : to.offset === toLength ? toIndex : -1;
  if (lastIndex < fromIndex) return null;
  return { firstBlockId: from.paragraphId, lastBlockId: ids[lastIndex]! };
}
