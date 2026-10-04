// `addRepeatingSectionItem` / `removeRepeatingSectionItem`: the store op, checked at plan time.
//
// The store validates the same way when the op lands; checking here as well answers a refusal
// with its reason in the planner's vocabulary instead of a refused transaction.

import type { OoxmlNode, OoxmlPart } from '../store/package/ooxml-tree.ts';
import {
  contentControlPropertiesOf,
  isRepeatingSectionControl,
  sdtPrChild,
} from '../store/store/tree-op-nodes.ts';
import {
  sectionItems,
  validateRepeatingSectionOp,
} from '../store/store/tree-op-repeating-section.ts';
import { collectStoryParagraphs } from '../store/package/story-blocks.ts';
import type { TreeDocOp } from '../store/store/tree-ops.ts';
import type { PlannedOperation } from './plan-types.ts';

const W15_NAMESPACE_URI = 'http://schemas.microsoft.com/office/word/2012/wordml';

type RepeatingSectionOperation =
  | { readonly op: 'addRepeatingSectionItem'; readonly index?: number }
  | { readonly op: 'removeRepeatingSectionItem'; readonly index: number };

const REASONS: Readonly<Record<string, string>> = {
  'not-a-content-control': 'that control is not a repeating section',
  'invalid-range': 'there is no item at that index',
  unsupported:
    'that item cannot be repeated: it holds a comment, note, tracked change, drawing or permission range, it is the last item, or the section has none',
};

export function planRepeatingSectionOp(
  operation: RepeatingSectionOperation,
  controlId: string,
  part: OoxmlPart
): { readonly ok: true; readonly op: TreeDocOp; readonly created: number } | PlannedOperation {
  const op: TreeDocOp =
    operation.op === 'addRepeatingSectionItem'
      ? {
          op: 'addRepeatingSectionItem',
          controlId,
          ...(operation.index === undefined ? {} : { index: operation.index }),
        }
      : { op: 'removeRepeatingSectionItem', controlId, index: operation.index };
  const refused = validateRepeatingSectionOp(part, op);
  if (!refused) {
    // A copy brings new paragraphs, which the planner must expect or it refuses the commit.
    const found = op.op === 'addRepeatingSectionItem' ? sectionItems(part, controlId) : null;
    const source = found?.ok ? found.items[op.index ?? found.items.length - 1] : undefined;
    const paragraphs: OoxmlNode[] = [];
    // The walk enters a block control from its parent, so hand it the item as a child.
    if (source) collectStoryParagraphs([source], paragraphs, 0);
    return { ok: true, op, created: paragraphs.length };
  }
  return {
    ok: false,
    error: {
      code: refused === 'invalid-range' ? 'invalid-offset' : 'unsupported-content',
      message: REASONS[refused] ?? 'that repeating-section change is refused',
      detail: refused,
    },
  };
}

/**
 * The subtype a repeating section or its item answers. The package read calls both `untyped`,
 * because the `w15` extensions sit outside the ECMA-376 type choice; Word reports them as
 * `RepeatingSection` controls, so a caller looking for a section can find one.
 */
export function repeatingSubtypeOf(
  node: OoxmlNode
): 'repeatingSection' | 'repeatingSectionItem' | null {
  if (isRepeatingSectionControl(node)) return 'repeatingSection';
  const properties = contentControlPropertiesOf(node);
  return sdtPrChild(properties, 'repeatingSectionItem', W15_NAMESPACE_URI)
    ? 'repeatingSectionItem'
    : null;
}
