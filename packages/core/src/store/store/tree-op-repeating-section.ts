// Repeating-section items (store lane): add one by cloning an item, remove one by index.
//
// A `w15:repeatingSection` control holds `w15:repeatingSectionItem` controls; Word's "+" adds a
// copy of an item after it and "x" removes one, never the last. A copy must not duplicate any
// identity the document holds once: control ids and paragraph ids are minted afresh, bookmarks
// are dropped from the copy, and an item that holds anything else with a document-wide identity
// (a comment, a note reference, a tracked change, a drawing, a permission range) is refused
// rather than copied into two places.

import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import {
  createNodeIdAllocator,
  findNode,
  replaceChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import { WML_NAMESPACE_URI } from '../package/ooxml-shared.ts';
import { mintBuildingBlockIdentities } from './building-block-identities.ts';
import { cloneWithFreshIds, contentControlEffect } from './tree-op-content-controls.ts';
import {
  contentControlPropertiesOf,
  fromEdit,
  isContentControlNode,
  isRepeatingSectionControl,
  sdtPrChild,
} from './tree-op-nodes.ts';
import type { TreeDocOp, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';

const W15_NAMESPACE_URI = 'http://schemas.microsoft.com/office/word/2012/wordml';

type ItemOp = Extract<TreeDocOp, { op: 'addRepeatingSectionItem' | 'removeRepeatingSectionItem' }>;

/** Elements whose identity is document-wide: a copy would point two places at one thing. */
const NOT_COPYABLE = new Set([
  'commentRangeStart',
  'commentRangeEnd',
  'commentReference',
  'footnoteReference',
  'endnoteReference',
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
  'drawing',
  'pict',
  'object',
  'permStart',
  'permEnd',
  'altChunk',
  'subDoc',
]);
/** Dropped from the copy: a bookmark name is unique, and the original keeps it. */
const DROPPED = new Set(['bookmarkStart', 'bookmarkEnd']);
const MAX_ITEM_NODES = 50_000;

const isWml = (node: OoxmlNode, names: ReadonlySet<string>): boolean =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && names.has(node.localName);

/** The section's `w:sdtContent` and the items in it, in order. */
export function sectionItems(
  part: OoxmlPart,
  controlId: string
):
  | { readonly ok: true; readonly content: OoxmlElement; readonly items: readonly OoxmlNode[] }
  | { readonly ok: false; readonly reason: TreeOpRejection } {
  const section = findNode(part, controlId);
  if (!section) return { ok: false, reason: 'unknown-content-control' };
  if (
    section.kind === 'textValue' ||
    !isContentControlNode(section) ||
    !isRepeatingSectionControl(section)
  )
    return { ok: false, reason: 'not-a-content-control' };
  const children: readonly OoxmlNode[] = section.children;
  const content = children.find(
    (child): child is OoxmlElement =>
      child.kind !== 'textValue' && child.kind === 'contentControlContent'
  );
  if (!content) return { ok: false, reason: 'tree-invariant' };
  const items = content.children.filter(
    (child: OoxmlNode) =>
      isContentControlNode(child) &&
      sdtPrChild(contentControlPropertiesOf(child), 'repeatingSectionItem', W15_NAMESPACE_URI) !==
        undefined
  );
  return { ok: true, content, items };
}

/** Why `item` cannot be copied, or null. Bounded: a deep or huge item refuses. */
function copyRefusal(item: OoxmlNode): TreeOpRejection | null {
  let nodes = 0;
  const visit = (node: OoxmlNode): TreeOpRejection | null => {
    if (++nodes > MAX_ITEM_NODES) return 'unsupported';
    if (node.kind === 'textValue') return null;
    if (isWml(node, NOT_COPYABLE)) return 'unsupported';
    for (const child of node.children) {
      const refused = visit(child);
      if (refused) return refused;
    }
    return null;
  };
  return visit(item);
}

function withoutBookmarks(node: OoxmlNode): OoxmlNode {
  if (node.kind === 'textValue') return node;
  return {
    ...node,
    children: node.children.filter((child) => !isWml(child, DROPPED)).map(withoutBookmarks),
  } as OoxmlNode;
}

/** The checks both ops share; `can` and the applier answer the same. */
export function validateRepeatingSectionOp(part: OoxmlPart, op: ItemOp): TreeOpRejection | null {
  const found = sectionItems(part, op.controlId);
  if (!found.ok) return found.reason;
  const count = found.items.length;
  // A section holding no item is a shape this lane does not repeat from.
  if (count === 0) return 'unsupported';
  if (op.op === 'removeRepeatingSectionItem') {
    if (!Number.isInteger(op.index) || op.index < 0 || op.index >= count) return 'invalid-range';
    // Word keeps at least one item; the section would otherwise have nothing to repeat.
    return count <= 1 ? 'unsupported' : null;
  }
  const index = op.index ?? count - 1;
  if (!Number.isInteger(index) || index < 0 || index >= count) return 'invalid-range';
  return copyRefusal(found.items[index]!);
}

export function applyRepeatingSectionOp(
  part: OoxmlPart,
  op: ItemOp,
  options?: EditOptions
): TreeOpResult {
  const refused = validateRepeatingSectionOp(part, op);
  if (refused) return { ok: false, reason: refused };
  const { content, items } = sectionItems(part, op.controlId) as Extract<
    ReturnType<typeof sectionItems>,
    { ok: true }
  >;
  let children: OoxmlNode[];
  if (op.op === 'removeRepeatingSectionItem') {
    const gone = items[op.index]!;
    children = content.children.filter((child) => child.id !== gone.id);
  } else {
    const source = items[op.index ?? items.length - 1]!;
    const copy = mintBuildingBlockIdentities(part)(
      cloneWithFreshIds(withoutBookmarks(source), createNodeIdAllocator(part))
    );
    if (!copy) return { ok: false, reason: 'unsupported' };
    const at = content.children.findIndex((child) => child.id === source.id);
    children = [...content.children.slice(0, at + 1), copy, ...content.children.slice(at + 1)];
  }
  return fromEdit(
    replaceChildren(part, content.id, children, options),
    contentControlEffect(op.controlId, 'flow-structural')
  );
}
