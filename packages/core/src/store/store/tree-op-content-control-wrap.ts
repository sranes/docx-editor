// Wrapping whole blocks in a new BLOCK-level content control (store lane).
//
// The inline insert wraps characters inside one paragraph. This wraps paragraphs and tables:
// the blocks move, unchanged and with their ids, into the new control's `w:sdtContent`, which
// takes their place in the parent. That is what Word authors when a selection covers whole
// paragraphs, and it is what makes a rule control removable with its paragraphs: deleting a
// block control with its content leaves no empty paragraph behind.

import { lockForbidsEdit } from '../package/content-control-nodes.ts';
import {
  createNodeIdAllocator,
  findNode,
  replaceChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import { WML_NAMESPACE_URI, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { controlElement, propertiesFor } from './tree-op-content-control-insert.ts';
import {
  contentControlEffect,
  contentControlLockAt,
  isWritableContentControlMetadata,
} from './tree-op-content-controls.ts';
import { fromEdit, parentOf } from './tree-op-nodes.ts';
import type { TreeDocOp, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';

type WrapOp = Extract<TreeDocOp, { op: 'wrapBlocksInContentControl' }>;

/** Block containers a control may be placed in. A paragraph holds runs, not blocks. */
const BLOCK_PARENTS = new Set([
  'body',
  'tc',
  'sdtContent',
  'hdr',
  'ftr',
  'footnote',
  'endnote',
  'txbxContent',
  'comment',
]);

const isWml = (node: OoxmlNode, localName: string): boolean =>
  node.kind !== 'textValue' &&
  node.namespaceUri === WML_NAMESPACE_URI &&
  node.localName === localName;

/** Whether a paragraph ends a section: a control must not span a section break. */
function carriesSectionMark(node: OoxmlNode): boolean {
  if (node.kind !== 'paragraph') return false;
  const properties = node.children.find((child) => isWml(child, 'pPr'));
  return properties?.children.some((child) => isWml(child, 'sectPr')) ?? false;
}

/** The sibling run `firstBlockId..lastBlockId`, or the reason it cannot be wrapped. */
export function wrappableBlocks(
  part: OoxmlPart,
  op: Pick<WrapOp, 'firstBlockId' | 'lastBlockId'>
):
  | { readonly ok: true; readonly parentId: string; readonly first: number; readonly last: number }
  | { readonly ok: false; readonly reason: TreeOpRejection } {
  const first = findNode(part, op.firstBlockId);
  const last = findNode(part, op.lastBlockId);
  if (!first || !last) return { ok: false, reason: 'unknown-block' };
  const parent = parentOf(part, op.firstBlockId);
  if (!parent || parentOf(part, op.lastBlockId)?.id !== parent.id)
    return { ok: false, reason: 'not-adjacent-siblings' };
  if (!BLOCK_PARENTS.has(parent.localName)) return { ok: false, reason: 'not-a-block' };
  const firstIndex = parent.children.findIndex((child) => child.id === op.firstBlockId);
  const lastIndex = parent.children.findIndex((child) => child.id === op.lastBlockId);
  if (firstIndex < 0 || lastIndex < firstIndex) return { ok: false, reason: 'invalid-range' };
  for (const block of parent.children.slice(firstIndex, lastIndex + 1)) {
    if (block.kind === 'textValue' || isWml(block, 'sectPr') || isWml(block, 'tcPr'))
      return { ok: false, reason: 'not-a-block' };
    if (carriesSectionMark(block)) return { ok: false, reason: 'carries-section-mark' };
  }
  return { ok: true, parentId: parent.id, first: firstIndex, last: lastIndex };
}

const LOCKS: ReadonlySet<string> = new Set([
  'unlocked',
  'sdtLocked',
  'contentLocked',
  'sdtContentLocked',
]);

/** Metadata and shape checks, for `validateTreeOp`. */
export function validateWrapBlocks(part: OoxmlPart, op: WrapOp): TreeOpRejection | null {
  for (const value of [op.tag, op.alias]) {
    if (!isWritableContentControlMetadata(value)) return 'invalid-property-value';
  }
  if (op.lock !== undefined && !LOCKS.has(op.lock)) return 'invalidArgs';
  const blocks = wrappableBlocks(part, op);
  return blocks.ok ? null : blocks.reason;
}

export function applyWrapBlocksInContentControl(
  part: OoxmlPart,
  op: WrapOp,
  options?: EditOptions
): TreeOpResult {
  const range = wrappableBlocks(part, op);
  if (!range.ok) return range;
  if (lockForbidsEdit(contentControlLockAt(part, op.firstBlockId)))
    return { ok: false, reason: 'locked' };
  const parent = findNode(part, range.parentId);
  if (!parent || parent.kind === 'textValue') return { ok: false, reason: 'tree-invariant' };
  const nextId = createNodeIdAllocator(part);
  const wrapped = parent.children.slice(range.first, range.last + 1);
  const control = controlElement(
    propertiesFor(part, { type: 'richText', tag: op.tag, alias: op.alias, lock: op.lock }, nextId),
    wrapped,
    nextId
  );
  const children = [
    ...parent.children.slice(0, range.first),
    control,
    ...parent.children.slice(range.last + 1),
  ];
  return fromEdit(
    replaceChildren(part, parent.id, children, options),
    contentControlEffect(parent.id, 'flow-structural')
  );
}
