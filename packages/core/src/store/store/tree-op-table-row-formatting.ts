// The paragraph and run formatting a fresh table cell inherits from the cell it copies.
//
// Word gives an inserted row the formatting of the row it was inserted from: alignment,
// spacing and paragraph style, and the font of the first run, so text typed or written into the
// new cell looks like the row above. The new cell carries that as its paragraph's `w:pPr` and
// one EMPTY run holding the source's first `w:rPr`; a text insert lands in that run.
//
// The nodes are copies of the same document's own properties, not file input arriving at a new
// boundary. Revision markup is stripped so an untracked row does not claim tracked formatting
// changes, and a section break never travels into a cell.

import type { OoxmlNode } from '../package/ooxml-tree.ts';
import { WML_NAMESPACE_URI } from '../package/ooxml-tree.ts';
import { cloneWithFreshIds } from './tree-op-content-controls.ts';

/** Children never copied: revision history, and a section break that has no place in a cell. */
const STRIPPED = new Set(['pPrChange', 'rPrChange', 'ins', 'del', 'moveFrom', 'moveTo', 'sectPr']);

const isWml = (node: OoxmlNode, localName: string): boolean =>
  node.kind !== 'textValue' &&
  node.namespaceUri === WML_NAMESPACE_URI &&
  node.localName === localName;

function stripped(node: OoxmlNode): OoxmlNode {
  if (node.kind === 'textValue') return node;
  return {
    ...node,
    children: node.children
      .filter((child) => !(child.kind !== 'textValue' && STRIPPED.has(child.localName)))
      .map(stripped),
  } as OoxmlNode;
}

function nodeCount(node: OoxmlNode): number {
  return node.kind === 'textValue'
    ? 1
    : 1 + node.children.reduce((total, child) => total + nodeCount(child), 0);
}

/** The nodes a fresh cell's paragraph should start with, and how many nodes they are. */
export function inheritedCellFormatting(
  sourceCell: OoxmlNode,
  nextId: () => string
): { readonly children: readonly OoxmlNode[]; readonly size: number } {
  if (sourceCell.kind === 'textValue') return { children: [], size: 0 };
  const paragraph = sourceCell.children.find((child) => child.kind === 'paragraph');
  if (!paragraph || paragraph.kind === 'textValue') return { children: [], size: 0 };
  const children: OoxmlNode[] = [];
  const properties = paragraph.children.find((child) => isWml(child, 'pPr'));
  if (properties) children.push(cloneWithFreshIds(stripped(properties), nextId));
  const run = paragraph.children.find(
    (child) => child.kind === 'run' && child.children.some((part) => isWml(part, 'rPr'))
  );
  if (run && run.kind !== 'textValue') {
    const runProperties = run.children.find((part) => isWml(part, 'rPr'))!;
    children.push(
      cloneWithFreshIds(stripped({ ...run, children: [runProperties] } as OoxmlNode), nextId)
    );
  }
  return { children, size: children.reduce((total, child) => total + nodeCount(child), 0) };
}
