// `applyVariables`: replace `{{name}}` template placeholders in the body, headers and footers.
//
// Every story's ops commit through `applyTreeOpsAtomic`, so the command is one undo step and
// all-or-nothing: a refusal in a header rolls the body back too. A replacement is the same
// op pair the automation `replaceSpan` writes — insert with right bias, then delete — so the
// value takes the formatting of the run the placeholder started in. Matching runs over the
// visible paragraph text, so a placeholder split across runs still matches. Notes are not
// searched; suggesting mode refuses, because these ops would not record tracked changes.

import type { StoryScope } from '@docx-editor.dev/core/store';
import { paragraphTextOf, type TreeDocOp } from '../store/store/tree-ops.ts';
import { storyParagraphs } from '../store/package/story-blocks.ts';
import { projectVisibleParagraphText } from '../store/store/text-projection.ts';
import type { OoxmlNode, OoxmlParagraphNode, OoxmlPart } from '../store/package/ooxml-tree.ts';
import type { ExecResult } from '../contracts/editor.ts';
import type { PaginatedSurface } from './paginated-surface-contract.ts';
import { storyScopeOfNodeId } from './surface-scope.ts';
import { checkVariables } from './variables-check.ts';

/** The story root of a part: the `w:body` of the main document, the part root otherwise. */
function storyRoot(part: OoxmlPart): OoxmlNode {
  return part.root.children.find((node) => node.kind === 'body') ?? part.root;
}

/** One paragraph's replacements, latest first, so earlier offsets stay valid. */
function paragraphOps(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  entries: readonly (readonly [string, string])[]
): TreeDocOp[] {
  const raw = paragraphTextOf(part, paragraph.id) ?? '';
  if (!raw.includes('{{')) return [];
  const projected = projectVisibleParagraphText(paragraph, raw);
  const found: { start: number; end: number; text: string }[] = [];
  for (const [name, text] of entries) {
    const occurrences = projected.findOccurrences(`{{${name}}}`, Number.MAX_SAFE_INTEGER, {
      matchCase: true,
    });
    for (const match of occurrences.matches)
      found.push({ start: match.rawStart, end: match.rawEnd, text });
  }
  found.sort((a, b) => b.start - a.start);
  const ops: TreeDocOp[] = [];
  let floor = Number.POSITIVE_INFINITY;
  for (const { start, end, text } of found) {
    // Two names cannot claim one stretch: `{{` and `}}` delimit, so matches never overlap.
    if (end > floor) continue;
    floor = start;
    if (text.length > 0)
      ops.push({ op: 'insertText', paragraphId: paragraph.id, offset: start, text, bias: 'right' });
    ops.push({
      op: 'deleteText',
      paragraphId: paragraph.id,
      start: start + text.length,
      end: end + text.length,
    });
  }
  return ops;
}

/**
 * Runs `applyVariables` over the mounted surface. Answers `null` on success, so the caller
 * derives `changed` from the package revision, and a refusal otherwise.
 */
export function execApplyVariables(surface: PaginatedSurface, values: unknown): ExecResult | null {
  const checked = checkVariables(values);
  if (!checked.ok) return { ok: false, code: 'invalidArgs', reason: checked.reason };
  if (checked.entries.length === 0) return null;
  if (surface.editingMode() === 'suggest')
    return {
      ok: false,
      code: 'unsupported',
      reason: 'applyVariables does not record tracked changes; turn off suggesting mode first',
    };
  const session = surface.session;
  const mainPart = session.part().name;
  const groups: { scope: StoryScope; ops: TreeDocOp[] }[] = [];
  for (const part of session.storyParts()) {
    const scope = storyScopeOfNodeId(session, part.root.id, { kind: 'body' });
    // A part that resolves to the body but is not the main part is a notes part.
    if (scope.kind === 'body' && part.name !== mainPart) continue;
    const ops = storyParagraphs(storyRoot(part)).flatMap((paragraph) =>
      paragraph.kind === 'paragraph' ? paragraphOps(part, paragraph, checked.entries) : []
    );
    if (ops.length > 0) groups.push({ scope, ops });
  }
  if (groups.length === 0) return null;
  let applied: { committed: boolean; reason?: unknown } | undefined;
  surface.commitReviewOps(() => {
    applied = session.applyTreeOpsAtomic(groups);
    return applied;
  });
  if (!applied?.committed)
    return {
      ok: false,
      code: 'locked',
      reason:
        typeof applied?.reason === 'string'
          ? `a placeholder cannot be replaced: ${applied.reason}`
          : 'a placeholder cannot be replaced',
    };
  return null;
}
