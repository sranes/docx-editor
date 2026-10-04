// Projected text search with model-offset results.

import { SEARCH_MATCH_LIMIT } from '../store/store/text-match.ts';
import type { ProjectedParagraphText } from '../store/store/text-projection.ts';
import type { AutomationHandleTable } from './handles.ts';
import type { AutomationSearchOptions } from './operations.ts';
import type { AutomationSpan } from './protocol.ts';
import type { AutomationStoryReads } from './reads.ts';
import { spanParagraphIds, type ResolvedSpan } from './spans.ts';

export type ProjectedSearchResult =
  | { readonly ok: true; readonly spans: readonly AutomationSpan[] }
  | { readonly ok: false; readonly paragraphId: string };

/**
 * Wildcard hits in one paragraph's visible text, mapped to raw offsets. A hit the projection
 * cannot map to one editable range, and an empty hit, is not a match.
 */
function wildcardMatches(
  projected: ProjectedParagraphText,
  pattern: RegExp,
  limit: number,
  bounds: { readonly from?: number; readonly to?: number }
): { rawStart: number; rawEnd: number }[] {
  const matches: { rawStart: number; rawEnd: number }[] = [];
  const to = Math.min(projected.text.length, bounds.to ?? projected.text.length);
  const scan = new RegExp(pattern.source, pattern.flags);
  scan.lastIndex = Math.max(0, bounds.from ?? 0);
  while (matches.length < limit) {
    const hit = scan.exec(projected.text);
    if (!hit) break;
    const end = hit.index + hit[0].length;
    if (end > to) break;
    if (hit[0].length === 0) {
      scan.lastIndex = hit.index + 1;
      continue;
    }
    const raw = projected.rawRange(hit.index, end);
    if (raw) matches.push({ rawStart: raw.start, rawEnd: raw.end });
  }
  return matches;
}

/** Search one story view and map each visible hit back to one editable model span. */
export function projectedSearchSpans(
  reads: AutomationStoryReads,
  scope: ResolvedSpan,
  handles: AutomationHandleTable,
  text: string,
  options: AutomationSearchOptions | undefined,
  /** A compiled `matchWildcards` pattern; when set, `text` is not scanned for literally. */
  wildcard?: RegExp
): ProjectedSearchResult {
  let budget = Math.min(options?.limit ?? SEARCH_MATCH_LIMIT, SEARCH_MATCH_LIMIT);
  const spans: AutomationSpan[] = [];
  const ids = spanParagraphIds(scope, reads);
  const last = ids.length - 1;
  for (const [position, paragraphId] of ids.entries()) {
    if (budget <= 0) break;
    const projected = reads.projectedText(paragraphId, options?.projection ?? 'allMarkup');
    if (!projected) continue;
    const bounds = {
      ...(position === 0 && scope ? { from: projected.projectedOffset(scope.start.offset) } : {}),
      ...(position === last && scope ? { to: projected.projectedOffset(scope.end.offset) } : {}),
    };
    const found = wildcard
      ? { matches: wildcardMatches(projected, wildcard, budget, bounds) }
      : projected.findOccurrences(text, budget, {
          matchCase: options?.matchCase === true,
          wholeWord: options?.matchWholeWord === true,
          ...bounds,
        });
    for (const occurrence of found.matches) {
      if (position === 0 && scope && occurrence.rawStart < scope.start.offset) continue;
      if (position === last && scope && occurrence.rawEnd > scope.end.offset) continue;
      const paragraph = handles.paragraph(paragraphId, reads.story);
      spans.push({
        start: { paragraph, offset: occurrence.rawStart },
        end: { paragraph, offset: occurrence.rawEnd },
      });
      budget -= 1;
    }
  }
  return { ok: true, spans };
}
