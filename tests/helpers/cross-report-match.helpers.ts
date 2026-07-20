/**
 * cross-report-match.helpers.ts
 *
 * Pure logic for deciding whether a page's discovered filters are the SAME
 * between a source (Import mode) and target (Direct Lake) report — i.e.
 * whether the desktop app can offer one shared picker, or must fall back to
 * two separate ones.
 *
 * "Same" here means: every filter title on one side has a title-matching
 * counterpart on the other side (case/whitespace/separator-insensitive, via
 * pageNameKey — the exact same rule used for page-name matching, since both
 * are "does this name from report A mean the same thing as this name from
 * report B"), AND the two sides agree on whether it's a hierarchy filter.
 * A title match with mismatched hierarchy-ness is deliberately NOT treated
 * as identical: applying a flat-vs-hierarchy field the wrong way would
 * either silently do nothing or throw a confusing error, not "mostly work".
 *
 * No I/O, no Playwright — takes two already-discovered field lists and
 * returns a verdict. Fully unit-testable.
 */

import type { DiscoveredSlicer } from './harness.helpers';
import { pageNameKey } from './report-export.helpers';

export interface FieldMatchResult {
  /** True only when every field on both sides has a title+hierarchy-ness
   *  match on the other side — i.e. a single shared picker is safe. */
  identical: boolean;
  /** Original (non-normalized) titles that matched, one per matched pair. */
  matchedTitles: string[];
  /** Original titles present only in source (no counterpart in target). */
  onlyInSource: string[];
  /** Original titles present only in target (no counterpart in source). */
  onlyInTarget: string[];
  /** Titles that matched by name but disagree on isHierarchy between sides —
   *  a real structural difference, not just cosmetic drift. */
  hierarchyMismatch: string[];
}

/**
 * Compares the field lists discovered on one page of the source report
 * against one page of the target report. Order-independent; a field
 * appearing more than once on one side (e.g. a duplicated slicer visual)
 * still only needs ONE counterpart on the other side to count as matched.
 */
export function matchDiscoveredFields(
  sourceFields: DiscoveredSlicer[],
  targetFields: DiscoveredSlicer[],
): FieldMatchResult {
  const bySourceKey = new Map<string, DiscoveredSlicer>();
  for (const f of sourceFields) if (!bySourceKey.has(pageNameKey(f.title))) bySourceKey.set(pageNameKey(f.title), f);

  const byTargetKey = new Map<string, DiscoveredSlicer>();
  for (const f of targetFields) if (!byTargetKey.has(pageNameKey(f.title))) byTargetKey.set(pageNameKey(f.title), f);

  const matchedTitles: string[] = [];
  const hierarchyMismatch: string[] = [];
  const onlyInSource: string[] = [];

  for (const [key, srcField] of bySourceKey) {
    const tgtField = byTargetKey.get(key);
    if (!tgtField) {
      onlyInSource.push(srcField.title);
      continue;
    }
    if (srcField.isHierarchy !== tgtField.isHierarchy) {
      hierarchyMismatch.push(srcField.title);
      continue;
    }
    matchedTitles.push(srcField.title);
  }

  const matchedKeys = new Set([...matchedTitles, ...hierarchyMismatch].map(pageNameKey));
  const onlyInTarget = [...byTargetKey.entries()]
    .filter(([key]) => !matchedKeys.has(key) && !bySourceKey.has(key))
    .map(([, f]) => f.title);

  const identical =
    onlyInSource.length === 0 &&
    onlyInTarget.length === 0 &&
    hierarchyMismatch.length === 0 &&
    matchedTitles.length > 0; // two empty pages are not a meaningful "match"

  return { identical, matchedTitles, onlyInSource, onlyInTarget, hierarchyMismatch };
}
