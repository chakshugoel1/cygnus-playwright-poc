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
 * A title that resolves to MORE THAN ONE distinct field binding on either
 * side (two slicers sharing a display title but bound to different
 * table.column targets — observed in production) is likewise never treated
 * as identical, even when both sides happen to have the same count of
 * variants: pairing them by array order would be a guess, and a wrong guess
 * here means a user's filter pick silently lands on the wrong field with no
 * warning. Same "don't auto-pair an ambiguous duplicate" principle already
 * used for visual matching in excel-compare.helpers.ts.
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
  /** Titles that resolve to more than one distinct field binding on at
   *  least one side — same title, can't safely tell which source variant
   *  corresponds to which target variant, so never auto-matched. */
  duplicateTitleConflicts: string[];
}

/** Identity for one distinct binding under a title — same title, different
 *  targetLabel (table.column) counts as a different variant. */
function variantKey(f: DiscoveredSlicer): string {
  return `${f.targetLabel ?? ''}`;
}

/** Groups fields by normalized title, deduping to distinct (title, binding)
 *  variants — a slicer repeated with an IDENTICAL binding (e.g. the same
 *  field discovered on more than one page) still collapses to one variant;
 *  only a title shared by genuinely DIFFERENT bindings produces more than
 *  one entry per title. */
function groupByTitle(fields: DiscoveredSlicer[]): Map<string, DiscoveredSlicer[]> {
  const byTitle = new Map<string, DiscoveredSlicer[]>();
  for (const f of fields) {
    const titleKey = pageNameKey(f.title);
    let variants = byTitle.get(titleKey);
    if (!variants) { variants = []; byTitle.set(titleKey, variants); }
    if (!variants.some(v => variantKey(v) === variantKey(f))) variants.push(f);
  }
  return byTitle;
}

/**
 * Compares the field lists discovered on one page of the source report
 * against one page of the target report. Order-independent; a field
 * appearing more than once on one side WITH THE SAME binding (e.g. a
 * duplicated slicer visual) still only needs ONE counterpart on the other
 * side to count as matched — see duplicateTitleConflicts above for the
 * different-binding case, which is handled separately and conservatively.
 */
export function matchDiscoveredFields(
  sourceFields: DiscoveredSlicer[],
  targetFields: DiscoveredSlicer[],
): FieldMatchResult {
  const bySourceTitle = groupByTitle(sourceFields);
  const byTargetTitle = groupByTitle(targetFields);

  const matchedTitles: string[] = [];
  const hierarchyMismatch: string[] = [];
  const duplicateTitleConflicts: string[] = [];
  const onlyInSource: string[] = [];

  for (const [key, srcVariants] of bySourceTitle) {
    const tgtVariants = byTargetTitle.get(key);
    if (!tgtVariants) {
      onlyInSource.push(srcVariants[0].title);
      continue;
    }
    if (srcVariants.length > 1 || tgtVariants.length > 1) {
      duplicateTitleConflicts.push(srcVariants[0].title);
      continue;
    }
    const srcField = srcVariants[0];
    const tgtField = tgtVariants[0];
    if (srcField.isHierarchy !== tgtField.isHierarchy) {
      hierarchyMismatch.push(srcField.title);
      continue;
    }
    matchedTitles.push(srcField.title);
  }

  const matchedKeys = new Set(
    [...matchedTitles, ...hierarchyMismatch, ...duplicateTitleConflicts].map(pageNameKey),
  );
  const onlyInTarget = [...byTargetTitle.entries()]
    .filter(([key]) => !matchedKeys.has(key) && !bySourceTitle.has(key))
    .map(([, variants]) => variants[0].title);

  const identical =
    onlyInSource.length === 0 &&
    onlyInTarget.length === 0 &&
    hierarchyMismatch.length === 0 &&
    duplicateTitleConflicts.length === 0 &&
    matchedTitles.length > 0; // two empty pages are not a meaningful "match"

  return { identical, matchedTitles, onlyInSource, onlyInTarget, hierarchyMismatch, duplicateTitleConflicts };
}
