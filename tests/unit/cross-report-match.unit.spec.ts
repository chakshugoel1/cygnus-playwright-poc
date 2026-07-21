/**
 * cross-report-match.unit.spec.ts
 *
 * Pure-logic tests for the cross-report filter matching feature: the
 * name-normalization rule shared by page names and filter titles
 * (pageNameKey), matchDiscoveredFields (deciding whether two reports' pages
 * have identical filters), and selectionsForSide (applying a scenario's
 * per-side picks only to the side they were picked for). Run with:
 * npm run test:unit
 */

import { test, expect } from '@playwright/test';
import { pageNameKey } from '../helpers/report-export.helpers';
import { matchDiscoveredFields } from '../helpers/cross-report-match.helpers';
import { selectionsForSide } from '../helpers/slicer-config.helpers';
import type { DiscoveredSlicer } from '../helpers/harness.helpers';

// ── pageNameKey's separator-stripping rule ────────────────────────────────────

test.describe('pageNameKey', () => {
  test('strips underscores, hyphens, and spaces entirely (not just to one style)', () => {
    expect(pageNameKey('a_bc')).toBe(pageNameKey('abc'));
    expect(pageNameKey('a-bc')).toBe(pageNameKey('abc'));
    expect(pageNameKey('a bc')).toBe(pageNameKey('abc'));
    expect(pageNameKey('Recruiting_Weg')).toBe(pageNameKey('Recruiting Weg'));
    expect(pageNameKey('RECRUITING-WEG')).toBe(pageNameKey('recruitingweg'));
  });

  test('is trimmed and case-insensitive', () => {
    expect(pageNameKey('  Vertragsversand  ')).toBe(pageNameKey('Vertragsversand'));
    expect(pageNameKey('Vertragsrücklauf - Cluster')).toBe(pageNameKey('vertragsrücklauf - cluster'));
  });

  test('does not collapse genuinely different names', () => {
    expect(pageNameKey('Page 1')).not.toBe(pageNameKey('Page 2'));
    expect(pageNameKey('Absagen')).not.toBe(pageNameKey('Absagegründe'));
  });
});

// ── matchDiscoveredFields ──────────────────────────────────────────────────────

function field(title: string, isHierarchy = false): DiscoveredSlicer {
  return {
    name: `v-${title}`, title, kind: isHierarchy ? 'tree' : 'flat',
    targetLabel: `T.${title}`, targets: [{ table: 'T', column: title }],
    options: ['A', 'B'], selected: [], isHierarchy, errorMessage: null,
  };
}

/** Same title, but an explicit (possibly different) underlying binding —
 *  for the duplicate-title-different-field scenario. */
function fieldWithBinding(title: string, binding: string, isHierarchy = false): DiscoveredSlicer {
  return {
    name: `v-${title}-${binding}`, title, kind: isHierarchy ? 'tree' : 'flat',
    targetLabel: binding, targets: [{ table: 'T', column: binding }],
    options: ['A', 'B'], selected: [], isHierarchy, errorMessage: null,
  };
}

test.describe('matchDiscoveredFields', () => {
  test('identical when every title matches on both sides (exact)', () => {
    const src = [field('Recruitingweg'), field('Status')];
    const tgt = [field('Recruitingweg'), field('Status')];
    const r = matchDiscoveredFields(src, tgt);
    expect(r.identical).toBe(true);
    expect(r.matchedTitles.sort()).toEqual(['Recruitingweg', 'Status']);
    expect(r.onlyInSource).toEqual([]);
    expect(r.onlyInTarget).toEqual([]);
  });

  test('identical when titles match only after normalization (case/separator drift)', () => {
    const src = [field('Recruiting_Weg')];
    const tgt = [field('RECRUITING WEG')];
    const r = matchDiscoveredFields(src, tgt);
    expect(r.identical).toBe(true);
    expect(r.matchedTitles).toEqual(['Recruiting_Weg']);
  });

  test('not identical when a field exists only in source', () => {
    const src = [field('Recruitingweg'), field('Status')];
    const tgt = [field('Recruitingweg')];
    const r = matchDiscoveredFields(src, tgt);
    expect(r.identical).toBe(false);
    expect(r.onlyInSource).toEqual(['Status']);
    expect(r.onlyInTarget).toEqual([]);
  });

  test('not identical when a field exists only in target', () => {
    const src = [field('Recruitingweg')];
    const tgt = [field('Recruitingweg'), field('Region')];
    const r = matchDiscoveredFields(src, tgt);
    expect(r.identical).toBe(false);
    expect(r.onlyInTarget).toEqual(['Region']);
  });

  test('a title match with mismatched hierarchy-ness is NOT identical', () => {
    const src = [field('BU / DU', true)];
    const tgt = [field('BU / DU', false)];
    const r = matchDiscoveredFields(src, tgt);
    expect(r.identical).toBe(false);
    expect(r.hierarchyMismatch).toEqual(['BU / DU']);
    expect(r.matchedTitles).toEqual([]);
  });

  test('two empty pages are not treated as a meaningful match', () => {
    const r = matchDiscoveredFields([], []);
    expect(r.identical).toBe(false);
  });

  test('duplicate titles on one side only need one counterpart on the other', () => {
    const src = [field('Status'), field('Status')]; // duplicated slicer visual
    const tgt = [field('Status')];
    const r = matchDiscoveredFields(src, tgt);
    expect(r.identical).toBe(true);
  });

  test('same title, SAME binding repeated (a genuinely duplicated visual) still collapses to one variant and matches normally', () => {
    const src = [fieldWithBinding('Status', 'T.Status'), fieldWithBinding('Status', 'T.Status')];
    const tgt = [fieldWithBinding('Status', 'T.Status')];
    const r = matchDiscoveredFields(src, tgt);
    expect(r.identical).toBe(true);
    expect(r.duplicateTitleConflicts).toEqual([]);
  });

  // Real production case: two slicers on the source report share a title
  // ("Bewerbungseingang Cluster") but are bound to two DIFFERENT fields.
  // Matching by title alone would silently pair the extra source variant to
  // whichever single target field shares that title - wrong for one of them,
  // with no warning. This must never auto-match, regardless of what the
  // single target-side field's binding is.
  test('same title, different field binding on source only: NOT identical, flagged as a duplicate-title conflict', () => {
    const src = [
      fieldWithBinding('Cluster', "'Germany - Extract'.Recruitment Area"),
      fieldWithBinding('Cluster', "'Fabric_application'.Recruitment Area"),
    ];
    const tgt = [fieldWithBinding('Cluster', "'V_FACT_APPLICATION_GERMANY'.RECRUITMENT_AREA")];
    const r = matchDiscoveredFields(src, tgt);
    expect(r.identical).toBe(false);
    expect(r.duplicateTitleConflicts).toEqual(['Cluster']);
    expect(r.matchedTitles).toEqual([]);
  });

  test('same title, different bindings with MATCHING variant counts on both sides: still not auto-matched (pairing would be a guess)', () => {
    const src = [fieldWithBinding('Card', 'A.One'), fieldWithBinding('Card', 'A.Two')];
    const tgt = [fieldWithBinding('Card', 'B.One'), fieldWithBinding('Card', 'B.Two')];
    const r = matchDiscoveredFields(src, tgt);
    expect(r.identical).toBe(false);
    expect(r.duplicateTitleConflicts).toEqual(['Card']);
  });

  test('a duplicate-title conflict on one title does not block other, cleanly-matched titles from matching', () => {
    const src = [
      fieldWithBinding('Cluster', 'A.1'), fieldWithBinding('Cluster', 'A.2'),
      field('Status'),
    ];
    const tgt = [fieldWithBinding('Cluster', 'B.1'), field('Status')];
    const r = matchDiscoveredFields(src, tgt);
    expect(r.identical).toBe(false); // overall still not safe to share
    expect(r.duplicateTitleConflicts).toEqual(['Cluster']);
    expect(r.matchedTitles).toEqual(['Status']);
  });
});

// ── selectionsForSide ──────────────────────────────────────────────────────────

interface FakeSel { values: string[]; title: string; side?: 'source' | 'target' }

test.describe('selectionsForSide', () => {
  test('keeps untagged selections for either side (default = applies to both)', () => {
    const sels: FakeSel[] = [{ values: ['x'], title: 'Status' }];
    expect(selectionsForSide(sels, 'source')).toEqual(sels);
    expect(selectionsForSide(sels, 'target')).toEqual(sels);
  });

  test('keeps a selection only for the side it is tagged for', () => {
    const sels: FakeSel[] = [
      { values: ['x'], title: 'Shared' },
      { values: ['y'], title: 'SourceOnly', side: 'source' },
      { values: ['z'], title: 'TargetOnly', side: 'target' },
    ];
    expect(selectionsForSide(sels, 'source').map(s => s.title)).toEqual(['Shared', 'SourceOnly']);
    expect(selectionsForSide(sels, 'target').map(s => s.title)).toEqual(['Shared', 'TargetOnly']);
  });
});
