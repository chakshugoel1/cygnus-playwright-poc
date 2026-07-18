/**
 * slicer-config.unit.spec.ts
 *
 * Pure-logic unit tests for slicer-config.helpers.ts — no browser, no network,
 * no auth. These cover the densest logic in that module: filter-application
 * planning (the global-batch split), scenario validation (including the
 * duplicate-name/slug guards), the case-insensitive page lookup, and slug
 * building. Run with: npm run test:unit
 */

import { test, expect } from '@playwright/test';
import {
  validateScenarios,
  scenarioSlug,
  getForPageCI,
  selectionsForPage,
  planFilterApplication,
  type SlicerScenario,
} from '../helpers/slicer-config.helpers';

// ── validateScenarios ─────────────────────────────────────────────────────────

test.describe('validateScenarios', () => {
  test('accepts a well-formed visual-based scenario', () => {
    const sc: SlicerScenario = {
      name: 'CXO-BU1',
      pages: { CXO: [{ visualName: 'v1', values: ['BU1'], isHierarchy: true }] },
    };
    expect(() => validateScenarios([sc])).not.toThrow();
  });

  test('accepts a well-formed field-based scenario (targets, no visualName)', () => {
    const sc: SlicerScenario = {
      name: 'Status-Active',
      pages: { CXO: [{ title: 'Status', values: ['Active'], targets: [{ table: 'T', column: 'Status' }] }] },
    };
    expect(() => validateScenarios([sc])).not.toThrow();
  });

  test('rejects a scenario missing a name', () => {
    const sc = { name: '', pages: {} } as SlicerScenario;
    expect(() => validateScenarios([sc])).toThrow(/missing a "name"/);
  });

  test('rejects a selection with neither visualName nor targets', () => {
    const sc: SlicerScenario = {
      name: 'bad', pages: { CXO: [{ values: ['x'] }] },
    };
    expect(() => validateScenarios([sc])).toThrow(/needs either "visualName".*or "targets"/s);
  });

  test('rejects a hierarchy selection without a visualName', () => {
    const sc: SlicerScenario = {
      name: 'bad-hier',
      pages: { CXO: [{ values: ['BU1'], isHierarchy: true, targets: [{ table: 'T', column: 'C' }] }] },
    };
    expect(() => validateScenarios([sc])).toThrow(/hierarchy selections require\s+"visualName"/s);
  });

  test('rejects an empty values array', () => {
    const sc: SlicerScenario = {
      name: 'no-vals', pages: { CXO: [{ visualName: 'v1', values: [] }] },
    };
    expect(() => validateScenarios([sc])).toThrow(/"values" must be a non-empty array/);
  });

  test('rejects two scenarios with the same name (case-insensitive)', () => {
    const a: SlicerScenario = { name: 'Sales', pages: { CXO: [{ visualName: 'v', values: ['x'] }] } };
    const b: SlicerScenario = { name: 'sales', pages: { CXO: [{ visualName: 'v', values: ['y'] }] } };
    expect(() => validateScenarios([a, b])).toThrow(/duplicates the name/);
  });

  test('rejects two different names that collapse to the same output slug', () => {
    // "CXO Sales" and "CXO/Sales" both slugify to "CXO_Sales" — they would
    // share an output folder and silently overwrite each other mid-run.
    const a: SlicerScenario = { name: 'CXO Sales', pages: { CXO: [{ visualName: 'v', values: ['x'] }] } };
    const b: SlicerScenario = { name: 'CXO/Sales', pages: { CXO: [{ visualName: 'v', values: ['y'] }] } };
    expect(() => validateScenarios([a, b])).toThrow(/same output folder/);
  });
});

// ── scenarioSlug ──────────────────────────────────────────────────────────────

test.describe('scenarioSlug', () => {
  test('keeps safe characters and replaces the rest with underscore', () => {
    expect(scenarioSlug('CXO Sales')).toBe('CXO_Sales');
    expect(scenarioSlug('CXO/Sales')).toBe('CXO_Sales');
    expect(scenarioSlug('a.b-c_d')).toBe('a.b-c_d');
  });

  test('trims leading/trailing underscores produced by stripping', () => {
    expect(scenarioSlug('  spaced  ')).toBe('spaced');
    expect(scenarioSlug('***weird***')).toBe('weird');
  });

  test('falls back to "scenario" when nothing usable remains', () => {
    expect(scenarioSlug('///')).toBe('scenario');
    expect(scenarioSlug('')).toBe('scenario');
  });
});

// ── getForPageCI / selectionsForPage ──────────────────────────────────────────

test.describe('getForPageCI', () => {
  test('finds a key regardless of case', () => {
    const rec = { CXO: [1, 2], Manager: [3] };
    expect(getForPageCI(rec, 'cxo')).toEqual([1, 2]);
    expect(getForPageCI(rec, 'MANAGER')).toEqual([3]);
  });

  test('returns [] for a page not present', () => {
    expect(getForPageCI({ CXO: [1] }, 'Employee')).toEqual([]);
  });

  test('selectionsForPage does the same case-insensitive lookup into a scenario', () => {
    const sc: SlicerScenario = {
      name: 's', pages: { CXO: [{ visualName: 'v', values: ['x'] }] },
    };
    expect(selectionsForPage(sc, 'cxo')).toHaveLength(1);
    expect(selectionsForPage(sc, 'nope')).toEqual([]);
  });
});

// ── planFilterApplication (the global-batch split) ────────────────────────────

test.describe('planFilterApplication', () => {
  const flat = (table: string, column: string, values: string[], title?: string) => ({
    title, values, targets: [{ table, column }],
  });

  test('batches a flat field with identical values on 2+ pages into one global group', () => {
    const sc: SlicerScenario = {
      name: 's',
      pages: {
        CXO:     [flat('Emp', 'Status', ['Active'])],
        Manager: [flat('Emp', 'Status', ['Active'])],
      },
    };
    const plan = planFilterApplication(sc);
    expect(plan.global).toHaveLength(1);
    expect(plan.global[0].values).toEqual(['Active']);
    expect(plan.global[0].pages.sort()).toEqual(['CXO', 'Manager']);
    // Both pages' per-page remainder should now be empty for that field.
    expect(plan.perPage['CXO']).toEqual([]);
    expect(plan.perPage['Manager']).toEqual([]);
  });

  test('does NOT batch a single-page field (no benefit — stays per-page)', () => {
    const sc: SlicerScenario = {
      name: 's', pages: { CXO: [flat('Emp', 'Status', ['Active'])] },
    };
    const plan = planFilterApplication(sc);
    expect(plan.global).toHaveLength(0);
    expect(plan.perPage['CXO']).toHaveLength(1);
  });

  test('does NOT batch the same field when the values differ per page', () => {
    const sc: SlicerScenario = {
      name: 's',
      pages: {
        CXO:     [flat('Emp', 'Status', ['Active'])],
        Manager: [flat('Emp', 'Status', ['Inactive'])],
      },
    };
    const plan = planFilterApplication(sc);
    expect(plan.global).toHaveLength(0);
    expect(plan.perPage['CXO']).toHaveLength(1);
    expect(plan.perPage['Manager']).toHaveLength(1);
  });

  test('never batches a hierarchy selection even if it repeats across pages', () => {
    const hier = { visualName: 'v', values: ['BU1'], isHierarchy: true, targets: [{ table: 'T', column: 'BU' }] };
    const sc: SlicerScenario = { name: 's', pages: { CXO: [hier], Manager: [hier] } };
    const plan = planFilterApplication(sc);
    expect(plan.global).toHaveLength(0);
    expect(plan.perPage['CXO']).toHaveLength(1);
    expect(plan.perPage['Manager']).toHaveLength(1);
  });

  test('never batches a selection with no targets (hand-written, visual-only)', () => {
    const sc: SlicerScenario = {
      name: 's',
      pages: {
        CXO:     [{ visualName: 'v', values: ['x'] }],
        Manager: [{ visualName: 'v', values: ['x'] }],
      },
    };
    const plan = planFilterApplication(sc);
    expect(plan.global).toHaveLength(0);
    expect(plan.perPage['CXO']).toHaveLength(1);
    expect(plan.perPage['Manager']).toHaveLength(1);
  });

  test('is order-independent on values when keying (["A","B"] === ["B","A"])', () => {
    const sc: SlicerScenario = {
      name: 's',
      pages: {
        CXO:     [flat('Emp', 'Dept', ['A', 'B'])],
        Manager: [flat('Emp', 'Dept', ['B', 'A'])],
      },
    };
    const plan = planFilterApplication(sc);
    expect(plan.global).toHaveLength(1);
    expect(plan.global[0].pages.sort()).toEqual(['CXO', 'Manager']);
  });

  test('moves work only FROM per-page TO global, never drops a selection', () => {
    const sc: SlicerScenario = {
      name: 's',
      pages: {
        CXO:     [flat('Emp', 'Status', ['Active']), { visualName: 'v', values: ['BU1'] }],
        Manager: [flat('Emp', 'Status', ['Active'])],
      },
    };
    const plan = planFilterApplication(sc);
    // Status batched globally; the visual-only selection on CXO stays per-page.
    expect(plan.global).toHaveLength(1);
    expect(plan.perPage['CXO']).toHaveLength(1);
    expect(plan.perPage['CXO'][0].visualName).toBe('v');
    expect(plan.perPage['Manager']).toEqual([]);
  });
});
