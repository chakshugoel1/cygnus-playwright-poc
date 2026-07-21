import { test, expect } from '@playwright/test';
import { computeAnalysisStats, buildAnalysisNarrative, summarizeForUi } from '../helpers/report-export.helpers';
import type { WorkbookDiffSummary, SheetDiffSummary, VisualBlockDiff } from '../helpers/excel-compare.helpers';

function visual(status: VisualBlockDiff['status'], title = 'V'): VisualBlockDiff {
  return {
    title, type: 'card', status,
    countExpected: 1, countActual: 1,
    rowsExpected: 1, rowsActual: 1,
    headerExpected: ['H'], headerActual: ['H'],
    sampleOnlyInExpected: [], sampleOnlyInActual: [],
    onlyInExpectedCount: 0, onlyInActualCount: 0,
  };
}

function sheet(visuals: VisualBlockDiff[], opts: Partial<SheetDiffSummary> = {}): SheetDiffSummary {
  return {
    sheet: opts.sheet ?? 'Page',
    inExpected: true, inActual: true,
    rowsExpected: 10, rowsActual: 10,
    identical: visuals.length === 0,
    dataIdentical: visuals.every(v => v.status === 'identical' || v.status === 'header-diff' || v.status === 'duplicate-count'),
    visuals,
    headerDiffCount: visuals.filter(v => v.status === 'header-diff').length,
    structuralDiffCount: visuals.filter(v => v.status === 'only-in-expected' || v.status === 'only-in-actual' || v.status === 'ambiguous').length,
    rowsOnlyInExpected: 0, rowsOnlyInActual: 0,
    sampleOnlyInExpected: [], sampleOnlyInActual: [],
    ...opts,
  };
}

function workbook(sheets: SheetDiffSummary[], opts: Partial<WorkbookDiffSummary> = {}): WorkbookDiffSummary {
  return {
    fileExpected: 'expected.xlsx', fileActual: 'actual.xlsx', runAt: new Date().toISOString(),
    sheetsOnlyInExpected: [], sheetsOnlyInActual: [],
    sheets,
    identical: sheets.every(s => s.identical),
    dataIdentical: sheets.every(s => s.dataIdentical),
    ...opts,
  };
}

test.describe('computeAnalysisStats', () => {
  test('a byte-identical (Tier 1 fast-pass) page counts as fastPassPages, not zero visuals', () => {
    const diff = workbook([sheet([], { identical: true })]);
    const stats = computeAnalysisStats(diff);
    expect(stats).toMatchObject({ totalVisuals: 0, fastPassPages: 1 });
  });

  test('buckets header-diff/duplicate-count as cosmetic, ambiguous as review, missing/data-diff as critical', () => {
    const diff = workbook([sheet([
      visual('identical'),
      visual('header-diff'),
      visual('duplicate-count'),
      visual('ambiguous'),
      visual('only-in-expected'),
      visual('only-in-actual'),
      visual('data-diff'),
    ])]);
    const stats = computeAnalysisStats(diff);
    expect(stats).toMatchObject({
      totalVisuals: 7, identicalVisuals: 1, cosmeticOnly: 2, needsReview: 1, critical: 3,
    });
  });

  test('a page missing entirely on one side is not counted in visual stats', () => {
    const diff = workbook([sheet([], { inActual: false, identical: false })]);
    const stats = computeAnalysisStats(diff);
    expect(stats).toMatchObject({ totalVisuals: 0, fastPassPages: 0 });
  });
});

test.describe('buildAnalysisNarrative', () => {
  test('all identical: one clean sentence, no noise', () => {
    const diff = workbook([sheet([visual('identical'), visual('identical')])]);
    const stats = computeAnalysisStats(diff);
    const lines = buildAnalysisNarrative(diff, stats, true, []);
    expect(lines).toHaveLength(1);
    expect(lines[0]).toMatch(/identical/i);
  });

  test('only cosmetic diffs: says nothing needs action', () => {
    const diff = workbook([sheet([visual('identical'), visual('header-diff'), visual('header-diff')])]);
    const stats = computeAnalysisStats(diff);
    const lines = buildAnalysisNarrative(diff, stats, true, []);
    expect(lines.join(' ')).toMatch(/cosmetic only/i);
    expect(lines.join(' ')).toMatch(/nothing here needs action/i);
  });

  test('mix of cosmetic and critical: reports counts and percentage, flags critical in bottom line', () => {
    const diff = workbook([sheet([
      visual('header-diff'), visual('header-diff'), visual('header-diff'),
      visual('only-in-expected'),
    ])]);
    const stats = computeAnalysisStats(diff);
    const lines = buildAnalysisNarrative(diff, stats, false, []);
    const text = lines.join(' ');
    expect(text).toMatch(/3 cosmetic label rename\(s\) \(75%/);
    expect(text).toMatch(/1 critical difference/);
    expect(text).toMatch(/Bottom line: 1 genuine issue/);
  });

  test('only ambiguous, no critical: bottom line says nothing critical but flags review', () => {
    const diff = workbook([sheet([visual('ambiguous')])]);
    const stats = computeAnalysisStats(diff);
    const lines = buildAnalysisNarrative(diff, stats, false, []);
    expect(lines.join(' ')).toMatch(/nothing critical, but 1 visual\(s\) need a quick manual check/);
  });

  test('caveats present: leads with a caution line before anything else', () => {
    const diff = workbook([sheet([visual('identical')])]);
    const stats = computeAnalysisStats(diff);
    const lines = buildAnalysisNarrative(diff, stats, false, ['filter field mismatch']);
    expect(lines[0]).toMatch(/NOT directly comparable/);
    expect(lines[0]).toMatch(/1 caveat/);
  });

  test('a page missing entirely on one side gets its own mention, separate from visual stats', () => {
    const diff = workbook(
      [sheet([visual('identical')])],
      { sheetsOnlyInExpected: ['Old Page'] },
    );
    const stats = computeAnalysisStats(diff);
    const lines = buildAnalysisNarrative(diff, stats, true, []);
    expect(lines.some(l => /1 entire page\(s\) exist on only one side/.test(l))).toBe(true);
  });
});

// ── summarizeForUi (the desktop app's in-page summary, NOT the xlsx) ─────────

test.describe('summarizeForUi', () => {
  test('all identical: verdict pass, no sample-row fields leak into the output', () => {
    const diff = workbook([sheet([visual('identical', 'KPI')])]);
    const ui = summarizeForUi(diff, [], true, 0);
    expect(ui.verdict).toBe('pass');
    expect(ui.passed).toBe(true);
    expect(ui.pages).toHaveLength(1);
    expect(ui.pages[0].visuals[0]).toEqual({
      title: 'KPI', type: 'card', severity: 'safe', statusLabel: expect.any(String), note: undefined,
    });
    // No sample-row arrays anywhere in the visual shape (would be a payload-size regression).
    expect(Object.keys(ui.pages[0].visuals[0])).not.toContain('sampleOnlyInExpected');
    expect(Object.keys(ui.pages[0].visuals[0])).not.toContain('sampleOnlyInActual');
  });

  test('real differences: verdict fail, per-page status matches the xlsx Page Comparison logic', () => {
    const diff = workbook([
      sheet([visual('data-diff')], { sheet: 'Different Page', identical: false, dataIdentical: false }),
      sheet([visual('header-diff')], { sheet: 'Header Only Page' }),
      sheet([visual('identical')], { sheet: 'Same Page', identical: true }),
    ]);
    const ui = summarizeForUi(diff, [], false, 1);
    expect(ui.verdict).toBe('fail');
    const byName = Object.fromEntries(ui.pages.map(p => [p.name, p.status]));
    expect(byName['Different Page']).toBe('different');
    expect(byName['Header Only Page']).toBe('header-only');
    expect(byName['Same Page']).toBe('identical');
  });

  test('a page missing on one side gets only-in-source/only-in-target status, not folded into pages[]', () => {
    const diff = workbook([
      sheet([], { sheet: 'Gone', inExpected: true, inActual: false, identical: false, dataIdentical: false }),
      sheet([], { sheet: 'New', inExpected: false, inActual: true, identical: false, dataIdentical: false }),
    ], { sheetsOnlyInExpected: ['Gone'], sheetsOnlyInActual: ['New'] });
    const ui = summarizeForUi(diff, [], false, 0);
    // These sheets are excluded from pages[] (no visuals to show) - they're
    // surfaced via pagesOnlyInSource/pagesOnlyInTarget instead.
    expect(ui.pages).toEqual([]);
    expect(ui.pagesOnlyInSource).toEqual(['Gone']);
    expect(ui.pagesOnlyInTarget).toEqual(['New']);
  });

  test('caveats present force verdict to not_comparable even when data otherwise passed', () => {
    const diff = workbook([sheet([visual('identical')])]);
    const ui = summarizeForUi(diff, ['filter field mismatch'], false, 0);
    expect(ui.verdict).toBe('not_comparable');
    expect(ui.caveats).toEqual(['filter field mismatch']);
  });

  test('narrative and stats match what buildAnalysisNarrative/computeAnalysisStats compute directly (single source of truth)', () => {
    const diff = workbook([sheet([visual('header-diff'), visual('only-in-expected')])]);
    const stats = computeAnalysisStats(diff);
    const narrative = buildAnalysisNarrative(diff, stats, false, []);
    const ui = summarizeForUi(diff, [], false, 1);
    expect(ui.stats).toEqual(stats);
    expect(ui.narrative).toEqual(narrative);
  });
});
