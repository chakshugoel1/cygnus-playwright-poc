/**
 * excel-compare.unit.spec.ts
 *
 * Offline end-to-end regression tests for compareWorkbooks: writes REAL xlsx
 * files with the same streaming writer the export pipeline uses (worksheets
 * placed before workbook.xml in the zip — the exact layout that broke the
 * ExcelJS streaming reader in production), then compares them. No browser, no
 * Power BI.
 *
 * Guards against the production incident where read-back sheet names came out
 * as "Sheet1/Sheet2/Sheet3" (or the reader crashed with "Cannot read
 * properties of undefined (reading 'sheets')"), which destroyed sheet pairing
 * and produced a garbage comparison. Run with: npm run test:unit
 */

import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import ExcelJS from 'exceljs';
import { compareWorkbooks, readWorkbookSheetInfo, normalizeRow } from '../helpers/excel-compare.helpers';
import { writeVisualsToSheet, diffPageSets } from '../helpers/report-export.helpers';
import type { VisualExport } from '../helpers/harness.helpers';

const tmpFiles: string[] = [];

function tmpPath(name: string): string {
  const p = path.join(os.tmpdir(), `cygnus-unit-${process.pid}-${name}`);
  tmpFiles.push(p);
  return p;
}

test.afterAll(() => {
  for (const f of tmpFiles) {
    try { fs.unlinkSync(f); } catch { /* already gone */ }
  }
});

function fakeVisual(title: string, rows: Array<[string, string]>): VisualExport {
  return {
    visualName:  `v-${title}`,
    visualTitle: title,
    visualType:  'tableEx',
    csvData:     null,
    headers:     ['Name', 'Value'],
    rows:        rows.map(([a, b]) => ({ Name: a, Value: b })),
    rowCount:    rows.length,
    errorMessage: null,
  };
}

/** Writes an xlsx exactly the way exportReportToWorkbook does: streaming
 *  writer, one committed sheet per page — worksheets land in the zip BEFORE
 *  workbook.xml, which is the trigger for the reader bug this file guards
 *  against. */
async function writeWorkbook(filePath: string, sheets: Array<{ name: string; visuals: VisualExport[] }>): Promise<void> {
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: filePath,
    useStyles: true,
    useSharedStrings: false,
  });
  for (const s of sheets) {
    const ws = wb.addWorksheet(s.name);
    writeVisualsToSheet(ws, s.visuals);
    ws.commit();
  }
  await wb.commit();
}

// German page names with umlauts + a name pair differing only by case, both
// taken from the real report pair that surfaced the bug.
const PAGE_A = 'Dashboard Bewerbungseingänge';
const PAGE_B = 'Dashboard Absagen';
const PAGE_C_SRC = 'Vertragsrücklauf - Cluster';
const PAGE_C_TGT = 'Vertragsrücklauf - cluster';
// Real production case: the exact same page lost its accent during
// migration ("Commodité" -> "Commodite") and was reported as two different
// pages instead of pairing up.
const PAGE_D_SRC = 'Commodit' + String.fromCharCode(0xe9); // "Commodité"
const PAGE_D_TGT = 'Commodite';

test.describe('compareWorkbooks on stream-written files', () => {
  test('reads back REAL page names, never Sheet1/Sheet2 defaults', async () => {
    const file = tmpPath('names.xlsx');
    await writeWorkbook(file, [
      { name: PAGE_A, visuals: [fakeVisual('KPI', [['a', '1']])] },
      { name: PAGE_B, visuals: [fakeVisual('KPI', [['b', '2']])] },
    ]);

    const infos = await readWorkbookSheetInfo(file);
    expect(infos.map(i => i.name)).toEqual([PAGE_A, PAGE_B]);
  });

  test('identical workbooks compare as identical under their real names', async () => {
    const sheets = [
      { name: PAGE_A, visuals: [fakeVisual('Eingänge nach Monat', [['Jan', '10'], ['Feb', '20']])] },
      { name: PAGE_B, visuals: [fakeVisual('Absagegründe', [['x', '5']])] },
    ];
    const fa = tmpPath('ident-a.xlsx');
    const fb = tmpPath('ident-b.xlsx');
    await writeWorkbook(fa, sheets);
    await writeWorkbook(fb, sheets);

    const diff = await compareWorkbooks(fa, fb);
    expect(diff.dataIdentical).toBe(true);
    expect(diff.sheets.map(s => s.sheet).sort()).toEqual([PAGE_A, PAGE_B].sort());
    expect(diff.sheetsOnlyInExpected).toEqual([]);
    expect(diff.sheetsOnlyInActual).toEqual([]);
  });

  test('pairs sheets whose names differ only by case, and flags a real value diff', async () => {
    const fa = tmpPath('case-a.xlsx');
    const fb = tmpPath('case-b.xlsx');
    await writeWorkbook(fa, [
      { name: PAGE_C_SRC, visuals: [fakeVisual('Cycle Time', [['2024', '12']])] },
    ]);
    await writeWorkbook(fb, [
      { name: PAGE_C_TGT, visuals: [fakeVisual('Cycle Time', [['2024', '99']])] }, // changed value
    ]);

    const diff = await compareWorkbooks(fa, fb);
    // Case-different names must PAIR (same page), not appear as missing sheets.
    expect(diff.sheetsOnlyInExpected).toEqual([]);
    expect(diff.sheetsOnlyInActual).toEqual([]);
    expect(diff.sheets).toHaveLength(1);
    expect(diff.sheets[0].dataIdentical).toBe(false); // and the value diff is real
  });

  test('a sheet genuinely on one side only is still reported as such', async () => {
    const fa = tmpPath('only-a.xlsx');
    const fb = tmpPath('only-b.xlsx');
    await writeWorkbook(fa, [
      { name: PAGE_A, visuals: [fakeVisual('KPI', [['a', '1']])] },
      { name: 'PAGES', visuals: [fakeVisual('Nav', [['n', '0']])] },
    ]);
    await writeWorkbook(fb, [
      { name: PAGE_A, visuals: [fakeVisual('KPI', [['a', '1']])] },
    ]);

    const diff = await compareWorkbooks(fa, fb);
    expect(diff.sheetsOnlyInExpected).toEqual(['PAGES']);
    expect(diff.dataIdentical).toBe(false);
  });

  test('pairs sheets whose names differ only by an accent (migration dropped the accent)', async () => {
    const fa = tmpPath('accent-a.xlsx');
    const fb = tmpPath('accent-b.xlsx');
    await writeWorkbook(fa, [
      { name: PAGE_D_SRC, visuals: [fakeVisual('KPI', [['a', '1']])] },
    ]);
    await writeWorkbook(fb, [
      { name: PAGE_D_TGT, visuals: [fakeVisual('KPI', [['a', '1']])] },
    ]);

    const diff = await compareWorkbooks(fa, fb);
    expect(diff.sheetsOnlyInExpected).toEqual([]);
    expect(diff.sheetsOnlyInActual).toEqual([]);
    expect(diff.sheets).toHaveLength(1);
    expect(diff.dataIdentical).toBe(true);
  });
});

test.describe('normalizeRow', () => {
  // Regression test for a production crash: ExcelJS's row.values can be a
  // SPARSE array (a row with no cell at all in some column leaves a real
  // hole, not an explicit undefined). Array.prototype.map skips holes, so
  // a naive `.slice(1).map(normalizeCell)` silently let a hole through as
  // `undefined` instead of '' - crashing the first caller that assumed
  // every cell was always a string (`cells[0].match(...)` in
  // streamSheetBlocks, on a 24k-row real table).
  test('a hole in a sparse row.values becomes "", never undefined', () => {
    const sparse: unknown[] = [];
    sparse[0] = undefined; // the always-unused 0-index placeholder
    sparse[2] = 'present'; // column B has a value; column A (index 1) is a genuine hole
    expect(1 in sparse).toBe(false); // confirms this really is a hole, not an explicit undefined

    const fakeRow = { values: sparse } as unknown as ExcelJS.Row;
    const cells = normalizeRow(fakeRow);

    expect(cells).toEqual(['', 'present']);
    expect(cells[0]).not.toBeUndefined();
  });

  test('a normal, fully-populated row is unaffected', () => {
    const fakeRow = { values: [undefined, 'a', 'b', 'c'] } as unknown as ExcelJS.Row;
    expect(normalizeRow(fakeRow)).toEqual(['a', 'b', 'c']);
  });
});

test.describe('diffPageSets', () => {
  test('trims and case-folds before matching (real drift seen in production)', () => {
    const d = diffPageSets(
      ['Vertragsversand', 'Vertragsrücklauf - Cluster', 'Absagen'],
      ['Vertragsversand ', 'Vertragsrücklauf - cluster', 'Absagen', 'PAGES', 'Page 4'],
    );
    expect(d.inBoth).toHaveLength(3);
    expect(d.onlyInSource).toEqual([]);
    expect(d.onlyInTarget).toEqual(['PAGES', 'Page 4']);
  });

  test('applies pageMap (source name -> target name) before matching', () => {
    const d = diffPageSets(['DU Head'], ['DU-Head'], { 'DU Head': 'DU-Head' });
    expect(d.inBoth).toEqual(['DU Head']);
    expect(d.onlyInSource).toEqual([]);
    expect(d.onlyInTarget).toEqual([]);
  });
});
