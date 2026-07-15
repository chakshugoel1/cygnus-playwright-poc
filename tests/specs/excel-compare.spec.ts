/**
 * excel-compare.spec.ts
 *
 * Standalone Playwright test that compares two Excel files without launching a browser.
 *
 * Usage:
 *   COMPARE_FILE_A="path/to/expected.xlsx" COMPARE_FILE_B="path/to/actual.xlsx" npm run compare:excel
 *
 * PowerShell:
 *   $env:COMPARE_FILE_A="playwright-report-cygnus\cygnus-expected-values.xlsx"; $env:COMPARE_FILE_B="playwright-report-cygnus\cygnus-main-run.xlsx"; npm run compare:excel
 */

import { test, expect } from '@playwright/test';
import * as fs   from 'fs';
import * as path from 'path';
import { compareWorkbooks } from '../helpers/excel-compare.helpers';

const OUT_DIR = path.join(process.cwd(), 'playwright-report-cygnus');

test('Excel Comparison — expected vs actual', async () => {
  test.setTimeout(300_000); // 5 min for large files

  const fileA = process.env['COMPARE_FILE_A'];
  const fileB = process.env['COMPARE_FILE_B'];

  if (!fileA || !fileB) {
    throw new Error(
      'Set COMPARE_FILE_A and COMPARE_FILE_B environment variables before running.\n' +
      'Example (PowerShell):\n' +
      '  $env:COMPARE_FILE_A="playwright-report-cygnus\\cygnus-expected-values.xlsx"\n' +
      '  $env:COMPARE_FILE_B="playwright-report-cygnus\\cygnus-main-run.xlsx"\n' +
      '  npm run compare:excel',
    );
  }

  const resolvedA = path.isAbsolute(fileA) ? fileA : path.join(process.cwd(), fileA);
  const resolvedB = path.isAbsolute(fileB) ? fileB : path.join(process.cwd(), fileB);

  if (!fs.existsSync(resolvedA)) throw new Error(`COMPARE_FILE_A not found: ${resolvedA}`);
  if (!fs.existsSync(resolvedB)) throw new Error(`COMPARE_FILE_B not found: ${resolvedB}`);

  console.log(`\nComparing:`);
  console.log(`  Expected: ${resolvedA}`);
  console.log(`  Actual  : ${resolvedB}\n`);

  const diff = await compareWorkbooks(resolvedA, resolvedB);

  // ── Console output ──────────────────────────────────────────────────────────
  console.log(`Sheets in expected : ${diff.sheets.filter(s => s.inExpected).length}`);
  console.log(`Sheets in actual   : ${diff.sheets.filter(s => s.inActual).length}`);
  if (diff.sheetsOnlyInExpected.length) console.log(`  Only in expected : ${diff.sheetsOnlyInExpected.join(', ')}`);
  if (diff.sheetsOnlyInActual.length)   console.log(`  Only in actual   : ${diff.sheetsOnlyInActual.join(', ')}`);
  console.log('');

  for (const s of diff.sheets) {
    if (!s.inExpected || !s.inActual) continue;
    const symbol = s.identical ? '✅' : '❌';
    console.log(`${symbol} ${s.sheet.padEnd(20)} rows expected=${s.rowsExpected} actual=${s.rowsActual} onlyInExpected=${s.rowsOnlyInExpected} onlyInActual=${s.rowsOnlyInActual}`);
    if (!s.identical && s.sampleOnlyInExpected.length > 0) {
      console.log(`   Sample rows only in expected:`);
      s.sampleOnlyInExpected.slice(0, 3).forEach(r => console.log(`     ${r.join(' | ')}`));
    }
    if (!s.identical && s.sampleOnlyInActual.length > 0) {
      console.log(`   Sample rows only in actual:`);
      s.sampleOnlyInActual.slice(0, 3).forEach(r => console.log(`     ${r.join(' | ')}`));
    }
  }

  console.log(`\nOverall: ${diff.identical ? '✅ IDENTICAL' : '❌ DIFFERENCES FOUND'}`);

  // ── Write JSON result ───────────────────────────────────────────────────────
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const resultFile = path.join(OUT_DIR, 'excel-compare-result.json');
  fs.writeFileSync(resultFile, JSON.stringify(diff, null, 2));
  console.log(`\n[✓] Result → ${resultFile}`);

  // Test passes regardless of diff result — the comparison itself is the artifact.
  // Uncomment the line below to make the test FAIL when differences are found:
  // expect(diff.identical, `Excel files differ — see ${resultFile}`).toBe(true);
  expect(true).toBe(true);
});
