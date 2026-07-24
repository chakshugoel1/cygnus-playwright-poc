/**
 * CYGNUS MAIN RUN: Power BI Harness (AAD token + visual.exportData)
 *
 * Purpose:
 *   1. Navigate to app.powerbi.com (uses existing cygnus.user.json session)
 *   2. Extract user's AAD token from browser MSAL cache
 *   3. loadHarnessPage():
 *        a. Call GenerateToken REST API → Power BI embed token (TokenType.Embed = 0)
 *        b. Serve harness/harness.html via page.route at http://localhost:3001/
 *        c. Intercept globalservice cluster call — proxy from Node.js with embed token
 *        d. Navigate to http://localhost:3001/ (standalone page, no CSP/AMD issues)
 *        e. Call window.__startEmbed(embedToken, embedUrl) — wait for 'rendered'
 *   4. List all report pages/tabs
 *   5. For each tab, export all visual data via visual.exportData()
 *   6. Save full results to playwright-report-cygnus/cygnus-main-run.json
 *
 * Success criteria:
 *   - Report embeds and renders on the standalone harness page
 *   - At least 1 visual on at least 1 page returns data rows
 *
 * Run with:
 *   npm run main:run
 */

import { test, expect } from '@playwright/test';
import * as fs   from 'fs';
import * as path from 'path';
import ExcelJS    from 'exceljs';
import { compareWorkbooks } from '../helpers/excel-compare.helpers';
import { getConfiguredTargetPages } from '../helpers/poc-config.helpers';
import {
  CYGNUS_REPORT_URL,
  waitForReportLoad,
} from '../helpers/cygnus.helpers';
import {
  extractUserTokenFromBrowser,
} from '../helpers/pbi-api.helpers';
import {
  loadHarnessPage,
  getReportPages,
  exportCurrentPageVisuals,
  setReportPage,
  type ReportPage,
  type VisualExport,
} from '../helpers/harness.helpers';

const OUT_DIR  = path.join(process.cwd(), 'playwright-report-cygnus');
const OUT_FILE = path.join(OUT_DIR, 'cygnus-main-run.json');
const EXCEL_MAX_SHEET_NAME = 31;

function toExcelSafeSheetBase(name: string): string {
  // Excel sheet names cannot contain: : \\ / ? * [ ]
  const cleaned = name.replace(/[:\\/?*\[\]]/g, ' ').trim();
  return cleaned || 'Sheet';
}

// Truncates by Unicode CODE POINT, not raw string index - see the same
// helper in report-export.helpers.ts for why plain .slice() risks splitting
// a surrogate pair (emoji, other astral-plane characters) in half.
function truncateSafe(text: string, maxLen: number): string {
  return Array.from(text).slice(0, maxLen).join('');
}

function toUniqueSheetName(name: string, usedNames: Set<string>): string {
  const base = toExcelSafeSheetBase(name);
  let candidate = truncateSafe(base, EXCEL_MAX_SHEET_NAME);
  let counter = 2;

  while (usedNames.has(candidate.toLowerCase())) {
    const suffix = `_${counter}`;
    const maxBaseLen = EXCEL_MAX_SHEET_NAME - suffix.length;
    candidate = `${truncateSafe(base, Math.max(1, maxBaseLen))}${suffix}`;
    counter += 1;
  }

  usedNames.add(candidate.toLowerCase());
  return candidate;
}

// Keep the browser open on failure for a short window so the DevTools console
// can be inspected, then let the test runner exit cleanly so the HTML report
// can be generated and opened with `npm run show:report`.
test.afterEach(async ({ page }, testInfo) => {
  if (testInfo.status !== testInfo.expectedStatus) {
    const PAUSE_SECONDS = 30;
    console.log(`\n[!] Test failed — browser stays open for ${PAUSE_SECONDS}s.`);
    console.log(`    Inspect DevTools now, then the browser will close automatically.`);
    console.log(`    After it closes, run: npm run show:report`);
    try {
      await page.waitForTimeout(PAUSE_SECONDS * 1000);
    } catch {
      // Page may already be closed by the time afterEach executes.
    }
  }
});

test('Cygnus Main Run — Power BI Harness visual.exportData()', async ({ page }) => {
  test.setTimeout(9_200_000); // 20 min — embed + export can be slow

  fs.mkdirSync(OUT_DIR, { recursive: true });

  interface PageResult {
    pageName:       string;
    pageDisplayName: string;
    visuals:        VisualExport[];
    visualsWithData: number;
    totalRows:      number;
  }

  const report: {
    runAt:       string;
    tokenFound:  boolean;
    embedOk:     boolean;
    pages:       ReportPage[];
    pageResults: PageResult[];
    conclusions: string[];
  } = {
    runAt:       new Date().toISOString(),
    tokenFound:  false,
    embedOk:     false,
    pages:       [],
    pageResults: [],
    conclusions: [],
  };

  // ── Step 1: Get user token from app.powerbi.com ───────────────────────────
  console.log('\n[1] Navigating to app.powerbi.com to acquire user token...');

  // PRIMARY: intercept Bearer tokens straight from Power BI network requests.
  // When the report loads, the browser immediately fires authenticated API calls
  // whose Authorization headers contain the live token — no MSAL storage write needed.
  let interceptedToken: string | null = null;
  page.on('request', (request) => {
    if (interceptedToken) return;
    const auth = request.headers()['authorization'];
    if (!auth?.startsWith('Bearer ')) return;
    const url = request.url();
    if (
      url.includes('powerbi.com') ||
      url.includes('analysis.windows.net') ||
      url.includes('pbidedicated.windows.net')
    ) {
      const t = auth.slice(7); // strip 'Bearer '
      if (t.length > 200) {   // genuine JWT, not a short placeholder
        interceptedToken = t;
      }
    }
  });

  await page.goto(CYGNUS_REPORT_URL);
  await waitForReportLoad(page);

  // Give MSAL storage extraction up to 15 s (validated audience in helper),
  // and only then fallback to the raw intercepted bearer token.
  let token: string | null = null;
  const tokenPollStart = Date.now();
  while (!token && Date.now() - tokenPollStart < 15_000) {
    token = await extractUserTokenFromBrowser(page);
    if (!token) {
      console.log('    ⏳ Token not yet captured — waiting 2 s...');
      await page.waitForTimeout(2000);
    }
  }
  if (!token) token = interceptedToken;
  console.log(token
    ? `    ✅ Token acquired after ${Math.round((Date.now() - tokenPollStart) / 1000)}s (length: ${token.length})`
    : '    ❌ Token not found after 15s');

  if (!token) {
    report.conclusions.push('❌ User token not found — run diag:option-e first to confirm token extraction works');
    fs.writeFileSync(OUT_FILE, JSON.stringify(report, null, 2));
    expect(false, 'Token not found').toBe(true);
    return;
  }
  report.tokenFound = true;

  // ── Step 2: Load harness page ─────────────────────────────────────────────
  console.log('\n[2] Loading harness page and embedding report...');
  try {
    await loadHarnessPage(page, token);
    report.embedOk = true;
  } catch (e: any) {
    report.conclusions.push(`❌ Embed failed: ${e.message}`);
    report.conclusions.push('Check: does the user have access to this report?');
    report.conclusions.push('Check: is Report.Read.All Delegated permission admin-consented?');
    fs.writeFileSync(OUT_FILE, JSON.stringify(report, null, 2));
    expect(false, `Embed failed: ${e.message}`).toBe(true);
    return;
  }

  // ── Step 3: Discover report pages ─────────────────────────────────────────
  console.log('\n[3] Discovering report pages...');
  report.pages = await getReportPages(page);
  console.log(`    Found ${report.pages.length} pages total.`);

  // ── Step 4: Export visuals from the 4 profile tabs (exact page names) ───────
  const DEFAULT_TARGET_PAGES = [
    { name: 'ReportSection',                     displayName: 'Employee' },
    { name: 'ReportSection044c6f061497c80b84d6', displayName: 'Manager'  },
    { name: 'ReportSection1c1b468f352e3b319033', displayName: 'DU Head'  },
    { name: 'ReportSection4d610368ccfea4f3d615', displayName: 'CXO'      },
  ];
  const TARGET_PAGES = getConfiguredTargetPages(DEFAULT_TARGET_PAGES);

  // If none of the configured target pages exist in this report, fall back to all discovered pages
  const matchedCount = TARGET_PAGES.filter(t => report.pages.some(p => p.name === t.name)).length;
  const effectiveTargets = matchedCount > 0 ? TARGET_PAGES : report.pages;

  console.log(`\n[4] Exporting visuals from ${effectiveTargets.length} profile tabs...`);

  for (const target of effectiveTargets) {
    const reportPage = report.pages.find(p => p.name === target.name);
    if (!reportPage) {
      console.log(`      ⚠ Page not found: "${target.displayName}" (${target.name})`);
      continue;
    }
    console.log(`\n    → Tab: "${reportPage.displayName}"`);

    try {
      await setReportPage(page, reportPage.name);
    } catch (e: any) {
      console.log(`      ⚠ Could not navigate to page: ${e.message}`);
      continue;
    }

    const visuals = await exportCurrentPageVisuals(page);
    const dataHits = visuals.filter(v => v.rowCount > 0);
    const totalRows = visuals.reduce((sum, v) => sum + v.rowCount, 0);

    console.log(`      ${visuals.length} visuals — ${dataHits.length} with data — ${totalRows} total rows`);
    visuals.forEach(v => {
      const status = v.rowCount > 0 ? '✅' : v.errorMessage ? '⚠' : '○';
      console.log(`        ${status} "${v.visualTitle}" [${v.visualType}] — ${v.rowCount} rows`);
      if (v.headers.length > 0) {
        console.log(`           Columns: ${v.headers.slice(0, 4).join(', ')}${v.headers.length > 4 ? '...' : ''}`);
      }
    });

    report.pageResults.push({
      pageName:        reportPage.name,
      pageDisplayName: reportPage.displayName,
      visuals,
      visualsWithData: dataHits.length,
      totalRows,
    });
  }

  // ── Step 5: Conclusions ───────────────────────────────────────────────────
  const allDataHits  = report.pageResults.reduce((sum, p) => sum + p.visualsWithData, 0);
  const allTotalRows = report.pageResults.reduce((sum, p) => sum + p.totalRows, 0);

  report.conclusions.push(
    allDataHits > 0
      ? `✅ ${allDataHits} visuals returned data across ${report.pageResults.length} tabs (${allTotalRows} total rows)`
      : '❌ Report embedded but no visuals returned data — check exportData() permissions',
  );

  // ── Step 6: Save JSON + Excel ─────────────────────────────────────────────
  fs.writeFileSync(OUT_FILE, JSON.stringify(report, null, 2));
  console.log(`\n[✓] JSON → ${OUT_FILE}`);

  // Excel: one sheet per tab + a Summary sheet.
  // Layout-only visual types (images, shapes, buttons, text boxes) are skipped
  // since they carry no tabular data — only data visuals are written.
  const SKIP_TYPES = new Set(['image', 'shape', 'textbox', 'actionButton']);

  const xlsxFile = path.join(OUT_DIR, 'cygnus-main-run.xlsx');
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Cygnus Playwright POC';
  wb.created = new Date();

  const HDR_FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
  const SEC_FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
  const ERR_FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  const EVEN_BG:  ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
  const ODD_BG:   ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
  const BD: ExcelJS.Border = { style: 'thin', color: { argb: 'FF000000' } };
  const CB: Partial<ExcelJS.Borders> = { top: BD, bottom: BD, left: BD, right: BD };

  // ── Summary sheet ─────────────────────────────────────────────────────────
  const sumWs = wb.addWorksheet('Summary');
  sumWs.columns = [
    { header: 'Tab',          key: 'tab',    width: 14 },
    { header: 'Visual Title', key: 'title',  width: 34 },
    { header: 'Type',         key: 'type',   width: 22 },
    { header: 'Rows',         key: 'rows',   width: 8  },
    { header: 'Status',       key: 'status', width: 14 },
    { header: 'Error',        key: 'error',  width: 50 },
  ];
  sumWs.getRow(1).eachCell(c => {
    c.fill = HDR_FILL; c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
    c.border = CB; c.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  for (const pr of report.pageResults) {
    for (const v of pr.visuals) {
      if (SKIP_TYPES.has(v.visualType)) continue;
      const status = v.rowCount > 0 ? 'Data ✓' : v.errorMessage ? 'No export' : 'Empty';
      const fill   = v.rowCount > 0 ? EVEN_BG : v.errorMessage ? ERR_FILL : ODD_BG;
      const r = sumWs.addRow({ tab: pr.pageDisplayName, title: v.visualTitle || v.visualName, type: v.visualType, rows: v.rowCount, status, error: v.errorMessage ?? '' });
      r.eachCell(c => { c.fill = fill; c.border = CB; c.alignment = { vertical: 'middle', wrapText: false }; });
    }
  }

  // ── One sheet per tab ─────────────────────────────────────────────────────
  const usedSheetNames = new Set<string>(['summary']);
  for (const pr of report.pageResults) {
    const wsName      = toUniqueSheetName(pr.pageDisplayName, usedSheetNames);
    const ws          = wb.addWorksheet(wsName);
    const dataVisuals = pr.visuals.filter(v => !SKIP_TYPES.has(v.visualType) && v.rowCount > 0 && v.headers.length > 0);

    if (dataVisuals.length === 0) {
      ws.addRow([`No data visuals found for the ${pr.pageDisplayName} tab`]);
      continue;
    }

    for (const v of dataVisuals) {
      const nc = v.headers.length;

      // Section title: visual title + type + row count
      const secRow = ws.addRow([`${v.visualTitle || v.visualName}  [${v.visualType}]  —  ${v.rowCount} rows`]);
      if (nc > 1) ws.mergeCells(secRow.number, 1, secRow.number, nc);
      secRow.getCell(1).fill      = SEC_FILL;
      secRow.getCell(1).font      = { bold: true, size: 11 };
      secRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
      secRow.getCell(1).border    = CB;
      secRow.height = 22;

      // Column headers from visual.exportData() CSV
      const hdrRow = ws.addRow(v.headers);
      hdrRow.eachCell((c, col) => {
        c.fill = HDR_FILL;
        c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
        c.border = CB;
        c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        const colObj = ws.getColumn(col);
        const maxVal = Math.max(...v.rows.slice(0, 20).map(r => String(r[v.headers[col - 1]] ?? '').length));
        colObj.width = Math.max(colObj.width ?? 10, Math.min(Math.max(v.headers[col - 1].length, maxVal) + 4, 45));
      });

      // Data rows
      v.rows.forEach((row, idx) => {
        const vals = v.headers.map(h => row[h] ?? '');
        const dr   = ws.addRow(vals);
        dr.eachCell(c => {
          c.fill   = idx % 2 === 0 ? EVEN_BG : ODD_BG;
          c.border = CB;
          c.alignment = { vertical: 'middle' };
        });
      });

      ws.addRow([]); // spacer between visuals
    }
  }

  await wb.xlsx.writeFile(xlsxFile);
  console.log(`[✓] Excel → ${xlsxFile}`);
  console.log(`    Sheets: ${report.pageResults.map(p => p.pageDisplayName).join(' | ')}`);

  // ── Step 7: Excel-to-Excel comparison against expected baseline ─────────────
  const EXPECTED_XLSX     = path.join(OUT_DIR, 'cygnus-expected-values.xlsx');
  const EXEC_SUMMARY_FILE = path.join(OUT_DIR, 'cygnus-execution-summary.xlsx');

  let diff: Awaited<ReturnType<typeof compareWorkbooks>> | null = null;

  if (fs.existsSync(EXPECTED_XLSX)) {
    console.log('\n[7] Comparing actual vs expected Excel (streaming two-tier)...');
    diff = await compareWorkbooks(EXPECTED_XLSX, xlsxFile);

    const diffSheets = diff.sheets.filter(s => !s.identical);
    console.log(`    Sheets compared : ${diff.sheets.length}`);
    console.log(`    Identical sheets: ${diff.sheets.filter(s => s.identical).length}`);
    console.log(`    Sheets with diffs: ${diffSheets.length}`);
    diffSheets.forEach(s => {
      console.log(`      ✗ ${s.sheet}: +${s.rowsOnlyInActual} rows in actual, -${s.rowsOnlyInExpected} rows only in expected`);
    });

    report.conclusions.push(
      diff.identical
        ? '✅ Excel comparison: actual matches expected exactly'
        : `⚠️ Excel comparison: ${diffSheets.length} sheet(s) differ — see cygnus-execution-summary.xlsx`,
    );
  } else {
    // No baseline yet — copy current run as the expected values file
    console.log('\n[7] No expected values Excel found — saving current run as baseline...');
    fs.copyFileSync(xlsxFile, EXPECTED_XLSX);
    console.log(`    ✅ Baseline saved → ${EXPECTED_XLSX}`);
    report.conclusions.push('✅ Expected values file created from this run. Next run will compare against it.');
  }

  // ── Step 8: Generate Execution Summary Excel ──────────────────────────────
  console.log('\n[8] Generating Execution Summary Excel...');
  const execWb = new ExcelJS.Workbook();
  execWb.creator = 'Cygnus Playwright POC';
  execWb.created = new Date();

  const E_PASS_FILL:  ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } };
  const E_FAIL_FILL:  ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
  const E_HDR_FILL:   ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
  const E_INFO_FILL:  ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9EFF7' } };
  const E_WARN_FILL:  ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } };
  const E_MISS_FILL:  ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  const E_BD:  ExcelJS.Border          = { style: 'thin', color: { argb: 'FFCCCCCC' } };
  const E_CB: Partial<ExcelJS.Borders> = { top: E_BD, bottom: E_BD, left: E_BD, right: E_BD };

  // Sheet 1: Run Info
  const infoWs = execWb.addWorksheet('Run Info');
  infoWs.columns = [{ width: 28 }, { width: 55 }];
  const diffSheetCount = diff ? diff.sheets.filter(s => !s.identical).length : 0;
  const overallResult  = !diff
    ? '⚠️ No baseline — created this run'
    : diff.identical ? '✅ PASSED — exact match' : `❌ FAILED — ${diffSheetCount} sheet(s) differ`;
  const infoData: [string, string | number][] = [
    // Pinned locale + named month - see report-export.helpers.ts's same fix
    // for why a bare .toLocaleString() (OS-locale-dependent, dd/mm vs mm/dd
    // is genuinely ambiguous) isn't used here.
    ['Run Time',          new Date(report.runAt).toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'medium' })],
    ['Embed OK',          report.embedOk   ? '✅ Yes' : '❌ No'],
    ['Token Found',       report.tokenFound ? '✅ Yes' : '❌ No'],
    ['Tabs Processed',    report.pageResults.length],
    ['Expected File',     diff ? diff.fileExpected : 'Created this run'],
    ['Actual File',       diff ? diff.fileActual   : xlsxFile],
    ['Sheets Compared',   diff ? diff.sheets.length : 0],
    ['Identical Sheets',  diff ? diff.sheets.filter(s => s.identical).length : 0],
    ['Sheets With Diffs', diffSheetCount],
    ['Overall Result',    overallResult],
    ['Conclusions',       report.conclusions.join(' | ')],
  ];
  infoData.forEach(([label, value]) => {
    const r = infoWs.addRow([label, String(value)]);
    r.getCell(1).font = { bold: true };
    r.getCell(1).fill = E_INFO_FILL;
    r.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
    r.height = 20;
  });

  // Sheet 2: Test Execution Summary
  if (diff) {
    const sheetCmpWs = execWb.addWorksheet('Test Execution Summary');
    sheetCmpWs.columns = [
      { header: 'Sheet',                   key: 'sheet',          width: 20 },
      { header: 'Rows in Expected',        key: 'rowsExpected',   width: 18 },
      { header: 'Rows in Actual',          key: 'rowsActual',     width: 16 },
      { header: 'Only in Expected',        key: 'onlyInExpected', width: 18 },
      { header: 'Only in Actual',          key: 'onlyInActual',   width: 16 },
      { header: 'Status',                  key: 'status',         width: 14 },
    ];
    sheetCmpWs.getRow(1).eachCell(c => {
      c.fill = E_HDR_FILL; c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      c.border = E_CB; c.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    sheetCmpWs.getRow(1).height = 22;

    for (const s of diff.sheets) {
      const statusText = s.identical ? '✅ Identical' : '❌ Different';
      const statusFill = s.identical ? E_PASS_FILL : E_FAIL_FILL;
      const row = sheetCmpWs.addRow({
        sheet: s.sheet,
        rowsExpected:   s.rowsExpected,
        rowsActual:     s.rowsActual,
        onlyInExpected: s.rowsOnlyInExpected,
        onlyInActual:   s.rowsOnlyInActual,
        status: statusText,
      });
      row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle' }; });
      const sc = row.getCell('status');
      sc.fill = statusFill;
      sc.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      sc.alignment = { horizontal: 'center', vertical: 'middle' };
    }

    // Add any sheets only in one file
    for (const s of diff.sheetsOnlyInExpected) {
      const row = sheetCmpWs.addRow({ sheet: s, rowsExpected: '-', rowsActual: 0, onlyInExpected: '-', onlyInActual: '-', status: '⚠️ Missing from actual' });
      row.getCell('status').fill = E_WARN_FILL;
      row.getCell('status').font = { bold: true };
      row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle' }; });
    }
    for (const s of diff.sheetsOnlyInActual) {
      const row = sheetCmpWs.addRow({ sheet: s, rowsExpected: 0, rowsActual: '-', onlyInExpected: '-', onlyInActual: '-', status: '🆕 New in actual' });
      row.getCell('status').fill = E_WARN_FILL;
      row.getCell('status').font = { bold: true };
      row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle' }; });
    }
  }

  // Sheet 3: Failed Logs
  const diffSheets = diff ? diff.sheets.filter(s => !s.identical && (s.sampleOnlyInExpected.length > 0 || s.sampleOnlyInActual.length > 0)) : [];
  if (diffSheets.length > 0) {
    const sampleWs = execWb.addWorksheet('Failed Logs');
    sampleWs.columns = [
      { header: 'Sheet',      key: 'sheet',  width: 18 },
      { header: 'Side',       key: 'side',   width: 22 },
      { header: 'Sample Row', key: 'row',    width: 90 },
    ];
    sampleWs.getRow(1).eachCell(c => {
      c.fill = E_HDR_FILL; c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      c.border = E_CB; c.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    sampleWs.getRow(1).height = 22;

    for (const s of diffSheets) {
      for (const r of s.sampleOnlyInExpected) {
        const row = sampleWs.addRow({ sheet: s.sheet, side: 'Only in Expected', row: r.join(' | ') });
        row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
        row.getCell('side').fill = E_MISS_FILL;
        row.getCell('side').font = { bold: true };
      }
      for (const r of s.sampleOnlyInActual) {
        const row = sampleWs.addRow({ sheet: s.sheet, side: 'Only in Actual', row: r.join(' | ') });
        row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
        row.getCell('side').fill = E_PASS_FILL;
        row.getCell('side').font = { bold: true, color: { argb: 'FFFFFFFF' } };
      }
    }
  }

  await execWb.xlsx.writeFile(EXEC_SUMMARY_FILE);
  console.log(`[✓] Execution Summary → ${EXEC_SUMMARY_FILE}`);

  // Update JSON with final conclusions
  fs.writeFileSync(OUT_FILE, JSON.stringify(report, null, 2));

  console.log('\n=== CONCLUSIONS ===');
  report.conclusions.forEach(c => console.log(c));

  expect(allDataHits, `No visual data returned. Check ${OUT_FILE}`).toBeGreaterThan(0);
});
