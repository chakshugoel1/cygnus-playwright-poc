/**
 * report-export.helpers.ts
 *
 * Report-agnostic export pipeline used by the report-parity (migration) flow.
 *
 *   exportReportToWorkbook()  embeds the CURRENTLY ACTIVE report (identity set
 *                             via applyReportIdentity), exports every visual on
 *                             every requested page, and writes a stable,
 *                             comparison-friendly .xlsx (one sheet per page).
 *
 *   writeParitySummary()      turns a WorkbookDiffSummary into a readable .xlsx
 *                             summary (overall verdict + per-page diffs + samples).
 *
 * The per-report workbooks contain ONLY data sheets (no run-info sheet) so the
 * two-tier row comparison in excel-compare.helpers stays clean.
 */

import type { Page } from '@playwright/test';
import * as fs   from 'fs';
import ExcelJS    from 'exceljs';
import {
  loadHarnessPage,
  getReportPages,
  setReportPage,
  exportCurrentPageVisuals,
  type ReportPage,
  type VisualExport,
} from './harness.helpers';
import type { WorkbookDiffSummary, VisualBlockDiff } from './excel-compare.helpers';

// Visual types that carry no tabular data — never written or compared.
const SKIP_TYPES = new Set(['image', 'shape', 'textbox', 'actionButton']);
const EXCEL_MAX_SHEET_NAME = 31;

// One retry, after a short pause, for the same class of transient hiccup
// readSlicerOptions() in harness.helpers.ts already retries around (observed
// in practice on busy pages) — cheap insurance against a single page silently
// looking like a genuine data difference between source and target.
const NAV_RETRY_WAIT_MS    = 1500;
const EXPORT_RETRY_WAIT_MS = 1500;

/**
 * Errors that cross the page.evaluate() boundary aren't always proper Error
 * instances with a usable .message — some Power BI SDK rejections are plain
 * objects that stringify down to "[object Object]" or just "Object" if you
 * assume .message exists. This pulls out whatever is actually readable.
 */
function describeError(e: unknown): string {
  if (e instanceof Error && e.message) return e.message;
  if (typeof e === 'string') return e;
  try {
    const s = JSON.stringify(e);
    if (s && s !== '{}') return s;
  } catch {
    // fall through
  }
  return String(e);
}

// ── Sheet-name helpers (Excel forbids : \ / ? * [ ] and 31-char names) ────────
function toExcelSafeSheetBase(name: string): string {
  const cleaned = name.replace(/[:\\/?*[\]]/g, ' ').trim();
  return cleaned || 'Sheet';
}

function toUniqueSheetName(name: string, usedNames: Set<string>): string {
  const base = toExcelSafeSheetBase(name);
  let candidate = base.slice(0, EXCEL_MAX_SHEET_NAME);
  let counter = 2;
  while (usedNames.has(candidate.toLowerCase())) {
    const suffix = `_${counter}`;
    const maxBaseLen = EXCEL_MAX_SHEET_NAME - suffix.length;
    candidate = `${base.slice(0, Math.max(1, maxBaseLen))}${suffix}`;
    counter += 1;
  }
  usedNames.add(candidate.toLowerCase());
  return candidate;
}

// ── Public result types ───────────────────────────────────────────────────────
export interface PageExportResult {
  sheetName:        string;
  pageDisplayName:  string;
  pageInternalName: string;
  visualsWithData:  number;
  totalRows:        number;
}

export interface ReportExportResult {
  pagesFound:  ReportPage[];
  exported:    PageExportResult[];
  outPath:     string;
}

export interface ExportOptions {
  /** Where to write the .xlsx. */
  outPath: string;
  /** If set, only export pages whose display name is in this list. */
  pagesFilter?: string[];
  /**
   * Map a page's actual display name → the sheet name to write it under.
   * Used to align renamed tabs between source and target. Default: identity.
   */
  sheetNameFor?: (pageDisplayName: string) => string;
  /** Label for console logging, e.g. "source (Import mode)". */
  label?: string;
  /**
   * Optional hook called ONCE, before any page is visited — for filters that
   * apply at the report level (affecting every page in one call), as opposed
   * to applySlicersForPage which runs once per page. If it throws, export
   * continues (a warning is logged) — the per-page hook is the safety net for
   * anything the global hook didn't (or couldn't) cover.
   */
  beforeExport?: (page: Page) => Promise<void>;
  /**
   * Optional hook called AFTER navigating to each page but BEFORE exporting its
   * visuals. Use it to apply slicer selections (filters) for that page. Receives
   * the Playwright page and the page's display name; should apply and settle any
   * filters, then resolve. If it throws, the page is still exported (unfiltered)
   * and a warning is logged, so one bad slicer can't abort the whole run.
   *
   * Kept as a generic callback so this module has no dependency on the slicer
   * scenario config — the spec wires the two together.
   */
  applySlicersForPage?: (page: Page, pageDisplayName: string) => Promise<void>;
}

// ── Styling (light — values are what matter, but readable output helps) ───────
const HDR_FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
const SEC_FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
const BD: ExcelJS.Border = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const CB: Partial<ExcelJS.Borders> = { top: BD, bottom: BD, left: BD, right: BD };

/**
 * Embeds the active report, exports all requested pages' visuals, and writes a
 * comparison-friendly workbook to options.outPath.
 *
 * The report identity MUST already be applied (applyReportIdentity) before this
 * call, and the user token acquired.
 */
export async function exportReportToWorkbook(
  page: Page,
  token: string,
  options: ExportOptions,
): Promise<ReportExportResult> {
  const label = options.label ?? 'report';
  const sheetNameFor = options.sheetNameFor ?? ((n: string) => n);

  // ── Embed ───────────────────────────────────────────────────────────────
  console.log(`\n[export:${label}] Embedding report...`);
  await loadHarnessPage(page, token);

  // ── Discover pages ──────────────────────────────────────────────────────
  const pagesFound = await getReportPages(page);
  console.log(`[export:${label}] Report has ${pagesFound.length} page(s): ${pagesFound.map(p => p.displayName).join(' | ')}`);

  // Select pages: filtered subset (by display name) or all.
  let selected: ReportPage[];
  if (options.pagesFilter && options.pagesFilter.length > 0) {
    const wanted = options.pagesFilter.map(s => s.toLowerCase());
    selected = pagesFound.filter(p => wanted.includes(p.displayName.toLowerCase()));
    const missing = options.pagesFilter.filter(
      w => !pagesFound.some(p => p.displayName.toLowerCase() === w.toLowerCase()),
    );
    if (missing.length > 0) {
      console.warn(`[export:${label}] Requested pages not found in report: ${missing.join(', ')}`);
    }
  } else {
    selected = pagesFound;
  }

  if (selected.length === 0) {
    console.warn(`[export:${label}] No pages selected to export.`);
  }

  // ── Build workbook ──────────────────────────────────────────────────────
  // Streaming writer: each row is flushed to disk (and freed from memory) as
  // soon as it's committed, so peak memory stays roughly proportional to ONE
  // page's data instead of the whole report's data — matters once you're
  // exporting many pages and/or large tables.
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: options.outPath,
    useStyles: true,          // keep the fill/font/border formatting
    useSharedStrings: false,  // shared strings trade memory for file size — skip, we're optimising for memory
  });
  wb.creator = 'Report Parity POC';
  wb.created = new Date();

  const usedSheetNames = new Set<string>();
  const exported: PageExportResult[] = [];

  if (options.beforeExport) {
    try {
      await options.beforeExport(page);
    } catch (e) {
      console.warn(`[export:${label}]   beforeExport hook failed — continuing without it: ${describeError(e)}`);
    }
  }

  for (const rp of selected) {
    console.log(`[export:${label}]   → page "${rp.displayName}"`);

    let navigated = true;
    try {
      await setReportPage(page, rp.name);
    } catch (e) {
      console.warn(`[export:${label}]     could not navigate to page "${rp.displayName}" — retrying once: ${describeError(e)}`);
      await page.waitForTimeout(NAV_RETRY_WAIT_MS);
      try {
        await setReportPage(page, rp.name);
      } catch (e2) {
        navigated = false;
        console.warn(`[export:${label}]     navigation to "${rp.displayName}" failed again — skipping this page (exporting it now would have captured the WRONG page's data): ${describeError(e2)}`);
      }
    }

    if (!navigated) {
      // Skip slicer application + export entirely — the report is still on
      // whatever page it was on before. Write an explicit marker sheet rather
      // than omitting it, so this reads as "we couldn't read this page" in the
      // comparison output instead of an ambiguous "page missing" diff.
      const desiredSheetName = sheetNameFor(rp.displayName);
      const wsName = toUniqueSheetName(desiredSheetName, usedSheetNames);
      const ws = wb.addWorksheet(wsName);
      const row = ws.addRow([`[PAGE NAVIGATION FAILED — could not reach "${rp.displayName}"; page skipped]`]);
      row.commit();
      ws.commit();
      exported.push({
        sheetName: wsName, pageDisplayName: rp.displayName, pageInternalName: rp.name,
        visualsWithData: 0, totalRows: 0,
      });
      continue;
    }

    // Apply slicer selections for this page (if a hook was provided), after
    // navigation and before export. Isolated so a slicer failure downgrades to
    // an unfiltered export of this page rather than failing the whole run.
    if (options.applySlicersForPage) {
      try {
        await options.applySlicersForPage(page, rp.displayName);
      } catch (e) {
        console.warn(`[export:${label}]     slicer application failed on "${rp.displayName}" — exporting UNFILTERED: ${describeError(e)}`);
      }
    }

    let visuals: VisualExport[] = [];
    try {
      visuals = await exportCurrentPageVisuals(page);
    } catch (e) {
      console.warn(`[export:${label}]     export failed on "${rp.displayName}" — retrying once: ${describeError(e)}`);
      await page.waitForTimeout(EXPORT_RETRY_WAIT_MS);
      try {
        visuals = await exportCurrentPageVisuals(page);
      } catch (e2) {
        console.warn(`[export:${label}]     export failed again on "${rp.displayName}" — page will be written with no data: ${describeError(e2)}`);
      }
    }

    const desiredSheetName = sheetNameFor(rp.displayName);
    const wsName = toUniqueSheetName(desiredSheetName, usedSheetNames);
    const ws = wb.addWorksheet(wsName);

    const { visualsWithData, totalRows } = writeVisualsToSheet(ws, visuals);
    ws.commit(); // flush this page's rows to disk now — don't hold them for the rest of the run

    exported.push({
      sheetName: wsName, pageDisplayName: rp.displayName, pageInternalName: rp.name,
      visualsWithData, totalRows,
    });
  }

  await wb.commit(); // finalises the xlsx file on disk
  const grandRows = exported.reduce((s, e) => s + e.totalRows, 0);
  console.log(`[export:${label}] Wrote ${exported.length} sheet(s), ${grandRows} data row(s) → ${options.outPath}`);

  return { pagesFound, exported, outPath: options.outPath };
}

// ── Column width bounds (characters) ───────────────────────────────────────
const MIN_COL_WIDTH = 12;
const MAX_COL_WIDTH = 42;
const WIDTH_PADDING = 2;

/**
 * Writes one page's visuals into an already-created worksheet: a marker row,
 * header row, and data rows per visual (skipping empty/non-data visuals),
 * sized so nothing is truncated.
 *
 * Exported (not just internal) so it can be exercised directly in tests
 * without a live Playwright page — see the parity smoke tests.
 *
 * Works with BOTH the in-memory Workbook and the streaming WorkbookWriter.
 * For the streaming writer specifically, two constraints from ExcelJS's
 * implementation shape this function's structure:
 *   1. Column widths (`ws.columns = [...]`) MUST be set before the FIRST row
 *      is committed — the streaming writer serialises `<cols>` the instant
 *      the first row is flushed, and never revisits it. So this does a
 *      lightweight first pass over the already-in-memory `visuals` data
 *      (just measuring string lengths — no ExcelJS calls) to compute widths
 *      BEFORE writing any row.
 *   2. Each row must be explicitly `.commit()`-ed as it's written, or it just
 *      sits in an internal buffer and is never actually freed — calling
 *      addRow() alone does NOT stream a row to disk.
 *
 * IMPORTANT: only styling (fill/font/border/width/wrap) is applied here.
 * Cell VALUES are written exactly as before — this has zero effect on
 * compareWorkbooks(), which reads only cell.value and ignores all styling.
 */
export function writeVisualsToSheet(
  ws: ExcelJS.Worksheet,
  visuals: VisualExport[],
): { visualsWithData: number; totalRows: number } {
  // Keep only data-bearing visuals, sorted for stable, aligned output.
  const dataVisuals = visuals
    .filter(v => !SKIP_TYPES.has(v.visualType) && v.rowCount > 0 && v.headers.length > 0)
    .sort((a, b) =>
      (a.visualTitle || a.visualName).localeCompare(b.visualTitle || b.visualName) ||
      a.visualName.localeCompare(b.visualName),
    );

  if (dataVisuals.length === 0) {
    ws.getColumn(1).width = 24;
    const row = ws.addRow(['[NO DATA VISUALS]']);
    row.commit();
    return { visualsWithData: 0, totalRows: 0 };
  }

  // ── PASS 1: compute per-column widths from data already in memory.
  // No ExcelJS writes yet — must finish this before touching a single row.
  const colWidths: number[] = [];
  const trackWidths = (vals: ReadonlyArray<string>) => {
    vals.forEach((v, i) => {
      const len = (v ?? '').toString().length;
      if (!colWidths[i] || len > colWidths[i]) colWidths[i] = len;
    });
  };
  for (const v of dataVisuals) {
    trackWidths(v.headers);
    for (const row of v.rows) trackWidths(v.headers.map(h => row[h] ?? ''));
  }
  ws.columns = colWidths.map(w => ({
    width: Math.min(Math.max(w + WIDTH_PADDING, MIN_COL_WIDTH), MAX_COL_WIDTH),
  }));

  // ── PASS 2: write + immediately commit every row, so the streaming writer
  // flushes each one to disk and frees it from memory right away instead of
  // holding the whole page (or worse, the whole report) in memory at once.
  let totalRows = 0;
  for (const v of dataVisuals) {
    const nc = v.headers.length;

    // Stable marker row — NO volatile row count, so identical data hashes equal.
    const secRow = ws.addRow([`[VISUAL] ${v.visualTitle || v.visualName} [${v.visualType}]`]);
    if (nc > 1) ws.mergeCells(secRow.number, 1, secRow.number, nc);
    const secCell = secRow.getCell(1);
    secCell.fill = SEC_FILL;
    secCell.font = { bold: true };
    secCell.border = CB;
    secCell.alignment = { vertical: 'middle', wrapText: true };
    secRow.height = 20;
    secRow.commit();

    const hdrRow = ws.addRow(v.headers);
    hdrRow.eachCell(c => {
      c.fill = HDR_FILL;
      c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      c.border = CB;
      c.alignment = { vertical: 'middle', wrapText: true };
    });
    hdrRow.height = 24;
    hdrRow.commit();

    for (const row of v.rows) {
      const vals = v.headers.map(h => row[h] ?? '');
      const dataRow = ws.addRow(vals);
      dataRow.eachCell(c => {
        c.border = CB;
        c.alignment = { vertical: 'middle' };
      });
      dataRow.commit();
    }
    totalRows += v.rowCount;

    const spacer = ws.addRow([]);
    spacer.commit();
  }

  return { visualsWithData: dataVisuals.length, totalRows };
}

// ── Parity summary writer ─────────────────────────────────────────────────────

export interface ParitySummaryMeta {
  pairName:     string;
  mode:         string;
  sourceLabel:  string;   // e.g. "Import mode"
  targetLabel:  string;   // e.g. "Direct Lake"
  sourceReportId: string;
  targetReportId: string;
  expectedFile: string;
  actualFile:   string;
}

/**
 * Whether a visual whose header TEXT differs but whose underlying VALUES match
 * (e.g. Import-mode friendly column name "Count of Application ID" vs
 * Direct-Lake raw name "Count of APPLICATION_ID") should fail the run.
 *
 * Default false: the migration is judged on whether the numbers match, not on
 * cosmetic label text. Header differences are still fully visible — every one
 * is listed on the "Visual Comparison" sheet and rolled into the summary
 * counts — they just don't flip the PASS/FAIL verdict on their own. Flip this
 * to true if you want label drift to be a hard failure instead.
 */
const TREAT_HEADER_DIFF_AS_FAILURE = false;

const E_PASS: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } };
const E_FAIL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
const E_HDR:  ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
const E_INFO: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9EFF7' } };
const E_WARN: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } };
const E_MISS: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
const E_BD: ExcelJS.Border = { style: 'thin', color: { argb: 'FFCCCCCC' } };
const E_CB: Partial<ExcelJS.Borders> = { top: E_BD, bottom: E_BD, left: E_BD, right: E_BD };

const VISUAL_STATUS_LABEL: Record<VisualBlockDiff['status'], string> = {
  'identical':         '✅ Identical',
  'header-diff':        '📝 Header text differs (values match)',
  'duplicate-count':    '🔁 Duplicated (values match)',
  'data-diff':          '❌ Values differ',
  'only-in-expected':   '⚠️ Only in Source',
  'only-in-actual':     '🆕 Only in Target',
  'ambiguous':          '❓ Ambiguous — needs manual check',
};

const VISUAL_STATUS_FILL: Record<VisualBlockDiff['status'], ExcelJS.Fill> = {
  'identical':        E_PASS,
  'header-diff':       E_INFO,
  'duplicate-count':   E_INFO,
  'data-diff':         E_FAIL,
  'only-in-expected':  E_MISS,
  'only-in-actual':    E_WARN,
  'ambiguous':         E_WARN,
};

/** True when a visual's status means "the data matches" (part of the relaxed gate). */
function visualDataMatches(status: VisualBlockDiff['status']): boolean {
  return status === 'identical' || status === 'header-diff' || status === 'duplicate-count';
}

/**
 * Writes the migration-parity summary workbook from a comparison result.
 * Returns { passed, differingSheets } for the caller's console verdict.
 *
 * PASS/FAIL is gated on `dataIdentical` by default (position-independent and
 * tolerant of header-text-only renames — see TREAT_HEADER_DIFF_AS_FAILURE),
 * not on the strict `identical` flag. Every difference, cosmetic or not, is
 * still written out — the relaxation only affects the verdict, never what's
 * shown.
 */
export async function writeParitySummary(
  diff: WorkbookDiffSummary,
  meta: ParitySummaryMeta,
  outPath: string,
): Promise<{ passed: boolean; differingSheets: number }> {
  const passed = TREAT_HEADER_DIFF_AS_FAILURE ? diff.identical : diff.dataIdentical;
  const differing = diff.sheets.filter(s => TREAT_HEADER_DIFF_AS_FAILURE ? !s.identical : !s.dataIdentical).length;

  const totalHeaderDiffs = diff.sheets.reduce((s, sh) => s + sh.headerDiffCount, 0);
  const totalStructuralDiffs = diff.sheets.reduce((s, sh) => s + sh.structuralDiffCount, 0);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Report Parity POC';
  wb.created = new Date();

  // ── Sheet 1: Summary ──────────────────────────────────────────────────────
  const infoWs = wb.addWorksheet('Parity Summary');
  infoWs.columns = [{ width: 30 }, { width: 60 }];
  const verdict = passed
    ? '✅ PASS — source and target data match'
    : `❌ FAIL — ${differing} page(s) have genuine data/structure differences`;
  const rows: [string, string][] = [
    ['Pair',                meta.pairName],
    ['Run Mode',            meta.mode],
    ['Run Time',            new Date().toLocaleString()],
    [`Source (${meta.sourceLabel})`, meta.sourceReportId],
    [`Target (${meta.targetLabel})`, meta.targetReportId],
    ['Expected File',       meta.expectedFile],
    ['Actual File',         meta.actualFile],
    ['Pages Compared',      String(diff.sheets.filter(s => s.inExpected && s.inActual).length)],
    ['Pages Fully Identical (strict)', String(diff.sheets.filter(s => s.identical).length)],
    ['Pages With Only Header-Text/Order Diffs', String(diff.sheets.filter(s => !s.identical && s.dataIdentical).length)],
    ['Pages With Real Data/Structure Diffs', String(differing)],
    ['Pages Only in Source', diff.sheetsOnlyInExpected.join(', ') || '(none)'],
    ['Pages Only in Target', diff.sheetsOnlyInActual.join(', ') || '(none)'],
    ['Header-Only Visual Diffs (see Visual Comparison)', String(totalHeaderDiffs)],
    ['Structural Visual Diffs — missing/duplicated/ambiguous (see Visual Comparison)', String(totalStructuralDiffs)],
    ['Header diffs count as failure?', TREAT_HEADER_DIFF_AS_FAILURE ? 'Yes' : 'No — see TREAT_HEADER_DIFF_AS_FAILURE'],
    ['Overall Result',      verdict],
  ];
  rows.forEach(([label, value]) => {
    const r = infoWs.addRow([label, value]);
    r.getCell(1).font = { bold: true };
    r.getCell(1).fill = E_INFO;
    r.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
  });
  const verdictRow = infoWs.getRow(rows.length);
  verdictRow.getCell(2).fill = passed ? E_PASS : E_FAIL;
  verdictRow.getCell(2).font = { bold: true, color: { argb: 'FFFFFFFF' } };

  // ── Sheet 2: Page Comparison ──────────────────────────────────────────────
  const cmpWs = wb.addWorksheet('Page Comparison');
  cmpWs.columns = [
    { header: 'Page (Sheet)',            key: 'sheet',   width: 26 },
    { header: 'Rows in Source',          key: 'exp',     width: 14 },
    { header: 'Rows in Target',          key: 'act',     width: 14 },
    { header: 'Header-Only Diffs',       key: 'hdrDiff', width: 16 },
    { header: 'Structural Diffs',        key: 'strDiff', width: 16 },
    { header: 'Only in Source (approx.)', key: 'onlyExp', width: 18 },
    { header: 'Only in Target (approx.)', key: 'onlyAct', width: 18 },
    { header: 'Status',                  key: 'status',  width: 30 },
  ];
  cmpWs.getRow(1).eachCell(c => {
    c.fill = E_HDR; c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
    c.border = E_CB; c.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  for (const s of diff.sheets) {
    let statusText: string;
    let statusFill: ExcelJS.Fill;
    if (s.inExpected && !s.inActual) { statusText = '⚠️ Missing in Target'; statusFill = E_WARN; }
    else if (!s.inExpected && s.inActual) { statusText = '🆕 Only in Target'; statusFill = E_WARN; }
    else if (s.identical) { statusText = '✅ Identical'; statusFill = E_PASS; }
    else if (s.dataIdentical) { statusText = '📝 Data OK — header text/order differs'; statusFill = E_INFO; }
    else { statusText = '❌ Different'; statusFill = E_FAIL; }

    const row = cmpWs.addRow({
      sheet: s.sheet,
      exp: s.rowsExpected,
      act: s.rowsActual,
      hdrDiff: s.headerDiffCount,
      strDiff: s.structuralDiffCount,
      onlyExp: s.rowsOnlyInExpected,
      onlyAct: s.rowsOnlyInActual,
      status: statusText,
    });
    row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle' }; });
    const sc = row.getCell('status');
    sc.fill = statusFill;
    sc.font = { bold: true, color: { argb: (statusFill === E_WARN || statusFill === E_INFO) ? 'FF000000' : 'FFFFFFFF' } };
    sc.alignment = { horizontal: 'center', vertical: 'middle' };
  }

  // ── Sheet 3: Visual Comparison (one row per Power BI visual, per page) ────
  // This is the actionable sheet: instead of a pile of raw row diffs, it says
  // exactly WHICH visual on WHICH page has a problem and what kind.
  const visWs = wb.addWorksheet('Visual Comparison');
  visWs.columns = [
    { header: 'Page (Sheet)',   key: 'sheet',   width: 24 },
    { header: 'Visual Title',   key: 'title',   width: 34 },
    { header: 'Type',           key: 'type',    width: 20 },
    { header: 'Status',         key: 'status',  width: 32 },
    { header: 'Copies (Src/Tgt)', key: 'counts', width: 16 },
    { header: 'Rows (Src/Tgt)', key: 'rows',    width: 16 },
    { header: 'Header (Source)', key: 'hdrExp', width: 34 },
    { header: 'Header (Target)', key: 'hdrAct', width: 34 },
    { header: 'Note',           key: 'note',    width: 60 },
  ];
  visWs.getRow(1).eachCell(c => {
    c.fill = E_HDR; c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
    c.border = E_CB; c.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  let anyVisualRows = false;
  for (const s of diff.sheets) {
    for (const v of s.visuals) {
      anyVisualRows = true;
      const row = visWs.addRow({
        sheet: s.sheet,
        title: v.title,
        type: v.type,
        status: VISUAL_STATUS_LABEL[v.status],
        counts: `${v.countExpected} / ${v.countActual}`,
        rows: `${v.rowsExpected} / ${v.rowsActual}`,
        hdrExp: v.headerExpected.join(' | '),
        hdrAct: v.headerActual.join(' | '),
        note: v.note ?? '',
      });
      row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
      const sc = row.getCell('status');
      const fill = VISUAL_STATUS_FILL[v.status];
      sc.fill = fill;
      sc.font = { bold: true, color: { argb: (fill === E_WARN || fill === E_INFO || fill === E_MISS) ? 'FF000000' : 'FFFFFFFF' } };
    }
  }
  if (!anyVisualRows) {
    visWs.addRow(['(No per-visual breakdown — every page matched at the fast whole-sheet check.)']);
  }

  // ── Sheet 4: Differences (sample rows for genuine data/structure problems) ─
  // Header-only and duplicate-count visuals are deliberately excluded here —
  // they're fully described on the Visual Comparison sheet already, and
  // mixing them in here is exactly the noise this rework was meant to remove.
  const problemVisuals: Array<{ sheet: string; visual: VisualBlockDiff }> = [];
  for (const s of diff.sheets) {
    for (const v of s.visuals) {
      if (!visualDataMatches(v.status)) problemVisuals.push({ sheet: s.sheet, visual: v });
    }
  }

  if (problemVisuals.length > 0) {
    const dWs = wb.addWorksheet('Differences');
    dWs.columns = [
      { header: 'Page (Sheet)', key: 'sheet',  width: 22 },
      { header: 'Visual',       key: 'visual', width: 30 },
      { header: 'Side',         key: 'side',   width: 26 },
      { header: 'Sample Row',   key: 'row',    width: 100 },
    ];
    dWs.getRow(1).eachCell(c => {
      c.fill = E_HDR; c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      c.border = E_CB; c.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    for (const { sheet, visual } of problemVisuals) {
      const label = `${visual.title} [${visual.type}]`;
      for (const r of visual.sampleOnlyInExpected) {
        const row = dWs.addRow({ sheet, visual: label, side: 'Only in Source (Expected)', row: r.join(' | ') });
        row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
        row.getCell('side').fill = E_MISS;
        row.getCell('side').font = { bold: true };
      }
      for (const r of visual.sampleOnlyInActual) {
        const row = dWs.addRow({ sheet, visual: label, side: 'Only in Target (Actual)', row: r.join(' | ') });
        row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
        row.getCell('side').fill = E_PASS;
        row.getCell('side').font = { bold: true, color: { argb: 'FFFFFFFF' } };
      }
      if (visual.sampleOnlyInExpected.length === 0 && visual.sampleOnlyInActual.length === 0) {
        // only-in-expected / only-in-actual / ambiguous visuals with no row-level
        // sample still deserve a line so the page isn't silently missing context.
        const row = dWs.addRow({ sheet, visual: label, side: VISUAL_STATUS_LABEL[visual.status], row: visual.note ?? '' });
        row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
      }
    }
  }

  fs.mkdirSync(require('path').dirname(outPath), { recursive: true });
  await wb.xlsx.writeFile(outPath);
  return { passed, differingSheets: differing };
}