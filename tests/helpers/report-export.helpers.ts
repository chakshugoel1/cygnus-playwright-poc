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
  /** pagesFilter entries that matched NO page in this report — a strong sign
   *  the two sides' page names have drifted (renamed/removed pages). */
  requestedPagesMissing: string[];
}

/**
 * Canonical form for matching a page/filter display name across two reports.
 * Migrated reports and cross-report filter titles drift in exactly these
 * ways while still meaning the same thing: trailing/leading whitespace
 * ("Vertragsversand " vs "Vertragsversand"), case ("- Cluster" vs
 * "- cluster"), and separator style ("Recruiting_Weg" vs "Recruiting Weg"
 * vs "RecruitingWeg" — separators are stripped entirely, not just
 * normalized to one style, so "a_bc" and "abc" compare equal). Shared by
 * page-name matching (this file) and cross-report filter-title matching
 * (cross-report-match.helpers.ts) — one rule, used everywhere names from
 * two different reports need to be compared.
 */
export function pageNameKey(name: string): string {
  return name.trim().toLowerCase().replace(/[\s_-]+/g, '');
}

/**
 * Pure page-set alignment between a source and target report, for the parity
 * summary. `pageMap` (source name → target name) is applied before matching;
 * names are matched trimmed + case-insensitively.
 */
export function diffPageSets(
  sourcePages: string[],
  targetPages: string[],
  pageMap?: Record<string, string>,
): { inBoth: string[]; onlyInSource: string[]; onlyInTarget: string[] } {
  const mapByKey = new Map(Object.entries(pageMap ?? {}).map(([s, t]) => [pageNameKey(s), t]));
  const targetKeys = new Set(targetPages.map(pageNameKey));

  const inBoth: string[] = [];
  const onlyInSource: string[] = [];
  const matchedTargetKeys = new Set<string>();

  for (const src of sourcePages) {
    const mapped = mapByKey.get(pageNameKey(src)) ?? src;
    const key = pageNameKey(mapped);
    if (targetKeys.has(key)) {
      inBoth.push(src);
      matchedTargetKeys.add(key);
    } else {
      onlyInSource.push(src);
    }
  }

  const onlyInTarget = targetPages.filter(t => !matchedTargetKeys.has(pageNameKey(t)));
  return { inBoth, onlyInSource, onlyInTarget };
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

  // Select pages: filtered subset (by display name, trimmed + case-insensitive
  // — migrated reports drift by trailing spaces and case) or all.
  let selected: ReportPage[];
  let requestedPagesMissing: string[] = [];
  if (options.pagesFilter && options.pagesFilter.length > 0) {
    const wanted = options.pagesFilter.map(pageNameKey);
    selected = pagesFound.filter(p => wanted.includes(pageNameKey(p.displayName)));
    requestedPagesMissing = options.pagesFilter.filter(
      w => !pagesFound.some(p => pageNameKey(p.displayName) === pageNameKey(w)),
    );
    if (requestedPagesMissing.length > 0) {
      console.warn(`[export:${label}] Requested pages not found in report: ${requestedPagesMissing.join(', ')}`);
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

  return { pagesFound, exported, outPath: options.outPath, requestedPagesMissing };
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
  /**
   * Problems that make the comparison itself untrustworthy (a filter that
   * applied on one side but failed on the other, requested pages missing on
   * one side, etc). Any entry here FORCES the verdict to a non-pass — a
   * filtered-vs-unfiltered comparison must never present itself as a clean
   * result, pass OR fail.
   */
  comparisonCaveats?: string[];
  /** Full page-set alignment between the two reports (informational). */
  pageAlignment?: { onlyInSource: string[]; onlyInTarget: string[]; inBothCount: number };
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

// ── Severity tiers ───────────────────────────────────────────────────────────
//
// A value difference and a header/label rename are not the same kind of
// problem — a renamed column is expected, cosmetic noise from a migration; a
// visual with different numbers (or one that vanished entirely) is a real
// finding. VISUAL_STATUS_LABEL/FILL above already say WHICH of 7 fine-grained
// statuses a visual has, but reading 7 different labels to work out "is this
// one bad?" doesn't scale. Severity collapses them to 3 buckets so the sheet
// can be scanned (or filtered) at a glance, and gives the row itself a light
// background tint (full-row, not just the Status cell) so a critical row is
// visually unmistakable even scrolling past it quickly.
type Severity = 'safe' | 'review' | 'critical';

const VISUAL_SEVERITY: Record<VisualBlockDiff['status'], Severity> = {
  'identical':        'safe',
  'header-diff':      'safe',     // values match — only the label text changed
  'duplicate-count':  'safe',     // values match — only the copy count changed
  'ambiguous':        'review',   // couldn't be auto-verified either way
  'only-in-expected': 'critical', // a visual genuinely vanished
  'only-in-actual':   'critical', // a visual genuinely appeared
  'data-diff':        'critical', // same visual, different numbers
};

const SEVERITY_LABEL: Record<Severity, string> = {
  safe:     '✅ Safe — cosmetic only',
  review:   '❓ Needs manual review',
  critical: '❌ Critical — real difference',
};

// Pale versions of the status fills, meant for tinting an entire row rather
// than one bold cell — strong enough to scan down a column of rows, light
// enough that black text stays readable on top of it.
const E_PASS_LT:     ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2F3E1' } };
const E_INFO_LT:     ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF1FB' } };
const E_WARN_LT:     ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF1D6' } };
const E_FAIL_LT:     ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE1E1' } };

const SEVERITY_ROW_FILL: Record<Severity, ExcelJS.Fill> = {
  safe:     E_PASS_LT,
  review:   E_WARN_LT,
  critical: E_FAIL_LT,
};
const SEVERITY_CELL_FILL: Record<Severity, ExcelJS.Fill> = {
  safe:     E_PASS,
  review:   E_WARN,
  critical: E_FAIL,
};

// header-diff/duplicate-count share the 'safe' severity but should still tint
// slightly differently from a truly byte-identical visual, so a reviewer can
// still tell "renamed" from "untouched" without reading the Status column —
// this only affects the ROW tint, never the severity bucket itself.
function rowFillFor(status: VisualBlockDiff['status']): ExcelJS.Fill {
  if (status === 'header-diff' || status === 'duplicate-count') return E_INFO_LT;
  return SEVERITY_ROW_FILL[VISUAL_SEVERITY[status]];
}

// ── Plain-English analysis ────────────────────────────────────────────────────
//
// The rest of this workbook answers "what differs, exactly, and where" — this
// answers the question a reviewer actually asks first: "do I need to worry
// about this?" It's generated purely from counts already computed elsewhere
// in this function (no new comparison logic), so it can never disagree with
// the detailed sheets — it's a summary OF them, not a second opinion.

export interface AnalysisStats {
  /** Visuals actually inspected in detail (Tier 2) — excludes pages that
   *  matched byte-for-byte at the fast Tier 1 check, which never get a
   *  per-visual breakdown because there's nothing to break down. */
  totalVisuals: number;
  identicalVisuals: number;
  /** header-diff + duplicate-count: values match, only cosmetic. */
  cosmeticOnly: number;
  /** ambiguous: could not be auto-verified either way. */
  needsReview: number;
  /** only-in-expected + only-in-actual + data-diff: a genuine difference. */
  critical: number;
  /** Pages that matched byte-for-byte at the fast check (no visual detail). */
  fastPassPages: number;
}

export function computeAnalysisStats(diff: WorkbookDiffSummary): AnalysisStats {
  const stats: AnalysisStats = {
    totalVisuals: 0, identicalVisuals: 0, cosmeticOnly: 0, needsReview: 0, critical: 0, fastPassPages: 0,
  };
  for (const s of diff.sheets) {
    if (!s.inExpected || !s.inActual) continue; // whole-page-missing handled separately, via sheetsOnlyIn*
    if (s.visuals.length === 0) {
      if (s.identical) stats.fastPassPages++;
      continue;
    }
    for (const v of s.visuals) {
      stats.totalVisuals++;
      switch (VISUAL_SEVERITY[v.status]) {
        case 'safe':     if (v.status === 'identical') stats.identicalVisuals++; else stats.cosmeticOnly++; break;
        case 'review':   stats.needsReview++; break;
        case 'critical': stats.critical++; break;
      }
    }
  }
  return stats;
}

/** Returns 2-4 short sentences, each its own row in the summary. */
export function buildAnalysisNarrative(
  diff: WorkbookDiffSummary,
  stats: AnalysisStats,
  passed: boolean,
  caveats: string[],
): string[] {
  const lines: string[] = [];
  const pagesMissing = diff.sheetsOnlyInExpected.length + diff.sheetsOnlyInActual.length;

  if (caveats.length > 0) {
    lines.push(
      `⚠ This run is NOT directly comparable — ${caveats.length} caveat(s) affected it (see the Caveat rows below). ` +
      `Resolve those before trusting the pass/fail verdict.`,
    );
  }

  const diffVisuals = stats.totalVisuals - stats.identicalVisuals;

  if (stats.totalVisuals === 0) {
    lines.push(stats.fastPassPages > 0
      ? `All ${stats.fastPassPages} compared page(s) matched byte-for-byte — no differences of any kind.`
      : 'No pages were compared in detail (nothing to analyze).');
  } else if (diffVisuals === 0) {
    lines.push(`All ${stats.totalVisuals} visual(s) inspected are identical — no differences of any kind.`);
  } else {
    const pctCosmetic = Math.round((stats.cosmeticOnly / diffVisuals) * 100);
    if (stats.critical === 0 && stats.needsReview === 0) {
      lines.push(
        `Good news: every difference found (${stats.cosmeticOnly} of ${diffVisuals}) is cosmetic only — ` +
        `column/visual labels were renamed during migration, but every underlying VALUE still matches. Nothing here needs action.`,
      );
    } else {
      const parts: string[] = [];
      if (stats.cosmeticOnly > 0) parts.push(`${stats.cosmeticOnly} cosmetic label rename(s) (${pctCosmetic}% of all differences — values match, safe to ignore)`);
      if (stats.critical > 0) parts.push(`${stats.critical} critical difference(s) (missing/added visuals or values that genuinely don't match)`);
      if (stats.needsReview > 0) parts.push(`${stats.needsReview} visual(s) needing a manual look (duplicate titles the tool couldn't auto-match)`);
      lines.push(`Of ${diffVisuals} visual(s) that differ: ${parts.join('; ')}.`);
      lines.push(stats.critical > 0
        ? `Bottom line: ${stats.critical} genuine issue(s) need review before sign-off${stats.cosmeticOnly > 0 ? ` — the other ${stats.cosmeticOnly} are just renamed labels and can be ignored` : ''}.`
        : `Bottom line: nothing critical, but ${stats.needsReview} visual(s) need a quick manual check (see "❓ Needs manual review" rows on Visual Comparison).`);
    }
  }

  if (pagesMissing > 0) {
    lines.push(`${pagesMissing} entire page(s) exist on only one side of the comparison (see "Pages Only in Source/Target" below) — not counted in the visual stats above.`);
  }

  return lines;
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
  const caveats = meta.comparisonCaveats ?? [];
  const dataPassed = TREAT_HEADER_DIFF_AS_FAILURE ? diff.identical : diff.dataIdentical;
  // Caveats override everything: a run where the two sides weren't filtered
  // identically (or pages went missing) is NOT comparable, so neither a green
  // PASS nor a plain FAIL would be honest.
  const passed = dataPassed && caveats.length === 0;
  const differing = diff.sheets.filter(s => TREAT_HEADER_DIFF_AS_FAILURE ? !s.identical : !s.dataIdentical).length;

  const totalHeaderDiffs = diff.sheets.reduce((s, sh) => s + sh.headerDiffCount, 0);
  const totalStructuralDiffs = diff.sheets.reduce((s, sh) => s + sh.structuralDiffCount, 0);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Report Parity POC';
  wb.created = new Date();

  // ── Sheet 1: Summary ──────────────────────────────────────────────────────
  const infoWs = wb.addWorksheet('Parity Summary');
  infoWs.columns = [{ width: 30 }, { width: 60 }];

  // Analysis narrative FIRST — the plain-English "do I need to worry about
  // this" answer a reviewer wants before any raw counts. Purely derived from
  // stats computed below; never a second source of truth.
  const analysisStats = computeAnalysisStats(diff);
  const analysisLines = buildAnalysisNarrative(diff, analysisStats, passed, caveats);
  const analysisHeaderRow = infoWs.addRow(['📋 Analysis', '']);
  infoWs.mergeCells(analysisHeaderRow.number, 1, analysisHeaderRow.number, 2);
  analysisHeaderRow.getCell(1).fill = E_HDR;
  analysisHeaderRow.getCell(1).font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
  analysisHeaderRow.getCell(1).alignment = { vertical: 'middle' };
  analysisHeaderRow.height = 22;
  for (const line of analysisLines) {
    const r = infoWs.addRow([line, '']);
    infoWs.mergeCells(r.number, 1, r.number, 2);
    const c = r.getCell(1);
    c.font = { italic: true };
    c.alignment = { vertical: 'middle', wrapText: true };
    c.fill = E_INFO;
    r.height = 32;
  }
  infoWs.addRow(['', '']); // spacer before the raw stats table below

  const verdict = caveats.length > 0
    ? `⚠️ NOT COMPARABLE — ${caveats.length} caveat(s) invalidate this run (see Caveat rows below)`
    : passed
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
  ];
  if (meta.pageAlignment) {
    rows.push(['Report Pages In Both Reports', String(meta.pageAlignment.inBothCount)]);
    rows.push(['Report Pages Only In Source Report', meta.pageAlignment.onlyInSource.join(', ') || '(none)']);
    rows.push(['Report Pages Only In Target Report', meta.pageAlignment.onlyInTarget.join(', ') || '(none)']);
  }
  caveats.forEach((c, i) => rows.push([`⚠ Caveat ${i + 1}`, c]));
  rows.push(['Overall Result', verdict]);
  let verdictRow: ExcelJS.Row | undefined;
  rows.forEach(([label, value]) => {
    const r = infoWs.addRow([label, value]);
    r.getCell(1).font = { bold: true };
    r.getCell(1).fill = label.startsWith('⚠ Caveat') ? E_WARN : E_INFO;
    r.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
    if (label === 'Overall Result') verdictRow = r;
  });
  verdictRow!.getCell(2).fill = caveats.length > 0 ? E_WARN : (passed ? E_PASS : E_FAIL);
  verdictRow!.getCell(2).font = { bold: true, color: { argb: caveats.length > 0 ? 'FF000000' : 'FFFFFFFF' } };

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
    { header: 'Severity',       key: 'severity',width: 26 },
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
  visWs.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: visWs.columns.length } };

  let anyVisualRows = false;
  for (const s of diff.sheets) {
    for (const v of s.visuals) {
      anyVisualRows = true;
      const severity = VISUAL_SEVERITY[v.status];
      const row = visWs.addRow({
        sheet: s.sheet,
        title: v.title,
        type: v.type,
        severity: SEVERITY_LABEL[severity],
        status: VISUAL_STATUS_LABEL[v.status],
        counts: `${v.countExpected} / ${v.countActual}`,
        rows: `${v.rowsExpected} / ${v.rowsActual}`,
        hdrExp: v.headerExpected.join(' | '),
        hdrAct: v.headerActual.join(' | '),
        note: v.note ?? '',
      });
      // Whole-row tint by severity first (so every cell — including Header
      // Source/Target, where a value-vs-label difference is actually visible
      // — carries the signal), then the two callout cells get their bolder,
      // more saturated fill on top so they still pop within the tinted row.
      const rowFill = rowFillFor(v.status);
      row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; c.fill = rowFill; });

      const sevCell = row.getCell('severity');
      sevCell.fill = SEVERITY_CELL_FILL[severity];
      sevCell.font = { bold: true, color: { argb: severity === 'critical' ? 'FFFFFFFF' : 'FF000000' } };

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