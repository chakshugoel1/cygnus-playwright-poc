/**
 * excel-compare.helpers.ts
 *
 * Streaming Excel workbook comparison, in three passes:
 *
 * Tier 1 — per-sheet ordered hash (fast path).
 *   Stream both files' matching sheet, feeding each row into a running SHA-1.
 *   If the two hashes match → sheet is byte-identical → done for that sheet.
 *   A mismatch at Tier 1 proves nothing (could be mere row reordering) so always fall
 *   through to Tier 2 on any mismatch.
 *
 * Tier 2 — visual-block-aware diff (the detailed pass).
 *   report-export.helpers.ts writes each Power BI visual as a self-contained block:
 *   a "[VISUAL] <title> [<type>]" marker row, a header row, its data rows, and a
 *   blank spacer. This tier re-groups the row stream back into those blocks and
 *   compares VISUAL TO VISUAL rather than treating the sheet as one flat bag of
 *   rows. That buys three things a flat row-diff cannot:
 *     - A table that simply moved to a different position/row range in the sheet
 *       is still matched correctly, because matching is by visual identity
 *       (title + type), not row position. A flat row-multiset already tolerated
 *       *row* reordering; this additionally tolerates *visual* reordering and
 *       keeps row matches scoped to the correct visual instead of letting two
 *       coincidentally-identical rows from unrelated visuals cancel out.
 *     - Header-text differences (e.g. an Import-mode friendly column name vs a
 *       Direct-Lake raw column name) are reported separately from genuine value
 *       differences, instead of being blended into one undifferentiated pile of
 *       "only in expected / only in actual" rows.
 *     - The same visual title appearing a different number of times on each side
 *       (duplicate slicers, a visual that only exists on one side, two distinctly
 *       different visuals that happen to share a title) is called out explicitly
 *       instead of silently fragmenting into dozens of unexplained row diffs.
 *
 * `identical` (per sheet and per workbook) stays STRICT: every visual must match
 * on title, type, count, headers AND data. `dataIdentical` is the relaxed sibling:
 * true when every visual's underlying VALUES match, even if headers were renamed,
 * columns were reordered, or a visual is duplicated identically on one side.
 * Callers that want "position/renaming-tolerant" pass/fail should gate on
 * `dataIdentical`; callers that want a byte-for-byte check should use `identical`.
 *
 * Design constraints (per implementation brief):
 *   - Uses ExcelJS streaming reader only — never document-mode readFile()
 *   - No new npm dependencies: exceljs (already installed) + Node built-in crypto
 *   - Memory-safe: Tier 1 hashes rows and discards them immediately. Tier 2 holds
 *     at most one visual's rows in memory at a time while parsing (not the whole
 *     sheet), matching the memory profile the streaming writer already targets.
 */

import * as crypto from 'crypto';
import * as path   from 'path';
import ExcelJS     from 'exceljs';

// Central-directory zip access (Open.file reads only the zip's directory, then
// inflates single entries on demand — no whole-file buffering). Already a
// transitive dependency of exceljs; declared as a direct dependency in
// package.json so an exceljs upgrade can never silently remove it.
// eslint-disable-next-line @typescript-eslint/no-var-requires
const unzipper = require('unzipper');

// ── Deterministic sheet-name resolution ──────────────────────────────────────
//
// ExcelJS's streaming READER resolves a worksheet's display name by looking it
// up in workbook.xml — but only if workbook.xml happened to appear EARLIER in
// the zip than the worksheet entry. Files produced by ExcelJS's streaming
// WRITER (i.e. our own exports) put the worksheets FIRST and workbook.xml
// last, and the relative order of workbook.xml vs its .rels file is not even
// deterministic across runs. Depending on that order the reader either:
//   - crashes ("Cannot read properties of undefined (reading 'sheets')" in
//     _parseWorksheet — rels parsed, workbook.xml not yet), or
//   - silently falls back to default names ("Sheet1", "Sheet2", ...), which
//     then destroys sheet pairing between the two workbooks being compared.
// Both were observed in production on a machine validating two real reports.
//
// Fix: read the sheet list ourselves, directly from workbook.xml (+ its rels
// for the sheetN.xml file mapping) via the zip central directory — fully
// order-independent — and PRE-SEED the ExcelJS reader's internal model before
// iterating, so its lookup always succeeds and its crash branch is
// unreachable. Belt-and-braces: sheet matching below also accepts the
// file-number id, and a sheet that is never emitted at all now throws instead
// of silently producing an empty result.

/** One sheet as resolved from workbook.xml: display name + sheetN.xml number. */
export interface SheetInfo {
  name:    string;  // real display name (page name), exactly as written
  fileNo:  string;  // N in xl/worksheets/sheetN.xml
  rId:     string;  // relationship id linking workbook.xml to the file
  sheetId: string;  // workbook.xml sheetId attribute
}

function decodeXmlEntities(text: string): string {
  return text
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

/** Reads the ordered sheet list from an xlsx via its central directory. */
export async function readWorkbookSheetInfo(filePath: string): Promise<SheetInfo[]> {
  const dir = await unzipper.Open.file(filePath);
  const entryFor = (p: string) => dir.files.find((f: any) => f.path === p);

  const wbEntry = entryFor('xl/workbook.xml');
  if (!wbEntry) throw new Error(`Not a valid xlsx (no xl/workbook.xml): ${filePath}`);
  const wbXml = (await wbEntry.buffer()).toString('utf8');

  // rId → sheet file number, from the workbook rels (may be absent in odd files)
  const relTargets = new Map<string, string>();
  const relsEntry = entryFor('xl/_rels/workbook.xml.rels');
  if (relsEntry) {
    const relsXml = (await relsEntry.buffer()).toString('utf8');
    for (const rel of relsXml.match(/<Relationship\b[^>]*>/g) ?? []) {
      const id     = rel.match(/\bId="([^"]+)"/)?.[1];
      const target = rel.match(/\bTarget="[^"]*worksheets\/sheet(\d+)\.xml"/)?.[1];
      if (id && target) relTargets.set(id, target);
    }
  }

  const sheets: SheetInfo[] = [];
  for (const tag of wbXml.match(/<sheet\b[^>]*>/g) ?? []) {
    const name    = tag.match(/\bname="([^"]*)"/)?.[1];
    const rId     = tag.match(/\br:id="([^"]+)"/)?.[1] ?? '';
    const sheetId = tag.match(/\bsheetId="([^"]+)"/)?.[1] ?? '';
    if (name === undefined) continue;
    // Fall back to declaration order when rels are missing — matches how the
    // ExcelJS writer numbers its sheet files.
    const fileNo = (rId && relTargets.get(rId)) || String(sheets.length + 1);
    sheets.push({ name: decodeXmlEntities(name), fileNo, rId, sheetId });
  }
  if (sheets.length === 0) throw new Error(`No sheets declared in workbook.xml: ${filePath}`);
  return sheets;
}

/**
 * Creates an ExcelJS streaming reader with its internal workbook model
 * PRE-SEEDED from our own central-directory parse, so _parseWorksheet's
 * name lookup always succeeds regardless of zip entry order (see the long
 * comment above). Touches two internal fields of the pinned exceljs 4.4.0
 * reader (model, workbookRels); the ws.id fallback in isTargetSheet keeps
 * working even if a future exceljs version ignores the seeding.
 */
function makeSeededReader(filePath: string, sheets: SheetInfo[]): any {
  const reader = new (ExcelJS as any).stream.xlsx.WorkbookReader(filePath, {
    sharedStrings: 'cache',
    hyperlinks:    'ignore',
    styles:        'ignore',
    entries:       'emit',
  });
  reader.model = {
    sheets: sheets.map((s, i) => ({
      id: Number(s.sheetId) || i + 1,
      name: s.name,
      rId: s.rId || `rId${s.fileNo}`,
    })),
  };
  reader.workbookRels = sheets.map(s => ({
    Id:     s.rId || `rId${s.fileNo}`,
    Target: `worksheets/sheet${s.fileNo}.xml`,
  }));
  return reader;
}

/** True when the emitted worksheet is the one described by `target`. */
function isTargetSheet(ws: any, target: SheetInfo, realNames: Set<string>): boolean {
  const wsName = typeof ws.name === 'string' ? ws.name : '';
  // A name we recognise from workbook.xml is authoritative; otherwise the
  // reader fell back to its default ("Sheet<fileNo>") and the id — which in
  // that fallback IS the file number — identifies the sheet instead.
  if (realNames.has(wsName)) return wsName === target.name;
  return String(ws.id) === String(target.fileNo);
}

/** Canonical pairing key: sheet names differing only by case/edge-whitespace
 *  (e.g. "…- Cluster" vs "…- cluster" between a source and a migrated report)
 *  refer to the same page and must pair up. */
function sheetKey(name: string): string {
  return name.trim().toLowerCase();
}

// ── Public types ─────────────────────────────────────────────────────────────

export type VisualDiffStatus =
  | 'identical'        // headers + data + counts all match
  | 'header-diff'      // data matches; header text and/or column order differs
  | 'duplicate-count'  // headers + data match; this visual appears a different
                        // number of times on each side (all copies mutually identical)
  | 'data-diff'        // matched 1:1 but the underlying values differ
  | 'only-in-expected' // visual (by title+type) exists only in the source export
  | 'only-in-actual'   // visual (by title+type) exists only in the target export
  | 'ambiguous';        // more than one distinct variant on at least one side —
                        // could not safely auto-pair, needs a human look

export interface VisualBlockDiff {
  title:          string;
  type:           string;
  status:         VisualDiffStatus;
  /** How many blocks with this exact (title, type) were found on each side. */
  countExpected:  number;
  countActual:    number;
  rowsExpected:   number;
  rowsActual:     number;
  headerExpected: string[];
  headerActual:   string[];
  /** Up to SAMPLE_SIZE data rows found only in the expected copy of this visual. */
  sampleOnlyInExpected: string[][];
  /** Up to SAMPLE_SIZE data rows found only in the actual copy of this visual. */
  sampleOnlyInActual:   string[][];
  /** True total (not capped by SAMPLE_SIZE) of differing rows, for data-diff visuals. */
  onlyInExpectedCount:  number;
  onlyInActualCount:    number;
  /** Human-readable explanation — why this status, what to go check. */
  note?: string;
}

export interface SheetDiffSummary {
  sheet:         string;
  inExpected:    boolean;
  inActual:      boolean;
  rowsExpected:  number;
  rowsActual:    number;
  /** STRICT: every visual matches on title, type, count, headers AND data. */
  identical:     boolean;
  /** RELAXED: every visual's underlying data matches (headers/order/duplicate
   *  count differences do not count against this). Use this to gate pass/fail
   *  when you only care that the migrated report shows the same numbers. */
  dataIdentical: boolean;
  /** Per-visual breakdown. Empty when the sheet was identical at Tier 1 (no
   *  detailed pass was needed) or when neither export has any visuals. */
  visuals: VisualBlockDiff[];
  /** Visuals whose header text/column order differs but whose data matches. */
  headerDiffCount: number;
  /** Visuals that are missing on one side, duplicated a different number of
   *  times, or otherwise couldn't be unambiguously matched. */
  structuralDiffCount: number;
  /** Approximate, human-scale row-diff counts for at-a-glance display — NOT
   *  used to determine identical/dataIdentical (those are computed from
   *  `visuals` directly). Kept for backward compatibility with existing
   *  consumers of this shape (cygnus-main-run.spec.ts, excel-compare.spec.ts). */
  rowsOnlyInExpected: number;
  rowsOnlyInActual:   number;
  /** Up to SAMPLE_SIZE rows only in the expected file (human-readable) */
  sampleOnlyInExpected: string[][];
  /** Up to SAMPLE_SIZE rows only in the actual file (human-readable) */
  sampleOnlyInActual:   string[][];
}

export interface WorkbookDiffSummary {
  fileExpected:     string;
  fileActual:       string;
  runAt:            string;
  sheetsOnlyInExpected: string[];
  sheetsOnlyInActual:   string[];
  sheets:           SheetDiffSummary[];
  /** STRICT — every sheet identical, no sheets missing on either side. */
  identical:        boolean;
  /** RELAXED — every sheet's data matches, no sheets missing on either side.
   *  This is what report-parity gates PASS/FAIL on by default. */
  dataIdentical:    boolean;
}

// ── Internal constants ────────────────────────────────────────────────────────

const SAMPLE_SIZE = 800; // max sample rows to keep per differing visual's VALUE diff
// A visual that's entirely missing on one side (or one of several ambiguous
// variants) doesn't need hundreds of rows reproduced in the Differences
// sheet to make the point — the Visual Comparison sheet already has its full
// row count. A short excerpt is plenty and keeps the sheet skimmable.
const MISSING_VISUAL_SAMPLE_CAP = 20;

// ── Cell normalisation ────────────────────────────────────────────────────────

// Only plain decimals (e.g. "18544.000001"), never bare integers — touching
// integers risks corrupting IDs like "007" that must keep their exact text.
const PLAIN_DECIMAL_RE = /^-?\d+\.\d+$/;
const NUMERIC_ROUND_DECIMALS = 4;
const LENIENT_TEXT_COMPARE = (process.env['CYGNUS_COMPARE_TEXT_LENIENT'] ?? '').trim() === '1';

/**
 * Collapses floating-point noise in decimal-looking text (e.g. a DAX SUM that
 * lands on "18543.999999997" instead of "18544") by rounding to a fixed number
 * of decimals and trimming trailing zeros. Bare integers and non-numeric text
 * are returned untouched, so this never risks reinterpreting an ID or code.
 */
function normalizeNumericDrift(text: string): string {
  if (!PLAIN_DECIMAL_RE.test(text)) return text;
  const n = Number(text);
  if (!Number.isFinite(n)) return text;
  const rounded = n.toFixed(NUMERIC_ROUND_DECIMALS).replace(/0+$/, '').replace(/\.$/, '');
  return rounded;
}

/**
 * Canonical form used for equality/hash checks only.
 * When lenient mode is on, case and common separator differences are ignored.
 */
function canonicalizeForCompare(text: string): string {
  if (!LENIENT_TEXT_COMPARE) return text;
  return text
    .toLocaleLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Converts a raw ExcelJS cell value to a consistent string so that
 * formatting-only quirks (extra spaces, number vs string, float drift) don't
 * produce false mismatches.
 */
function normalizeCell(value: ExcelJS.CellValue): string {
  let raw: string;
  if (value === null || value === undefined) {
    raw = '';
  } else if (value instanceof Date) {
    raw = value.toISOString();
  } else if (typeof value === 'object') {
    // RichText, hyperlink, formula result etc.
    const v = value as any;
    if (v.result !== undefined) raw = String(v.result).trim();
    else if (v.text !== undefined) raw = String(v.text).trim();
    else raw = JSON.stringify(value);
  } else {
    raw = String(value).trim();
  }
  return normalizeNumericDrift(raw);
}

/**
 * Normalises a row's cell value array into a plain string array.
 * row.values is 1-indexed in ExcelJS (index 0 is undefined).
 */
function normalizeRow(row: ExcelJS.Row): string[] {
  const vals = row.values as ExcelJS.CellValue[];
  // slice(1) to drop the 0-index placeholder
  return (vals || []).slice(1).map(normalizeCell);
}

// ── Hashing ───────────────────────────────────────────────────────────────────

function hashRow(cells: string[]): string {
  return crypto.createHash('sha1').update(JSON.stringify(cells.map(canonicalizeForCompare))).digest('hex');
}

function hashRunning(existing: string, rowHash: string): string {
  return crypto.createHash('sha1').update(existing + rowHash).digest('hex');
}

// ── Tier 1: whole-sheet ordered hash ──────────────────────────────────────────

interface SheetHashResult {
  orderedHash: string;
  rowCount:    number;
}

/**
 * One retry for a failed streaming read. The failure mode this covers was
 * observed in production: the ExcelJS reader ending its iteration early
 * without emitting every sheet. A sheet not being emitted now THROWS (see the
 * `found` checks in the readers below) instead of silently returning an empty
 * result that would then be "compared" as if it were real data.
 */
async function withSheetRetry<T>(filePath: string, target: SheetInfo, read: () => Promise<T>): Promise<T> {
  try {
    return await read();
  } catch (e) {
    console.warn(`[excel-compare] Re-reading "${target.name}" from ${path.basename(filePath)} after a failed streaming read: ${(e as Error).message}`);
    return await read();
  }
}

function sheetNotEmitted(filePath: string, target: SheetInfo): Error {
  return new Error(
    `Sheet "${target.name}" (worksheets/sheet${target.fileNo}.xml) exists in ${path.basename(filePath)}'s ` +
    `workbook.xml but was never emitted by the streaming reader — the file may be corrupted or the read ` +
    `was interrupted. Refusing to treat it as empty.`,
  );
}

async function hashSheetOrdered(filePath: string, target: SheetInfo, allSheets: SheetInfo[]): Promise<SheetHashResult> {
  const realNames = new Set(allSheets.map(s => s.name));
  return withSheetRetry(filePath, target, async () => {
    let orderedHash = '';
    let rowCount    = 0;
    let found       = false;

    const reader = makeSeededReader(filePath, allSheets);
    for await (const ws of reader) {
      if (!isTargetSheet(ws, target, realNames)) {
        for await (const _ of ws) { /* drain so the stream doesn't stall */ }
        continue;
      }
      found = true;
      for await (const row of ws) {
        const cells = normalizeRow(row as ExcelJS.Row);
        if (cells.every(c => c === '')) continue; // skip fully-blank spacer rows
        rowCount++;
        orderedHash = hashRunning(orderedHash, hashRow(cells));
      }
    }
    if (!found) throw sheetNotEmitted(filePath, target);
    return { orderedHash, rowCount };
  });
}

// ── Tier 2: visual-block parsing ──────────────────────────────────────────────
//
// report-export.helpers.ts (writeVisualsToSheet) writes each visual as:
//   [VISUAL] <title> [<type>]   ← marker row, column A only
//   <header1> <header2> ...     ← header row
//   <data rows...>
//   (blank spacer row)
// This re-groups the flat row stream back into those blocks. The regex is
// greedy on the title so it always splits on the LAST "[...]" in the line,
// which is where writeVisualsToSheet always puts the type — safe even if a
// visual's title itself happens to contain brackets.
const VISUAL_MARKER_RE = /^\[VISUAL\]\s+(.+)\s\[([^[\]]+)\]$/;

interface VisualBlockRaw {
  title:   string;
  type:    string;
  headers: string[];
  rows:    string[][];
}

async function streamSheetBlocks(filePath: string, target: SheetInfo, allSheets: SheetInfo[]): Promise<VisualBlockRaw[]> {
  const realNames = new Set(allSheets.map(s => s.name));
  return withSheetRetry(filePath, target, async () => {
    const blocks: VisualBlockRaw[] = [];
    let current: VisualBlockRaw | null = null;
    let sawHeaderForCurrent = false;
    let found = false;

    const reader = makeSeededReader(filePath, allSheets);
    for await (const ws of reader) {
      if (!isTargetSheet(ws, target, realNames)) {
        for await (const _ of ws) { /* drain */ }
        continue;
      }
      found = true;
      for await (const row of ws) {
        const cells = normalizeRow(row as ExcelJS.Row);
        const isBlank = cells.every(c => c === '');
        const marker  = !isBlank ? cells[0].match(VISUAL_MARKER_RE) : null;

        if (marker) {
          if (current) blocks.push(current);
          current = { title: marker[1].trim(), type: marker[2].trim(), headers: [], rows: [] };
          sawHeaderForCurrent = false;
          continue;
        }
        if (isBlank) continue;       // spacer row between visuals
        if (!current) continue;      // defensive: content before any marker row

        if (!sawHeaderForCurrent) {
          current.headers = cells;
          sawHeaderForCurrent = true;
        } else {
          current.rows.push(cells);
        }
      }
    }
    if (!found) throw sheetNotEmitted(filePath, target);
    if (current) blocks.push(current);
    return blocks;
  });
}

// ── Tier 2: block matching + comparison ───────────────────────────────────────

function headerKey(headers: string[]): string {
  return headers.map(canonicalizeForCompare).join('\u0001');
}

/** Order-independent content signature for a block's data rows. */
function blockDataSignature(rows: string[][]): string {
  const hashes = rows.map(hashRow).sort();
  return crypto.createHash('sha1').update(hashes.join('|')).digest('hex');
}

/**
 * Collapses blocks that are byte-for-byte identical to each other (same
 * headers, same data, any row order) into one representative + a count.
 * This is what lets a slicer that was innocently duplicated on the report
 * canvas — two visual objects, same title, same bound field, same values —
 * be treated as "1 visual, ×2 copies" instead of contaminating the diff.
 */
function dedupeBlocks(blocks: VisualBlockRaw[]): { representative: VisualBlockRaw; count: number }[] {
  const groups: { key: string; representative: VisualBlockRaw; count: number }[] = [];
  for (const b of blocks) {
    const key = headerKey(b.headers) + '::' + blockDataSignature(b.rows);
    const existing = groups.find(g => g.key === key);
    if (existing) existing.count++;
    else groups.push({ key, representative: b, count: 1 });
  }
  return groups.map(({ representative, count }) => ({ representative, count }));
}

/**
 * If `headers` and `targetOrder` contain the exact same column names (just
 * shuffled), remaps every row so column i lines up with targetOrder[i].
 * Returns null when the header sets genuinely differ (renamed/added/removed
 * columns) — in that case realignment isn't meaningful and headers are
 * compared positionally instead.
 */
function realignByHeader(headers: string[], rows: string[][], targetOrder: string[]): string[][] | null {
  if (headers.length !== targetOrder.length) return null;
  const sortedA = headers.map(canonicalizeForCompare).sort().join('\u0001');
  const sortedB = targetOrder.map(canonicalizeForCompare).sort().join('\u0001');
  if (sortedA !== sortedB) return null;

  const indexOf = new Map<string, number>();
  for (let i = 0; i < headers.length; i++) {
    const key = canonicalizeForCompare(headers[i]);
    if (indexOf.has(key)) {
      // Ambiguous (two headers collapse to same canonical key); keep positional compare.
      return null;
    }
    indexOf.set(key, i);
  }
  return rows.map(r => targetOrder.map(h => r[indexOf.get(canonicalizeForCompare(h))!] ?? ''));
}

/** Order-independent row diff between two row sets, with true counts + capped samples. */
function diffRows(rowsExp: string[][], rowsAct: string[][]) {
  const expMap = new Map<string, { count: number; sample: string[] }>();
  for (const r of rowsExp) {
    const h = hashRow(r);
    const e = expMap.get(h);
    if (e) e.count++; else expMap.set(h, { count: 1, sample: r });
  }
  const actMap = new Map<string, { count: number; sample: string[] }>();
  for (const r of rowsAct) {
    const h = hashRow(r);
    const e = actMap.get(h);
    if (e) e.count++; else actMap.set(h, { count: 1, sample: r });
  }

  let onlyInExpectedCount = 0;
  let onlyInActualCount   = 0;
  const sampleOnlyInExpected: string[][] = [];
  const sampleOnlyInActual:   string[][] = [];

  const allHashes = new Set([...expMap.keys(), ...actMap.keys()]);
  for (const h of allHashes) {
    const ce = expMap.get(h)?.count ?? 0;
    const ca = actMap.get(h)?.count ?? 0;
    const matched = Math.min(ce, ca);
    onlyInExpectedCount += ce - matched;
    onlyInActualCount   += ca - matched;
    for (let i = 0; i < ce - matched && sampleOnlyInExpected.length < SAMPLE_SIZE; i++) {
      sampleOnlyInExpected.push(expMap.get(h)!.sample);
    }
    for (let i = 0; i < ca - matched && sampleOnlyInActual.length < SAMPLE_SIZE; i++) {
      sampleOnlyInActual.push(actMap.get(h)!.sample);
    }
  }

  return { onlyInExpectedCount, onlyInActualCount, sampleOnlyInExpected, sampleOnlyInActual };
}

/** Compares one matched (title, type) pair of representative blocks. */
function compareBlockPair(
  exp: VisualBlockRaw, expCount: number,
  act: VisualBlockRaw, actCount: number,
): VisualBlockDiff {
  const headersEqual = headerKey(exp.headers) === headerKey(act.headers);

  let actRows = act.rows;
  let headerNote: string | undefined;
  if (!headersEqual) {
    const realigned = realignByHeader(act.headers, act.rows, exp.headers);
    if (realigned) {
      actRows = realigned;
      headerNote = 'Same columns on both sides, different order — realigned by column name before comparing values.';
    } else {
      headerNote = 'Column headers differ in text between source and target.';
    }
  }

  const expSig  = blockDataSignature(exp.rows);
  const actSig  = blockDataSignature(actRows);
  const dataEqual = expSig === actSig;

  let onlyInExpectedCount = 0;
  let onlyInActualCount   = 0;
  let sampleOnlyInExpected: string[][] = [];
  let sampleOnlyInActual:   string[][] = [];
  if (!dataEqual) {
    const d = diffRows(exp.rows, actRows);
    onlyInExpectedCount = d.onlyInExpectedCount;
    onlyInActualCount   = d.onlyInActualCount;
    sampleOnlyInExpected = d.sampleOnlyInExpected;
    sampleOnlyInActual   = d.sampleOnlyInActual;
  }

  let status: VisualDiffStatus;
  if (headersEqual && dataEqual && expCount === actCount) status = 'identical';
  else if (headersEqual && dataEqual)                      status = 'duplicate-count';
  else if (dataEqual)                                       status = 'header-diff';
  else                                                       status = 'data-diff';

  const notes: string[] = [];
  if (headerNote) notes.push(headerNote);
  if (expCount !== actCount) {
    notes.push(`Source had ${expCount} identical cop${expCount === 1 ? 'y' : 'ies'} of this visual; target had ${actCount}.`);
  }

  return {
    title: exp.title, type: exp.type, status,
    countExpected: expCount, countActual: actCount,
    rowsExpected: exp.rows.length, rowsActual: act.rows.length,
    headerExpected: exp.headers, headerActual: act.headers,
    sampleOnlyInExpected, sampleOnlyInActual,
    onlyInExpectedCount, onlyInActualCount,
    note: notes.length > 0 ? notes.join(' ') : undefined,
  };
}

/**
 * Matches every visual block from the expected export against every visual
 * block from the actual export, by (title, type) identity rather than row
 * position, and returns one verdict per visual. See the module-level comment
 * for why this is more accurate than a flat row diff.
 */
function matchVisualBlocks(expectedBlocks: VisualBlockRaw[], actualBlocks: VisualBlockRaw[]): VisualBlockDiff[] {
  const results: VisualBlockDiff[] = [];
  const titleKeyToDisplay = new Map<string, string>();
  const addTitle = (t: string) => {
    const key = canonicalizeForCompare(t);
    if (!titleKeyToDisplay.has(key)) titleKeyToDisplay.set(key, t);
  };
  expectedBlocks.forEach(b => addTitle(b.title));
  actualBlocks.forEach(b => addTitle(b.title));
  const titleKeys = [...titleKeyToDisplay.keys()];

  for (const titleKey of titleKeys) {
    const title = titleKeyToDisplay.get(titleKey)!;
    const expForTitle = expectedBlocks.filter(b => canonicalizeForCompare(b.title) === titleKey);
    const actForTitle = actualBlocks.filter(b => canonicalizeForCompare(b.title) === titleKey);
    const typeKeyToDisplay = new Map<string, string>();
    const addType = (t: string) => {
      const key = canonicalizeForCompare(t);
      if (!typeKeyToDisplay.has(key)) typeKeyToDisplay.set(key, t);
    };
    expForTitle.forEach(b => addType(b.type));
    actForTitle.forEach(b => addType(b.type));
    const typeKeys = [...typeKeyToDisplay.keys()];

    for (const typeKey of typeKeys) {
      const type = typeKeyToDisplay.get(typeKey)!;
      const expList = expForTitle.filter(b => canonicalizeForCompare(b.type) === typeKey);
      const actList = actForTitle.filter(b => canonicalizeForCompare(b.type) === typeKey);

      if (expList.length === 0) {
        for (const { representative, count } of dedupeBlocks(actList)) {
          results.push({
            title, type, status: 'only-in-actual',
            countExpected: 0, countActual: count,
            rowsExpected: 0, rowsActual: representative.rows.length,
            headerExpected: [], headerActual: representative.headers,
            sampleOnlyInExpected: [], sampleOnlyInActual: representative.rows.slice(0, MISSING_VISUAL_SAMPLE_CAP),
            onlyInExpectedCount: 0, onlyInActualCount: representative.rows.length * count,
            note: count > 1 ? `${count} identical copies found, only in target.` : 'Only in target — not present in source.',
          });
        }
        continue;
      }
      if (actList.length === 0) {
        for (const { representative, count } of dedupeBlocks(expList)) {
          results.push({
            title, type, status: 'only-in-expected',
            countExpected: count, countActual: 0,
            rowsExpected: representative.rows.length, rowsActual: 0,
            headerExpected: representative.headers, headerActual: [],
            sampleOnlyInExpected: representative.rows.slice(0, MISSING_VISUAL_SAMPLE_CAP), sampleOnlyInActual: [],
            onlyInExpectedCount: representative.rows.length * count, onlyInActualCount: 0,
            note: count > 1 ? `${count} identical copies found, only in source.` : 'Only in source — missing from target.',
          });
        }
        continue;
      }

      // Both sides have at least one block for this (title, type). Dedupe
      // exact duplicates within each side first — that's what turns "2
      // identical copies vs 1" into a clean 1:1 comparison instead of noise.
      const expDedup = dedupeBlocks(expList);
      const actDedup = dedupeBlocks(actList);

      if (expDedup.length === 1 && actDedup.length === 1) {
        results.push(compareBlockPair(
          expDedup[0].representative, expDedup[0].count,
          actDedup[0].representative, actDedup[0].count,
        ));
        continue;
      }

      // More than one DISTINCT variant remains on at least one side after
      // dedup (e.g. two genuinely different visuals share this title). Auto-
      // pairing here would be a guess, so report it plainly instead of
      // silently picking a pairing that might be wrong.
      results.push({
        title, type, status: 'ambiguous',
        countExpected: expList.length, countActual: actList.length,
        rowsExpected: expList.reduce((s, b) => s + b.rows.length, 0),
        rowsActual:   actList.reduce((s, b) => s + b.rows.length, 0),
        headerExpected: expDedup[0]?.representative.headers ?? [],
        headerActual:   actDedup[0]?.representative.headers ?? [],
        sampleOnlyInExpected: expDedup.flatMap(d => d.representative.rows.slice(0, 5)),
        sampleOnlyInActual:   actDedup.flatMap(d => d.representative.rows.slice(0, 5)),
        onlyInExpectedCount: expList.reduce((s, b) => s + b.rows.length, 0),
        onlyInActualCount:   actList.reduce((s, b) => s + b.rows.length, 0),
        note: `Found ${expDedup.length} distinct variant(s) of "${title} [${type}]" in source and ${actDedup.length} in target — ` +
              `could not unambiguously match them to each other. Open both files and check this visual manually.`,
      });
    }
  }

  return results;
}

// ── Fallback for sheets with no [VISUAL] markers at all ──────────────────────
//
// Used only when neither side's sheet parsed into any visual blocks — most
// often an arbitrary xlsx passed to the standalone compare:excel tool, not a
// Cygnus report-export.helpers.ts output. Reads every row directly (same
// blank-row skipping as Tier 1) and diffs them as one plain block, so a
// non-Cygnus-format sheet still gets a real comparison instead of a vacuous
// "nothing parsed, so call it identical".

async function streamPlainRows(filePath: string, target: SheetInfo, allSheets: SheetInfo[]): Promise<string[][]> {
  const realNames = new Set(allSheets.map(s => s.name));
  return withSheetRetry(filePath, target, async () => {
    const rows: string[][] = [];
    let found = false;
    const reader = makeSeededReader(filePath, allSheets);
    for await (const ws of reader) {
      if (!isTargetSheet(ws, target, realNames)) {
        for await (const _ of ws) { /* drain */ }
        continue;
      }
      found = true;
      for await (const row of ws) {
        const cells = normalizeRow(row as ExcelJS.Row);
        if (cells.every(c => c === '')) continue;
        rows.push(cells);
      }
    }
    if (!found) throw sheetNotEmitted(filePath, target);
    return rows;
  });
}

async function comparePlainSheet(
  fileExpected: string, fileActual: string,
  expSheet: SheetInfo, actSheet: SheetInfo,
  expAll: SheetInfo[], actAll: SheetInfo[],
): Promise<VisualBlockDiff[]> {
  const [rowsExp, rowsAct] = await Promise.all([
    streamPlainRows(fileExpected, expSheet, expAll),
    streamPlainRows(fileActual,   actSheet, actAll),
  ]);

  const expSig = blockDataSignature(rowsExp);
  const actSig = blockDataSignature(rowsAct);
  const base = {
    title: '(whole sheet)', type: 'sheet',
    countExpected: 1, countActual: 1,
    rowsExpected: rowsExp.length, rowsActual: rowsAct.length,
    headerExpected: [], headerActual: [],
  };

  if (expSig === actSig) {
    return [{
      ...base, status: 'identical',
      sampleOnlyInExpected: [], sampleOnlyInActual: [],
      onlyInExpectedCount: 0, onlyInActualCount: 0,
    }];
  }

  const d = diffRows(rowsExp, rowsAct);
  return [{
    ...base, status: 'data-diff',
    sampleOnlyInExpected: d.sampleOnlyInExpected, sampleOnlyInActual: d.sampleOnlyInActual,
    onlyInExpectedCount: d.onlyInExpectedCount, onlyInActualCount: d.onlyInActualCount,
    note: 'This sheet has no [VISUAL] marker rows — compared as one plain block instead of per-visual.',
  }];
}

// ── Main public function ──────────────────────────────────────────────────────

/**
 * Compares two Excel workbooks: Tier 1 fast ordered-hash check per sheet,
 * falling through to the visual-block-aware Tier 2 comparison on any
 * mismatch. See the module-level comment for the full design.
 *
 * @param fileExpected  Path to the expected (baseline) .xlsx file
 * @param fileActual    Path to the actual (current run) .xlsx file
 */
export async function compareWorkbooks(
  fileExpected: string,
  fileActual:   string,
): Promise<WorkbookDiffSummary> {
  // Sheet lists come from each file's workbook.xml via the central directory —
  // deterministic and complete, never dependent on zip entry order (see the
  // module comment on readWorkbookSheetInfo).
  const expSheets = await readWorkbookSheetInfo(fileExpected);
  const actSheets = await readWorkbookSheetInfo(fileActual);

  // Pair case-insensitively and whitespace-trimmed: "…- Cluster" vs
  // "…- cluster" (or a trailing space) between a source and a migrated
  // report is the same page, not a missing sheet. Excel forbids two sheets
  // in ONE workbook differing only by case, so keys are unique per side.
  const actByKey = new Map(actSheets.map(s => [sheetKey(s.name), s]));
  const expKeys  = new Set(expSheets.map(s => sheetKey(s.name)));

  const sheetsOnlyInExpected = expSheets.filter(s => !actByKey.has(sheetKey(s.name))).map(s => s.name);
  const sheetsOnlyInActual   = actSheets.filter(s => !expKeys.has(sheetKey(s.name))).map(s => s.name);
  const commonPairs = expSheets
    .filter(s => actByKey.has(sheetKey(s.name)))
    .map(s => ({ exp: s, act: actByKey.get(sheetKey(s.name))! }));

  const sheetResults: SheetDiffSummary[] = [];

  const emptySheetResult = (sheet: string, inExpected: boolean, inActual: boolean): SheetDiffSummary => ({
    sheet, inExpected, inActual,
    rowsExpected: 0, rowsActual: 0,
    identical: false, dataIdentical: false,
    visuals: [], headerDiffCount: 0, structuralDiffCount: 0,
    rowsOnlyInExpected: 0, rowsOnlyInActual: 0,
    sampleOnlyInExpected: [], sampleOnlyInActual: [],
  });

  for (const s of sheetsOnlyInExpected) sheetResults.push(emptySheetResult(s, true, false));
  for (const s of sheetsOnlyInActual)   sheetResults.push(emptySheetResult(s, false, true));

  // Common sheets — Tier 1 then Tier 2. Reported under the EXPECTED side's
  // name when the two sides' names differ only by case/whitespace.
  for (const { exp, act } of commonPairs) {
    const sheetName = exp.name;
    const [expHash, actHash] = await Promise.all([
      hashSheetOrdered(fileExpected, exp, expSheets),
      hashSheetOrdered(fileActual,   act, actSheets),
    ]);

    if (expHash.orderedHash === actHash.orderedHash) {
      // Fast path: byte-identical, no need for the detailed pass.
      sheetResults.push({
        sheet: sheetName, inExpected: true, inActual: true,
        rowsExpected: expHash.rowCount, rowsActual: actHash.rowCount,
        identical: true, dataIdentical: true,
        visuals: [], headerDiffCount: 0, structuralDiffCount: 0,
        rowsOnlyInExpected: 0, rowsOnlyInActual: 0,
        sampleOnlyInExpected: [], sampleOnlyInActual: [],
      });
      continue;
    }

    // Tier 2: visual-block-aware detailed comparison.
    const [expBlocks, actBlocks] = await Promise.all([
      streamSheetBlocks(fileExpected, exp, expSheets),
      streamSheetBlocks(fileActual,   act, actSheets),
    ]);

    let visuals: VisualBlockDiff[];
    if (expBlocks.length === 0 && actBlocks.length === 0) {
      // Neither side has any [VISUAL]-marker rows — this sheet isn't in the
      // report-export.helpers.ts format (e.g. an arbitrary xlsx passed to the
      // standalone `compare:excel` tool, not a Cygnus visual export). Tier 1
      // already proved these sheets are NOT byte-identical (that's the only
      // way execution reaches here), so falling through to an empty visuals
      // list would be a vacuous truth — "nothing to compare" is not the same
      // as "everything matches". Compare the raw rows directly instead.
      visuals = await comparePlainSheet(fileExpected, fileActual, exp, act, expSheets, actSheets);
    } else {
      visuals = matchVisualBlocks(expBlocks, actBlocks);
    }

    const identical      = visuals.length > 0 ? visuals.every(v => v.status === 'identical') : true;
    const dataIdentical   = visuals.every(v =>
      v.status === 'identical' || v.status === 'header-diff' || v.status === 'duplicate-count',
    );
    const headerDiffCount = visuals.filter(v => v.status === 'header-diff').length;
    const structuralDiffCount = visuals.filter(v =>
      v.status === 'only-in-expected' || v.status === 'only-in-actual' ||
      v.status === 'ambiguous' || v.status === 'duplicate-count',
    ).length;

    // Approximate, display-only aggregate row counts — NOT the source of
    // truth for identical/dataIdentical (see field docs above).
    let rowsOnlyInExpected = 0;
    let rowsOnlyInActual   = 0;
    const sampleOnlyInExpected: string[][] = [];
    const sampleOnlyInActual:   string[][] = [];
    const pushCapped = (arr: string[][], items: string[][]) => {
      for (const r of items) {
        if (arr.length >= SAMPLE_SIZE) break;
        arr.push(r);
      }
    };
    for (const v of visuals) {
      if (v.status === 'identical' || v.status === 'duplicate-count') continue;
      if (v.status === 'header-diff') {
        rowsOnlyInExpected += 1;
        rowsOnlyInActual   += 1;
        pushCapped(sampleOnlyInExpected, [v.headerExpected]);
        pushCapped(sampleOnlyInActual,   [v.headerActual]);
        continue;
      }
      rowsOnlyInExpected += v.onlyInExpectedCount;
      rowsOnlyInActual   += v.onlyInActualCount;
      pushCapped(sampleOnlyInExpected, v.sampleOnlyInExpected);
      pushCapped(sampleOnlyInActual,   v.sampleOnlyInActual);
    }

    sheetResults.push({
      sheet: sheetName, inExpected: true, inActual: true,
      rowsExpected: expHash.rowCount, rowsActual: actHash.rowCount,
      identical, dataIdentical,
      visuals, headerDiffCount, structuralDiffCount,
      rowsOnlyInExpected, rowsOnlyInActual,
      sampleOnlyInExpected, sampleOnlyInActual,
    });
  }

  const noSheetsMissing = sheetsOnlyInExpected.length === 0 && sheetsOnlyInActual.length === 0;
  const identical     = noSheetsMissing && sheetResults.every(s => s.identical);
  const dataIdentical = noSheetsMissing && sheetResults.every(s => s.dataIdentical);

  return {
    fileExpected: path.basename(fileExpected),
    fileActual:   path.basename(fileActual),
    runAt:        new Date().toISOString(),
    sheetsOnlyInExpected,
    sheetsOnlyInActual,
    sheets: sheetResults,
    identical,
    dataIdentical,
  };
}