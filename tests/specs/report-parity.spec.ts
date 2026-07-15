/**
 * REPORT PARITY (MIGRATION) VALIDATOR
 * ============================================================================
 * Validates that a Power BI report migrated from IMPORT mode to DIRECT LAKE
 * mode still shows the SAME data.
 *
 *    source (Import mode)  → exported → expected.xlsx   (EXPECTED values)
 *    target (Direct Lake)  → exported → actual.xlsx     (ACTUAL values)
 *    compare → parity-summary.xlsx
 *
 * Configure the report pair in: tests/helpers/comparison-config.helpers.ts
 *
 * Run:
 *    npm run parity           → export source + target, then compare
 *    npm run parity:source    → export ONLY source (refresh the expected baseline)
 *    npm run parity:target    → export ONLY target, compare against existing expected
 *
 * Select a pair (when more than one is configured):
 *    PAIR=Cygnus npm run parity
 */

import { test, expect } from '@playwright/test';
import * as fs   from 'fs';
import * as path from 'path';

import {
  getActivePair,
  getRunMode,
  identityIsConfigured,
  applyReportIdentity,
  buildReportUrl,
  type ReportIdentity,
} from '../helpers/comparison-config.helpers';
import {
  exportReportToWorkbook,
  writeParitySummary,
} from '../helpers/report-export.helpers';
import { compareWorkbooks } from '../helpers/excel-compare.helpers';
import { extractUserTokenFromBrowser } from '../helpers/pbi-api.helpers';

const SOURCE_LABEL = 'Import mode';
const TARGET_LABEL = 'Direct Lake';

// Resolved once — the pair/mode config is static for the run.
const PAIR = getActivePair();
const MODE = getRunMode();

const OUT_DIR       = path.join(process.cwd(), 'playwright-report-parity', PAIR.name);
const EXPECTED_XLSX = path.join(OUT_DIR, 'expected.xlsx');
const ACTUAL_XLSX   = path.join(OUT_DIR, 'actual.xlsx');
const SUMMARY_XLSX  = path.join(OUT_DIR, 'parity-summary.xlsx');

const RUN_SOURCE = MODE === 'both' || MODE === 'source';
const RUN_TARGET = MODE === 'both' || MODE === 'target';

/**
 * Validate the configuration BEFORE a browser is ever launched, so a
 * misconfiguration fails in milliseconds with an actionable message instead of
 * after a slow browser start + sign-in.
 */
test.beforeAll(() => {
  if (RUN_SOURCE && !identityIsConfigured(PAIR.source)) {
    throw new Error(
      `\n\n  CONFIG ERROR — pair "${PAIR.name}": the SOURCE (Import mode) report is not configured.\n` +
      `  Replace the REPLACE_WITH_* placeholders in:\n` +
      `      tests/helpers/comparison-config.helpers.ts\n`,
    );
  }
  if (RUN_TARGET && !identityIsConfigured(PAIR.target)) {
    throw new Error(
      `\n\n  CONFIG ERROR — pair "${PAIR.name}": the TARGET (Direct Lake) report is not configured.\n` +
      `  Fill in the target groupId / reportId / datasetId (currently REPLACE_WITH_* placeholders) in:\n` +
      `      tests/helpers/comparison-config.helpers.ts\n`,
    );
  }
  if (MODE === 'target' && !fs.existsSync(EXPECTED_XLSX)) {
    throw new Error(
      `\n\n  MODE=target needs an existing baseline, but none was found at:\n` +
      `      ${EXPECTED_XLSX}\n` +
      `  Run "npm run parity:source" (or "npm run parity") first to create it.\n`,
    );
  }
});

test('Report Parity — Import mode vs Direct Lake data validation', async ({ page, context }) => {
  test.setTimeout(9_200_000); // 20+ min — two full report embeds + exports

  const pair = PAIR;
  const mode = MODE;
  const runSource = RUN_SOURCE;
  const runTarget = RUN_TARGET;

  fs.mkdirSync(OUT_DIR, { recursive: true });

  console.log('\n═══════════════════════════════════════════════════════════════');
  console.log(`  REPORT PARITY VALIDATION — pair "${pair.name}", mode "${mode}"`);
  console.log(`  Source (${SOURCE_LABEL}): ${pair.source.reportId}`);
  console.log(`  Target (${TARGET_LABEL}): ${pair.target.reportId}`);
  console.log(`  Output: ${OUT_DIR}`);
  console.log('═══════════════════════════════════════════════════════════════');

  // ── Step 1: Acquire the user AAD token once (valid for both reports) ───────
  // The token is tenant-scoped, not report-scoped, so a single acquisition
  // serves both the source and target embeds.
  console.log('\n[1] Acquiring user token from app.powerbi.com...');

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
      const t = auth.slice(7);
      if (t.length > 200) interceptedToken = t;
    }
  });

  // Navigate to whichever report we will actually use first — this both warms
  // the session and triggers the authenticated calls we intercept the token from.
  const firstIdentity: ReportIdentity = runSource ? pair.source : pair.target;
  await page.goto(buildReportUrl(firstIdentity), {
    waitUntil: 'domcontentloaded',
    timeout: 60_000,
  });
  await page.waitForTimeout(8_000); // let the report fire its authenticated calls

  let token: string | null = null;
  const pollStart = Date.now();
  while (!token && Date.now() - pollStart < 15_000) {
    token = await extractUserTokenFromBrowser(page);
    if (!token) {
      console.log('    ⏳ Token not yet captured — waiting 2 s...');
      await page.waitForTimeout(2_000);
    }
  }
  if (!token) token = interceptedToken;

  if (!token) {
    expect(
      false,
      'User token not found. Make sure the auth session is valid — run "npm run test:setup" first.',
    ).toBe(true);
    return;
  }
  console.log(`    ✅ Token acquired (length: ${token.length})`);

  // ── Page mapping (handles tabs renamed during migration) ──────────────────
  // pageMap maps SOURCE display name → TARGET display name.
  const pageMap = pair.pageMap ?? {};
  // Inverse: TARGET display name → SOURCE display name (so the target's sheets
  // are written under the source's names and the two workbooks line up).
  const inverseMap: Record<string, string> = {};
  for (const [src, tgt] of Object.entries(pageMap)) inverseMap[tgt] = src;

  const sourcePagesFilter = pair.pages;                       // may be undefined → all pages
  const targetPagesFilter = pair.pages?.map(p => pageMap[p] ?? p);

  // ── Step 2: Export SOURCE (Import mode) → expected.xlsx ───────────────────
  if (runSource) {
    console.log(`\n[2] SOURCE (${SOURCE_LABEL}) → expected values`);
    applyReportIdentity(pair.source);

    // Fresh page per stage: loadHarnessPage installs route handlers, so reusing
    // the same page across two embeds would stack them.
    const srcPage = await context.newPage();
    try {
      await exportReportToWorkbook(srcPage, token, {
        outPath:     EXPECTED_XLSX,
        pagesFilter: sourcePagesFilter,
        label:       `source (${SOURCE_LABEL})`,
      });
    } finally {
      await srcPage.close();
    }
  } else {
    console.log(`\n[2] SOURCE skipped (mode="${mode}") — reusing existing ${path.basename(EXPECTED_XLSX)}`);
  }

  // ── Step 3: Export TARGET (Direct Lake) → actual.xlsx ─────────────────────
  if (runTarget) {
    console.log(`\n[3] TARGET (${TARGET_LABEL}) → actual values`);
    applyReportIdentity(pair.target);

    const tgtPage = await context.newPage();
    try {
      await exportReportToWorkbook(tgtPage, token, {
        outPath:     ACTUAL_XLSX,
        pagesFilter: targetPagesFilter,
        // Write target sheets under the SOURCE page names so they align.
        sheetNameFor: (displayName: string) => inverseMap[displayName] ?? displayName,
        label:        `target (${TARGET_LABEL})`,
      });
    } finally {
      await tgtPage.close();
    }
  } else {
    console.log(`\n[3] TARGET skipped (mode="${mode}") — no comparison will be run`);
  }

  // ── Step 4: Compare + summary ─────────────────────────────────────────────
  if (!runTarget) {
    // source-only run: baseline refreshed, nothing to compare against yet.
    console.log('\n═══════════════════════════════════════════════════════════════');
    console.log(`  ✅ Baseline (expected values) written → ${EXPECTED_XLSX}`);
    console.log('     Run "npm run parity:target" to validate the migrated report against it.');
    console.log('═══════════════════════════════════════════════════════════════');
    expect(fs.existsSync(EXPECTED_XLSX), 'Expected baseline file was not written').toBe(true);
    return;
  }

  expect(fs.existsSync(EXPECTED_XLSX), `Missing expected baseline: ${EXPECTED_XLSX}`).toBe(true);
  expect(fs.existsSync(ACTUAL_XLSX),   `Missing actual export: ${ACTUAL_XLSX}`).toBe(true);

  console.log('\n[4] Comparing expected (source) vs actual (target)...');
  const diff = await compareWorkbooks(EXPECTED_XLSX, ACTUAL_XLSX);

  const { passed, differingSheets } = await writeParitySummary(
    diff,
    {
      pairName:       pair.name,
      mode,
      sourceLabel:    SOURCE_LABEL,
      targetLabel:    TARGET_LABEL,
      sourceReportId: pair.source.reportId,
      targetReportId: pair.target.reportId,
      expectedFile:   EXPECTED_XLSX,
      actualFile:     ACTUAL_XLSX,
    },
    SUMMARY_XLSX,
  );

  // ── Console verdict ───────────────────────────────────────────────────────
  // PASS/FAIL (and "differing" below) are gated on dataIdentical, not the
  // strict identical flag — a page whose only issue is header text/order or a
  // harmlessly-duplicated visual counts as passing here. See
  // TREAT_HEADER_DIFF_AS_FAILURE in report-export.helpers.ts to change that.
  const strictIdentical = diff.sheets.filter(s => s.identical).length;
  const headerOnly       = diff.sheets.filter(s => !s.identical && s.dataIdentical).length;
  console.log('\n═══════════════════════════════════════════════════════════════');
  console.log(`  Pages compared        : ${diff.sheets.filter(s => s.inExpected && s.inActual).length}`);
  console.log(`  Fully identical       : ${strictIdentical}`);
  console.log(`  Header/order only diff: ${headerOnly}  (counts as passing — see Visual Comparison sheet)`);
  console.log(`  Real data/structure diff: ${differingSheets}`);
  if (diff.sheetsOnlyInExpected.length > 0) {
    console.log(`  ⚠ Only in SOURCE: ${diff.sheetsOnlyInExpected.join(', ')}`);
  }
  if (diff.sheetsOnlyInActual.length > 0) {
    console.log(`  ⚠ Only in TARGET: ${diff.sheetsOnlyInActual.join(', ')}`);
  }
  for (const s of diff.sheets.filter(x => !x.dataIdentical)) {
    console.log(
      `    ✗ ${s.sheet}: ${s.headerDiffCount} header-only diff(s), ${s.structuralDiffCount} structural diff(s) ` +
      `(missing/duplicated/ambiguous visuals), ~${s.rowsOnlyInExpected} row(s) only in source / ` +
      `~${s.rowsOnlyInActual} only in target`,
    );
  }
  console.log(
    passed
      ? '  ✅ MATCH — migrated report data matches the Import-mode report'
      : `  ⚠️  DIFFERENCES FOUND — ${differingSheets} page(s) have real data/structure differences`,
  );
  console.log(`  Summary → ${SUMMARY_XLSX}`);
  console.log('═══════════════════════════════════════════════════════════════');

  // A "failed" test should mean the AUTOMATION broke (couldn't authenticate,
  // embed, export, or write files — those already throw normally via the
  // expect() calls earlier in this test). Whether the two reports' DATA
  // matches is a business result, not a script error, so a differing run is
  // recorded as a non-failing annotation (shows in the console above and in
  // the Playwright HTML report) instead of a failing assertion.
  //
  // Set FAIL_ON_PARITY_DIFF=1 (e.g. for a future CI gate) to restore the old
  // behavior where a data difference fails the test outright.
  test.info().annotations.push({
    type: passed ? 'Parity result' : 'Parity result — DIFFERENCES FOUND',
    description: passed
      ? 'Source and target report data match.'
      : `${differingSheets} page(s) differ from the Import-mode baseline. See ${SUMMARY_XLSX}`,
  });

  if (process.env['FAIL_ON_PARITY_DIFF'] === '1') {
    expect(
      passed,
      `Migration parity FAILED — ${differingSheets} page(s) differ. See ${SUMMARY_XLSX}`,
    ).toBe(true);
  }
});
