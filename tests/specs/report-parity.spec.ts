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
  diffPageSets,
  pageNameKey,
  type ReportExportResult,
} from '../helpers/report-export.helpers';
import { compareWorkbooks } from '../helpers/excel-compare.helpers';
import { pollForUserToken, validateDatasetField } from '../helpers/pbi-api.helpers';
import {
  getScenarios,
  selectionsForPage,
  scenarioSlug,
  planFilterApplication,
  globalFiltersEnabled,
  getForPageCI,
  type SlicerScenario,
} from '../helpers/slicer-config.helpers';
import { setSlicerSelection, setGlobalFieldFilters, clearGlobalFilters, setPageFieldFilters } from '../helpers/harness.helpers';

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
      `  Open the desktop app and enter the Source report's Tenant ID / Group ID / Report ID /\n` +
      `  Dataset ID, then run Parity from there.\n` +
      `  (Advanced: a permanent default can also be set in tests/helpers/comparison-config.helpers.ts.)\n`,
    );
  }
  if (RUN_TARGET && !identityIsConfigured(PAIR.target)) {
    throw new Error(
      `\n\n  CONFIG ERROR — pair "${PAIR.name}": the TARGET (Direct Lake) report is not configured.\n` +
      `  Open the desktop app and enter the Target report's Tenant ID / Group ID / Report ID /\n` +
      `  Dataset ID, then run Parity from there.\n` +
      `  (Advanced: a permanent default can also be set in tests/helpers/comparison-config.helpers.ts.)\n`,
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

  let token: string | null = await pollForUserToken(page);
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

  // ── Scenario setup ────────────────────────────────────────────────────────
  // Each scenario applies a set of slicer selections before exporting. With no
  // scenarios configured, we run ONE synthetic "unfiltered" pass whose output
  // paths are exactly the originals — so default behaviour is byte-for-byte
  // unchanged from before slicers existed.
  const configuredScenarios = getScenarios();
  const UNFILTERED = Symbol('unfiltered');
  type ScenarioOrNone = SlicerScenario | typeof UNFILTERED;
  const scenarioRuns: ScenarioOrNone[] =
    configuredScenarios.length > 0 ? configuredScenarios : [UNFILTERED];

  if (configuredScenarios.length > 0) {
    console.log(`\n[scenarios] ${configuredScenarios.length} filter scenario(s) configured: ${configuredScenarios.map(s => s.name).join(', ')}`);
  }

  // ── Pre-flight: validate filter fields against BOTH datasets ──────────────
  // Power BI's embed SDK silently ACCEPTS a filter whose table/column doesn't
  // exist in a report's model — no error, the data just isn't filtered. When
  // source and target sit on different datasets (the normal migration case), a
  // field discovered on one side may not exist on the other, which produces a
  // filtered-vs-unfiltered comparison that LOOKS like real data differences.
  // Observed in production. So before embedding anything, probe each flat
  // filter field against both datasets via executeQueries and refuse to
  // proceed on a mismatch (CYGNUS_FILTER_FIELD_MISMATCH=skip drops the filter
  // from BOTH sides instead; =ignore keeps the old blind behaviour, with the
  // mismatch recorded as a caveat in the summary either way).
  const FIELD_MISMATCH_POLICY = (process.env['CYGNUS_FILTER_FIELD_MISMATCH'] ?? 'abort').trim().toLowerCase();
  const preflightCaveats: string[] = [];
  if (configuredScenarios.length > 0) {
    const flatFieldKey = (t: { table?: string; column?: string }) => `${t.table}${t.column}`;
    const fields = new Map<string, { table: string; column: string }>();
    let hasHierarchy = false;
    for (const sc of configuredScenarios) {
      for (const sels of Object.values(sc.pages)) {
        for (const sel of sels) {
          if (sel.isHierarchy) { hasHierarchy = true; continue; }
          const t = sel.targets?.[0];
          if (sel.targets?.length === 1 && t?.table && t?.column) {
            fields.set(flatFieldKey(t), { table: t.table, column: t.column });
          }
        }
      }
    }

    if (fields.size > 0) {
      console.log(`\n[preflight] Validating ${fields.size} filter field(s) against ${runSource && runTarget ? 'both datasets' : 'the dataset in scope'}...`);
      const sides: Array<[string, string]> = [];
      if (runSource) sides.push(['SOURCE', pair.source.datasetId]);
      if (runTarget) sides.push(['TARGET', pair.target.datasetId]);

      const missingKeys = new Set<string>();
      const missingDescriptions: string[] = [];
      const unknownSides = new Set<string>();
      for (const [key, f] of fields) {
        for (const [sideName, dsId] of sides) {
          const res = await validateDatasetField(token, dsId, f.table, f.column);
          if (res.status === 'missing') {
            missingKeys.add(key);
            missingDescriptions.push(`'${f.table}'[${f.column}] does not exist in the ${sideName} dataset (${dsId})`);
            console.warn(`    x '${f.table}'[${f.column}] - MISSING in ${sideName} dataset`);
          } else if (res.status === 'unknown') {
            if (!unknownSides.has(sideName)) {
              unknownSides.add(sideName);
              console.warn(`    ? cannot pre-validate fields against the ${sideName} dataset (${res.detail ?? 'probe failed'}) - continuing without validation for that side.`);
            }
          } else {
            console.log(`    + '${f.table}'[${f.column}] exists in ${sideName} dataset`);
          }
        }
      }
      if (hasHierarchy) {
        console.log('    (hierarchy filters cannot be pre-validated this way - a missing hierarchy visual is caught at apply time instead)');
      }

      if (missingDescriptions.length > 0) {
        if (FIELD_MISMATCH_POLICY === 'skip') {
          for (const sc of configuredScenarios) {
            for (const pageName of Object.keys(sc.pages)) {
              sc.pages[pageName] = sc.pages[pageName].filter(sel => {
                const t = sel.targets?.[0];
                return !(sel.targets?.length === 1 && t && missingKeys.has(flatFieldKey(t)));
              });
            }
          }
          for (const d of missingDescriptions) {
            preflightCaveats.push(`Filter dropped from BOTH sides (CYGNUS_FILTER_FIELD_MISMATCH=skip): ${d}`);
          }
          console.warn(`    -> ${missingKeys.size} filter field(s) dropped from BOTH sides so the comparison stays like-for-like.`);
        } else if (FIELD_MISMATCH_POLICY === 'ignore') {
          for (const d of missingDescriptions) {
            preflightCaveats.push(`Filter field mismatch IGNORED (CYGNUS_FILTER_FIELD_MISMATCH=ignore) - one side will not actually be filtered: ${d}`);
          }
        } else {
          throw new Error(
            `\n\n  FILTER FIELD MISMATCH - the comparison would be invalid (one side filtered, the other not):\n` +
            missingDescriptions.map(m => `    - ${m}`).join('\n') +
            `\n\n  Unselect these filters in the app (or discover filters on the side you are validating), then rerun.\n` +
            `  Advanced: CYGNUS_FILTER_FIELD_MISMATCH=skip drops them from BOTH sides; =ignore proceeds anyway.\n`,
          );
        }
      }
    }
  }

  // Per-scenario output paths. Unfiltered → original OUT_DIR (no subfolder).
  // Filtered → OUT_DIR/<scenario-slug>/ so each summary stays as readable as
  // the original, one folder per scenario.
  const pathsFor = (sc: ScenarioOrNone) => {
    if (sc === UNFILTERED) {
      return { dir: OUT_DIR, expected: EXPECTED_XLSX, actual: ACTUAL_XLSX, summary: SUMMARY_XLSX, label: '(unfiltered)' };
    }
    const dir = path.join(OUT_DIR, scenarioSlug(sc.name));
    return {
      dir,
      expected: path.join(dir, 'expected.xlsx'),
      actual:   path.join(dir, 'actual.xlsx'),
      summary:  path.join(dir, 'parity-summary.xlsx'),
      label:    sc.name,
    };
  };

  // Builds the per-page slicer hook for a scenario. Returns undefined for the
  // unfiltered pass so exportReportToWorkbook does no slicer work at all.
  //
  // When CYGNUS_SLICER_GLOBAL_FILTERS=1, planFilterApplication() splits the
  // scenario into (a) fields eligible to apply ONCE at the report level
  // (same flat field + same values on 2+ pages) and (b) everything else,
  // applied per-page exactly as before. This is purely a speed optimization —
  // if it's off, or a field isn't eligible, behavior is identical to before
  // this existed. If the global application itself fails for any reason, nothing
  // was removed from anyone's per-page plan in that case (see below), so the
  // proven per-page path still runs for those fields as a safety net.
  const useGlobalFilters = globalFiltersEnabled();

  // Filter-application failures per side, per scenario. A failure here means
  // that page was exported UNFILTERED on that side (report-export's safety
  // catch) — if the OTHER side applied the same filter successfully, the
  // comparison for that run is invalid, and that must surface as a caveat in
  // the summary rather than masquerade as a data difference.
  type FilterIssue = { page: string; error: string };

  const makeHooks = (sc: ScenarioOrNone, filterIssues: FilterIssue[]) => {
    if (sc === UNFILTERED) return { beforeExport: undefined, perPageHook: undefined };

    const plan = useGlobalFilters ? planFilterApplication(sc) : null;
    if (plan && plan.global.length > 0) {
      console.log(`      (global-filter fast path: ${plan.global.length} field(s) will apply once instead of per-page)`);
    }

    const beforeExport = (!plan || plan.global.length === 0) ? undefined : async (p: import('@playwright/test').Page): Promise<void> => {
      await clearGlobalFilters(p);
      // IMPORTANT: all global fields go into ONE setGlobalFieldFilters() call.
      // report.setFilters() REPLACES the whole filter set every time it's
      // called — calling it once per field would silently wipe out every
      // field applied before the last one.
      const specs = plan.global.map(g => ({ target: g.targets[0], values: g.values }));
      const labelFor = (g: typeof plan.global[number]) => g.targets.map(t => `${t.table}.${t.column}`).join('+');
      console.log(`      · GLOBAL filters (all pages): ${plan.global.map(g => `${labelFor(g)}=[${g.values.join(', ')}]`).join('; ')}`);
      try {
        await setGlobalFieldFilters(p, specs);
      } catch (e) {
        // Fall through — these fields' selections were REMOVED from the
        // per-page plan already, so without a rescue here they'd be lost
        // entirely. Re-apply ALL of them per-page as the safety net (can't
        // tell from a batch failure which single field was the problem).
        console.warn(`      ⚠ global filter batch failed — falling back to per-page for all ${plan.global.length} field(s): ${(e as Error).message}`);
        for (const g of plan.global) {
          for (const pageName of g.pages) {
            const original = selectionsForPage(sc, pageName).find(
              s => s.targets && s.targets.length === 1 &&
                   s.targets[0].table === g.targets[0].table && s.targets[0].column === g.targets[0].column,
            );
            if (original) (plan.perPage[pageName] ??= []).push(original);
          }
        }
      }
    };

    const perPageHook = async (p: import('@playwright/test').Page, pageDisplayName: string): Promise<void> => {
      try {
        await applySelectionsForPage(p, pageDisplayName);
      } catch (e) {
        filterIssues.push({ page: pageDisplayName, error: (e as Error).message });
        throw e; // report-export's catch still logs it and exports the page unfiltered
      }
    };

    const applySelectionsForPage = async (p: import('@playwright/test').Page, pageDisplayName: string): Promise<void> => {
      const selections = plan ? getForPageCI(plan.perPage, pageDisplayName) : selectionsForPage(sc, pageDisplayName);
      if (selections.length === 0) return; // "None" for this page

      // Visual-based selections (hierarchy, or explicit visualName) are each
      // an independent visual object — safe to apply one at a time in a loop.
      const visualBased = selections.filter(s => s.visualName);
      // Field-based selections (no visualName) all share ONE page-level
      // filter set — MUST be applied together in a single call. Like the
      // global path above, page.setFilters() REPLACES rather than adds, so
      // applying these one at a time would silently discard all but the last.
      const fieldBased  = selections.filter(s => !s.visualName && s.targets && s.targets.length > 0);
      const unmatched    = selections.filter(s => !s.visualName && (!s.targets || s.targets.length === 0));

      for (const sel of visualBased) {
        console.log(`      · "${sel.title ?? sel.visualName}" → [${sel.values.join(', ')}]`);
        // Passing isHierarchy/targets explicitly — confirmed in practice that
        // a fresh live getSlicerState() query can report a different target
        // count for the exact same visual between separate embed sessions,
        // so re-deriving hierarchy-ness live is not reliable; trust what
        // discovery already established instead.
        await setSlicerSelection(p, sel.visualName!, sel.values, sel.isHierarchy, sel.targets);
      }

      if (fieldBased.length > 0) {
        console.log(`      · fields on "${pageDisplayName}": ${fieldBased.map(s => `${s.title ?? '(field)'}=[${s.values.join(', ')}]`).join('; ')}`);
        // This is what lets a field discovered on ONE page apply cleanly to
        // OTHER pages that may not have their own visible slicer for it.
        await setPageFieldFilters(p, fieldBased.map(s => ({ target: s.targets![0], values: s.values })));
      }

      for (const sel of unmatched) {
        // Guarded against by validateScenarios() for JSON-sourced configs,
        // but kept here too since scenarios can also be built inline in code.
        console.warn(`      ⚠ selection "${sel.title ?? '(untitled)'}" on "${pageDisplayName}" has neither visualName nor targets — skipped.`);
      }
    };

    return { beforeExport, perPageHook };
  };


  // Aggregate verdict across all scenarios for the final annotation.
  const scenarioResults: Array<{ name: string; passed: boolean; differingSheets: number; summary: string }> = [];

  for (const sc of scenarioRuns) {
    const P = pathsFor(sc);
    fs.mkdirSync(P.dir, { recursive: true });

    if (sc !== UNFILTERED) {
      console.log('\n───────────────────────────────────────────────────────────────');
      console.log(`  SCENARIO: ${sc.name}`);
      console.log('───────────────────────────────────────────────────────────────');
    }

    // ── Export SOURCE (Import mode) → expected.xlsx ─────────────────────────
    let srcResult: ReportExportResult | null = null;
    const srcFilterIssues: { page: string; error: string }[] = [];
    if (runSource) {
      console.log(`\n[2] SOURCE (${SOURCE_LABEL}) → expected values ${P.label}`);
      applyReportIdentity(pair.source);
      const srcHooks = makeHooks(sc, srcFilterIssues); // fresh plan — independent from target's
      const srcPage = await context.newPage();
      try {
        srcResult = await exportReportToWorkbook(srcPage, token, {
          outPath:     P.expected,
          pagesFilter: sourcePagesFilter,
          label:       `source (${SOURCE_LABEL}) ${P.label}`,
          beforeExport: srcHooks.beforeExport,
          applySlicersForPage: srcHooks.perPageHook,
        });
      } finally {
        await srcPage.close();
      }
    } else {
      console.log(`\n[2] SOURCE skipped (mode="${mode}") — reusing existing ${path.basename(P.expected)}`);
    }

    // ── Export TARGET (Direct Lake) → actual.xlsx ───────────────────────────
    let tgtResult: ReportExportResult | null = null;
    const tgtFilterIssues: { page: string; error: string }[] = [];
    if (runTarget) {
      console.log(`\n[3] TARGET (${TARGET_LABEL}) → actual values ${P.label}`);
      applyReportIdentity(pair.target);
      const tgtHooks = makeHooks(sc, tgtFilterIssues); // fresh plan — independent from source's
      const tgtPage = await context.newPage();
      try {
        tgtResult = await exportReportToWorkbook(tgtPage, token, {
          outPath:      P.actual,
          pagesFilter:  targetPagesFilter,
          sheetNameFor: (displayName: string) => inverseMap[displayName] ?? displayName,
          label:        `target (${TARGET_LABEL}) ${P.label}`,
          beforeExport: tgtHooks.beforeExport,
          applySlicersForPage: tgtHooks.perPageHook,
        });
      } finally {
        await tgtPage.close();
      }
    } else {
      console.log(`\n[3] TARGET skipped (mode="${mode}") — no comparison will be run`);
    }

    // ── Source-only run: baseline refreshed, nothing to compare yet ─────────
    if (!runTarget) {
      console.log('\n═══════════════════════════════════════════════════════════════');
      console.log(`  ✅ Baseline (expected values) written → ${P.expected}`);
      console.log('     Run "npm run parity:target" to validate the migrated report against it.');
      console.log('═══════════════════════════════════════════════════════════════');
      expect(fs.existsSync(P.expected), 'Expected baseline file was not written').toBe(true);
      continue; // next scenario (or end, if unfiltered)
    }

    expect(fs.existsSync(P.expected), `Missing expected baseline: ${P.expected}`).toBe(true);
    expect(fs.existsSync(P.actual),   `Missing actual export: ${P.actual}`).toBe(true);

    // ── Comparison validity: caveats + page alignment ───────────────────────
    // Anything in `caveats` means the two exports were NOT produced under
    // identical conditions, so the comparison can't be trusted either way —
    // the summary's verdict becomes "NOT COMPARABLE" instead of PASS/FAIL.
    const caveats: string[] = [...preflightCaveats];
    for (const i of srcFilterIssues) {
      caveats.push(`Filter application FAILED on SOURCE page "${i.page}" — that page was exported UNFILTERED on the source side. (${i.error})`);
    }
    for (const i of tgtFilterIssues) {
      caveats.push(`Filter application FAILED on TARGET page "${i.page}" — that page was exported UNFILTERED on the target side. (${i.error})`);
    }
    if (srcResult && srcResult.requestedPagesMissing.length > 0) {
      caveats.push(`Requested page(s) not found in the SOURCE report (renamed or removed?): ${srcResult.requestedPagesMissing.join(', ')}`);
    }
    if (tgtResult && tgtResult.requestedPagesMissing.length > 0) {
      caveats.push(`Requested page(s) not found in the TARGET report (renamed or removed?): ${tgtResult.requestedPagesMissing.join(', ')}`);
    }

    // Full page-set alignment between the two reports (informational — extra
    // junk pages on one side don't invalidate a comparison that was scoped to
    // specific pages, but the summary should say the reports' page sets differ).
    let pageAlignment: { onlyInSource: string[]; onlyInTarget: string[]; inBothCount: number } | undefined;
    if (srcResult && tgtResult) {
      const align = diffPageSets(
        srcResult.pagesFound.map(p => p.displayName),
        tgtResult.pagesFound.map(p => p.displayName),
        pageMap,
      );
      pageAlignment = { onlyInSource: align.onlyInSource, onlyInTarget: align.onlyInTarget, inBothCount: align.inBoth.length };
      if (align.onlyInSource.length > 0 || align.onlyInTarget.length > 0) {
        console.log('\n[page-check] The two reports do NOT have identical page sets:');
        if (align.onlyInSource.length > 0) console.log(`    only in SOURCE report: ${align.onlyInSource.join(' | ')}`);
        if (align.onlyInTarget.length > 0) console.log(`    only in TARGET report: ${align.onlyInTarget.join(' | ')}`);
        console.log(`    in both: ${align.inBoth.length} page(s)`);
      }
    }

    // ── Compare + summary ───────────────────────────────────────────────────
    console.log(`\n[4] Comparing expected (source) vs actual (target) ${P.label}...`);
    const diff = await compareWorkbooks(P.expected, P.actual);

    // Read-back sanity check: every sheet we KNOW was just written must have
    // been seen by the comparator. Guards against the corrupted/partial-read
    // failure mode where a comparison quietly runs against a subset of the
    // data (observed once as sheets coming back as "Sheet1/Sheet2/Sheet3").
    const readBackCheck = (written: ReportExportResult | null, side: 'expected' | 'actual', file: string) => {
      if (!written) return;
      const seen = new Set(
        diff.sheets
          .filter(s => (side === 'expected' ? s.inExpected : s.inActual))
          .map(s => pageNameKey(s.sheet)),
      );
      for (const e of written.exported) {
        if (!seen.has(pageNameKey(e.sheetName))) {
          throw new Error(
            `Read-back mismatch: sheet "${e.sheetName}" was written to ${path.basename(file)} this run ` +
            `but the comparator did not see it. The file may be corrupted — rerun this scenario.`,
          );
        }
      }
    };
    readBackCheck(srcResult, 'expected', P.expected);
    readBackCheck(tgtResult, 'actual', P.actual);

    const { passed, differingSheets } = await writeParitySummary(
      diff,
      {
        pairName:       sc === UNFILTERED ? pair.name : `${pair.name} — ${sc.name}`,
        mode,
        sourceLabel:    SOURCE_LABEL,
        targetLabel:    TARGET_LABEL,
        sourceReportId: pair.source.reportId,
        targetReportId: pair.target.reportId,
        expectedFile:   P.expected,
        actualFile:     P.actual,
        comparisonCaveats: caveats,
        pageAlignment,
      },
      P.summary,
    );

    if (caveats.length > 0) {
      console.log('\n  ⚠⚠⚠ COMPARISON CAVEATS — this run is NOT a valid like-for-like comparison:');
      for (const c of caveats) console.log(`    - ${c}`);
    }

    const strictIdentical = diff.sheets.filter(s => s.identical).length;
    const headerOnly       = diff.sheets.filter(s => !s.identical && s.dataIdentical).length;
    console.log('\n═══════════════════════════════════════════════════════════════');
    if (sc !== UNFILTERED) console.log(`  Scenario              : ${sc.name}`);
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
    console.log(`  Summary → ${P.summary}`);
    console.log('═══════════════════════════════════════════════════════════════');

    scenarioResults.push({
      name: sc === UNFILTERED ? '(unfiltered)' : sc.name,
      passed, differingSheets, summary: P.summary,
    });
  }

  // ── Source-only run ends here (nothing was compared) ──────────────────────
  if (!runTarget) return;

  // ── Aggregate result across scenarios ─────────────────────────────────────
  // A "failed" test means the AUTOMATION broke (auth/embed/export/file-write —
  // those already throw via the expect() calls above). Whether the reports'
  // DATA matches is a business result, recorded as a non-failing annotation.
  // Set FAIL_ON_PARITY_DIFF=1 to make any data difference fail the test (CI).
  const allPassed = scenarioResults.every(r => r.passed);
  const failing   = scenarioResults.filter(r => !r.passed);

  for (const r of scenarioResults) {
    test.info().annotations.push({
      type: r.passed ? `Parity: ${r.name}` : `Parity: ${r.name} — DIFFERENCES FOUND`,
      description: r.passed
        ? 'Source and target report data match.'
        : `${r.differingSheets} page(s) differ. See ${r.summary}`,
    });
  }

  if (process.env['FAIL_ON_PARITY_DIFF'] === '1') {
    expect(
      allPassed,
      `Migration parity FAILED — ${failing.length} scenario(s) differ: ${failing.map(f => f.name).join(', ')}`,
    ).toBe(true);
  }
});