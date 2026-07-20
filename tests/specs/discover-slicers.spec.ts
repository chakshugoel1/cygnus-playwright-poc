/**
 * DISCOVER SLICERS
 * ============================================================================
 * Two things, in order:
 *   1. GLOBAL CHECK (always runs, ~instant): reads report-level filters
 *      ("Filters on all pages") via one API call — no page navigation at all.
 *      Most reports you validate use this pattern; if it finds anything, page
 *      crawling usually isn't needed for those fields at all.
 *   2. PAGE CRAWL (opt-in — only runs for pages you name): lists every slicer
 *      on each named page, including the Power BI visual `name` — the opaque
 *      internal id you need for a hierarchy field's scenario JSON entry. Flat
 *      fields don't need this at all if you already know the field from the
 *      global check or from discovering it on any ONE page — see
 *      tests/helpers/slicer-config.helpers.ts for why.
 *
 * Run:
 *    npm run discover:slicers
 *        → global check only. No pages named = no page crawl at all.
 *    DISCOVER_PAGES="Employee,Manager,CXO" npm run discover:slicers
 *        → global check + crawl exactly those 3 pages.
 *    DISCOVER_PAGE=CXO npm run discover:slicers
 *        → same as DISCOVER_PAGES, singular form, one page (back-compat).
 *    DISCOVER_ALL_PAGES=1 npm run discover:slicers
 *        → old bulk behavior: crawl every page except probable tooltip pages.
 *          Slow on a page-heavy report — see the timeout note below.
 *    DISCOVER_FIRST_MATCH=1 npm run discover:slicers
 *        → no pages named: walk pages in order (skipping likely tooltip
 *          pages) and crawl only the FIRST one that actually has slicers,
 *          then stop. Most reports repeat the same fields across pages, so
 *          one representative page is usually enough — far cheaper than
 *          DISCOVER_ALL_PAGES. Flat fields found this way still apply to
 *          every other page in a later run (targets-based, no visual needed
 *          there); a hierarchy field found here only ever applies back to
 *          THIS page, since its visual only exists here.
 *    DISCOVER_SIDE=target npm run discover:slicers
 *        → discover on the target (Direct Lake) report instead of source.
 *    DISCOVER_PAGES="Dashboard Overview" DISCOVER_FALLBACK_FIRST_MATCH=1 npm run discover:slicers
 *        → cross-report mode (used by the desktop app's cross-report filter
 *          matching, not typically set by hand): try the exact named page
 *          (case/whitespace/separator-insensitive), and if THIS report
 *          doesn't have a page by that name, silently fall back to a
 *          first-match scan instead of throwing "Page(s) not found". Only
 *          meaningful with exactly one page named.
 *    DISCOVER_SKIP_GLOBAL_CHECK=1 npm run discover:slicers
 *        → skip the global (report-level) filter check entirely. Still on
 *          by default (it's cheap and useful for reports that DO use
 *          report-level filters) — this is for when you already know YOUR
 *          report always comes back empty there and don't want to see it
 *          in the output every time. The desktop app sets this by default
 *          for its own discovery calls; the raw CLI command above still
 *          checks unless you set it yourself.
 *
 * Output:
 *    - Printed to console: global filters first, then each crawled page
 *    - Written to playwright-report-parity/<pair>/discovered-slicers.json as
 *      { globalFilters: [...], pages: { "PageName": [...] } }
 *
 * A flat slicer's `options` list is read directly. A tree/hierarchy slicer
 * (kind: "tree", e.g. a BU/DU-style filter) is reported with isHierarchy=true
 * and an empty top-level options list — its options are read step-by-step as
 * parent levels are chosen, not all at once (see discoverDependentOptions in
 * harness.helpers.ts). This test does not attempt to auto-walk a hierarchy;
 * it flags where one exists so you know to investigate it directly.
 *
 * PAGE SELECTION NOTE: this report hides its real content pages from Power
 * BI's native tab bar (it uses custom in-report navigation buttons instead),
 * so the SDK's page `visibility` flag does NOT mean "not a real page" here —
 * confirmed in practice (CXO/Manager/DU Head all report visibility=1). So
 * DISCOVER_PAGES/DISCOVER_PAGE search the FULL page list, visibility ignored;
 * only DISCOVER_ALL_PAGES's tooltip-name filter uses NAME as a signal instead,
 * which is what actually avoided the hang a genuine tooltip page caused
 * during testing (visibility could not be trusted for that either).
 */

import { test, type Page } from '@playwright/test';
import * as fs   from 'fs';
import * as path from 'path';

import { getActivePair, applyReportIdentity, buildReportUrl } from '../helpers/comparison-config.helpers';
import {
  loadHarnessPage, getReportPages, setReportPage, discoverPageSlicers, getGlobalFilters,
  type DiscoveredSlicer, type ReportPage, type GlobalFilterInfo,
} from '../helpers/harness.helpers';
import { pollForUserToken } from '../helpers/pbi-api.helpers';
import { pageNameKey } from '../helpers/report-export.helpers';

const PAIR = getActivePair();
const SIDE = ((process.env['DISCOVER_SIDE'] ?? 'source').trim().toLowerCase() === 'target') ? 'target' : 'source';

// Accept either the plural (new, GUI-facing) or singular (original) form.
const PAGES_CSV = (process.env['DISCOVER_PAGES'] ?? process.env['DISCOVER_PAGE'] ?? '').trim();
const NAMED_PAGES = PAGES_CSV ? PAGES_CSV.split(',').map(s => s.trim()).filter(Boolean) : [];
const CRAWL_ALL = (process.env['DISCOVER_ALL_PAGES'] ?? '').trim() === '1';
const FIRST_MATCH = (process.env['DISCOVER_FIRST_MATCH'] ?? '').trim() === '1';
const FALLBACK_FIRST_MATCH = (process.env['DISCOVER_FALLBACK_FIRST_MATCH'] ?? '').trim() === '1';
const INCLUDE_HIDDEN = (process.env['DISCOVER_INCLUDE_HIDDEN'] ?? '').trim() === '1';
const SKIP_GLOBAL_CHECK = (process.env['DISCOVER_SKIP_GLOBAL_CHECK'] ?? '').trim() === '1';

const OUT_DIR  = path.join(process.cwd(), 'playwright-report-parity', PAIR.name);
const OUT_JSON = path.join(OUT_DIR, 'discovered-slicers.json');

function looksLikeTooltipPage(name: string): boolean {
  return /tooltip/i.test(name);
}

/** Shared print block for one page's discovered slicers — used by both the
 * named/all-pages crawl loop and the first-match scan below. */
function logPageSlicers(slicers: DiscoveredSlicer[]): void {
  if (slicers.length === 0) {
    console.log('      (no slicers on this page)');
    return;
  }
  for (const s of slicers) {
    console.log(`      • "${s.title}"  [${s.kind}${s.isHierarchy ? ' — HIERARCHY' : ''}]`);
    console.log(`          visualName : ${s.name}`);
    console.log(`          bound to   : ${s.targetLabel}`);
    if (s.errorMessage) {
      console.log(`          ⚠ could not fully read: ${s.errorMessage}`);
    } else {
      const preview = s.options.slice(0, 8).join(', ') + (s.options.length > 8 ? `, … (${s.options.length} total)` : '');
      if (s.isHierarchy) {
        console.log(`          top-level options: ${preview || '(none found)'} (deeper levels not listed — pick a top-level value first)`);
      } else {
        console.log(`          options    : ${preview || '(none found)'}`);
      }
    }
    if (s.selected.length > 0) {
      console.log(`          currently selected: ${s.selected.join(', ')}`);
    }
  }
}

/**
 * Walks pages in declared order, skipping likely tooltip pages, and stops at
 * the first one that actually has slicers — see the module comment for why
 * one representative page is usually enough. Shared by DISCOVER_FIRST_MATCH
 * and the DISCOVER_FALLBACK_FIRST_MATCH cross-report mode below. Mutates
 * `pagesResult` in place; returns whether it found anything.
 */
async function scanForFirstPageWithSlicers(
  page: Page,
  allPages: ReportPage[],
  pagesResult: Record<string, DiscoveredSlicer[]>,
): Promise<boolean> {
  console.log('\n[4] Scanning for the first page with slicers...');
  const candidates = allPages.filter(p => !looksLikeTooltipPage(p.displayName));
  for (const p of candidates) {
    console.log(`\n[4]   checking "${p.displayName}"...`);
    await setReportPage(page, p.name);
    await page.waitForTimeout(3_000);

    const slicers = await discoverPageSlicers(page);
    if (slicers.length === 0) {
      console.log('      (no slicers here, trying next page)');
      continue;
    }

    pagesResult[p.displayName] = slicers;
    console.log(`      found ${slicers.length} slicer(s) — stopping scan.`);
    logPageSlicers(slicers);
    return true;
  }
  console.log('\n    No page with slicers was found while scanning.');
  return false;
}

/** Returns [] when neither named pages nor bulk mode was requested — global check only. */
function resolveTargetPages(allPages: ReportPage[]): ReportPage[] {
  if (NAMED_PAGES.length > 0) {
    // Explicit page request(s) — search the FULL list, ignore visibility
    // (see the module comment for why visibility can't be trusted here).
    const found: ReportPage[] = [];
    const missing: string[] = [];
    for (const wanted of NAMED_PAGES) {
      const matches = allPages.filter(p => p.displayName.trim().toLowerCase() === wanted.trim().toLowerCase());
      if (matches.length === 0) {
        missing.push(wanted);
      } else {
        if (matches.length > 1) {
          console.warn(
            `\n    ⚠ ${matches.length} pages are named "${wanted}" (internal names: ${matches.map(p => p.name).join(', ')}) — ` +
            `crawling all of them, since display name alone can't disambiguate.`,
          );
        }
        found.push(...matches);
      }
    }
    if (missing.length > 0) {
      throw new Error(`Page(s) not found: ${missing.join(', ')}. Available: ${allPages.map(p => p.displayName).join(', ')}`);
    }
    return found;
  }

  if (!CRAWL_ALL) return []; // global check only — no page crawl requested

  if (INCLUDE_HIDDEN) return allPages;
  const candidates = allPages.filter(p => !looksLikeTooltipPage(p.displayName));
  const skipped = allPages.length - candidates.length;
  if (skipped > 0) {
    console.log(`\n    (skipping ${skipped} page(s) whose name suggests a tooltip page. Set DISCOVER_INCLUDE_HIDDEN=1 to include them.)`);
  }
  return candidates;
}

test('Discover Slicers', async ({ page }) => {
  // 20 min — mainly matters for DISCOVER_ALL_PAGES; a dense page can have
  // 15-20+ slicers, and reading each one's full option list (some have 900+
  // values) via exportData() takes real time. The global check and named-page
  // modes are far faster — this ceiling is sized for the slow path.
  test.setTimeout(1_200_000);

  const identity = SIDE === 'target' ? PAIR.target : PAIR.source;
  console.log('\n═══════════════════════════════════════════════════════════════');
  console.log(`  DISCOVER SLICERS — pair "${PAIR.name}", side "${SIDE}"`);
  console.log(`  Report: ${identity.reportId}`);
  console.log('═══════════════════════════════════════════════════════════════');

  applyReportIdentity(identity);

  // Same token-acquisition pattern as report-parity.spec.ts.
  console.log('\n[1] Acquiring user token...');
  await page.goto(buildReportUrl(identity), { waitUntil: 'domcontentloaded', timeout: 60_000 });
  await page.waitForTimeout(8_000);

  const token = await pollForUserToken(page);
  if (!token) {
    throw new Error('User token not found. Run "npm run test:setup" first.');
  }
  console.log('    ✅ Token acquired');

  console.log('\n[2] Embedding report...');
  await loadHarnessPage(page, token);

  // ── Step 1: global check — always, fast, no page navigation ──────────────
  let globalFilters: GlobalFilterInfo[] | null = [];
  if (SKIP_GLOBAL_CHECK) {
    console.log('\n[3] Skipping report-level ("global") filter check (DISCOVER_SKIP_GLOBAL_CHECK=1).');
    globalFilters = null;
  } else {
    console.log('\n[3] Checking for report-level ("global") filters...');
    try {
      globalFilters = await getGlobalFilters(page);
    } catch (e) {
      console.warn(`      ⚠ could not read global filters: ${(e as Error).message}`);
    }
    if (globalFilters.length === 0) {
      console.log('      (none found — this report likely repeats fields as per-page slicers instead; see [4] below)');
    } else {
      for (const f of globalFilters) {
        const t = f.target as any;
        const label = t?.table && t?.column ? `${t.table}.${t.column}` : JSON.stringify(t);
        console.log(`      • ${label} → [${(f.values ?? ['(hierarchy selection)']).join(', ')}]`);
      }
    }
  }

  // ── Step 2: page crawl — only for pages actually requested ────────────────
  const allPages = await getReportPages(page);
  const pagesResult: Record<string, DiscoveredSlicer[]> = {};

  if (NAMED_PAGES.length === 1 && FALLBACK_FIRST_MATCH) {
    // Cross-report mode: never calls resolveTargetPages (and so never hits
    // its throw-on-missing behavior) — a missing page here is an expected,
    // handled case, not an error.
    const wanted = NAMED_PAGES[0];
    const match = allPages.find(p => pageNameKey(p.displayName) === pageNameKey(wanted));
    if (match) {
      console.log(`\n[4] Found matching page "${match.displayName}" for "${wanted}" — crawling it...`);
      await setReportPage(page, match.name);
      await page.waitForTimeout(3_000);
      const slicers = await discoverPageSlicers(page);
      pagesResult[match.displayName] = slicers;
      logPageSlicers(slicers);
    } else {
      console.log(`\n[4] Page "${wanted}" not found on this report — falling back to a first-match scan...`);
      await scanForFirstPageWithSlicers(page, allPages, pagesResult);
    }
  } else {
    const targetPages = resolveTargetPages(allPages);

    if (targetPages.length > 0) {
      for (const p of targetPages) {
        console.log(`\n[4] Page "${p.displayName}"`);
        await setReportPage(page, p.name);
        await page.waitForTimeout(3_000); // let the page settle before reading visuals

        const slicers = await discoverPageSlicers(page);
        pagesResult[p.displayName] = slicers;
        logPageSlicers(slicers);
      }
    } else if (FIRST_MATCH) {
      // No specific pages requested, but the caller wants something useful
      // without paying for a full-report crawl (DISCOVER_ALL_PAGES).
      await scanForFirstPageWithSlicers(page, allPages, pagesResult);
    } else {
      console.log('\n[4] No pages requested — skipping page crawl.');
      console.log('    Set DISCOVER_PAGES="PageA,PageB" to crawl specific pages, DISCOVER_FIRST_MATCH=1 for one representative page, or DISCOVER_ALL_PAGES=1 for every page.');
    }
  }

  const allPageNames = allPages.map(p => p.displayName);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  fs.writeFileSync(
    OUT_JSON,
    JSON.stringify({ globalFilters, pages: pagesResult, allPages: allPageNames }, null, 2),
    'utf8',
  );

  console.log('\n═══════════════════════════════════════════════════════════════');
  console.log(`  ✅ Discovery complete → ${OUT_JSON}`);
  console.log('  Copy the visualName/target values you need into a scenario JSON');
  console.log('  (see tests/helpers/slicer-config.helpers.ts for the shape).');
  console.log('═══════════════════════════════════════════════════════════════');
});