/**
 * Harness helpers — Power BI visual data extraction via powerbi-client JS API.
 *
 * Architecture:
 *   1. Playwright serves harness.html from localhost:3001 via page.route()
 *   2. A CORS proxy (page.route) forwards Power BI API calls through Node.js
 *      — The cluster resolution call (api.powerbi.com/globalservice/clusterUri)
 *        is blocked by CORS when made from localhost:3001. The proxy bypasses this.
 *      — The actual visual data calls go through the embedded iframe
 *        (app.powerbi.com origin), so those don't need proxying.
 *   3. powerbi.embed() creates an iframe → report renders
 *   4. visual.exportData() returns CSV data per visual
 */

import * as fs   from 'fs';
import * as path from 'path';
import type { Page } from '@playwright/test';
import { getReportMetadata, generateEmbedTokenWithSP, getClusterDetails, extractEmailFromToken, extractUserTokenFromBrowser, getPbiAccessToken } from './pbi-api.helpers';

// ── Types ─────────────────────────────────────────────────────────────────────

export interface ReportPage {
  name:        string;   // internal Power BI page name (used for setPage)
  displayName: string;   // human-readable name shown in the report
  /** 0 = normal/visible, 1 = hidden (tooltip pages, drillthrough targets, etc). */
  visibility?: number;
}

export interface RawVisualExport {
  name:         string;
  title:        string;
  type:         string;
  layout:       unknown;
  csvData:      string | null;
  errorMessage: string | null;
}

export interface VisualExport {
  visualName:   string;
  visualTitle:  string;
  visualType:   string;
  csvData:      string | null;
  headers:      string[];
  rows:         Record<string, string>[];
  rowCount:     number;
  errorMessage: string | null;
}

/** A slicer's binding target — either a plain column or a hierarchy level. */
export interface SlicerTarget {
  table?:          string;
  column?:         string;
  hierarchy?:      string;
  hierarchyLevel?: string;
  [key: string]:   unknown;
}

/** Raw slicer info as returned by the browser-side __listPageSlicers(). */
export interface RawSlicerInfo {
  name:         string;
  title:        string;
  type:         string;         // "slicer" | "advancedSlicer"
  kind:         'flat' | 'tree';
  targets:      SlicerTarget[];
  selected:     string[];
  errorMessage: string | null;
}

/** A discovered slicer, enriched with its full available option list. */
export interface DiscoveredSlicer {
  name:         string;         // Power BI visual name (stable id used to set state)
  title:        string;         // human-readable slicer title
  kind:         'flat' | 'tree';
  targetLabel:  string;         // e.g. "Employee.Department" — for display
  targets:      SlicerTarget[]; // raw binding data — needed for the global-filter fast path
  options:      string[];       // full available values (flat slicers only)
  selected:     string[];       // currently-selected values
  isHierarchy:  boolean;        // true for tree/BU-DU style slicers
  errorMessage: string | null;  // non-null if this slicer couldn't be read
}

// ── Harness page URL (served by Playwright via page.route — no real HTTP server) ────────────────

// Serve harness at an app.powerbi.com path so the embed iframe is same-origin.
// The Power BI iframe rejects postMessage from HTTP origins (e.g. localhost);
// serving from app.powerbi.com means parent + iframe share the same HTTPS origin,
// postMessage works, and WABI natively accepts requests from app.powerbi.com.
const HARNESS_URL = 'https://app.powerbi.com/__playwright_harness/';

// ── CSV parser ────────────────────────────────────────────────────────────────

/**
 * Parses a CSV string returned by visual.exportData() into headers + rows.
 * Power BI exportData() returns well-formed CSV with a header row.
 */
export function parseVisualCsv(csv: string): { headers: string[]; rows: Record<string, string>[] } {
  if (!csv || !csv.trim()) return { headers: [], rows: [] };

  const lines   = csv.split('\n').map(l => l.replace(/\r$/, ''));
  const nonEmpty = lines.filter(l => l.trim());
  if (nonEmpty.length < 1) return { headers: [], rows: [] };

  const parseLine = (line: string): string[] => {
    const result: string[] = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"' && !inQuotes) { inQuotes = true; continue; }
      if (ch === '"' && inQuotes && line[i + 1] === '"') { current += '"'; i++; continue; }
      if (ch === '"' && inQuotes) { inQuotes = false; continue; }
      if (ch === ',' && !inQuotes) { result.push(current); current = ''; continue; }
      current += ch;
    }
    result.push(current);
    return result;
  };

  const headers = parseLine(nonEmpty[0]);
  const rows    = nonEmpty.slice(1).map(line => {
    const values = parseLine(line);
    const row: Record<string, string> = {};
    headers.forEach((h, i) => { row[h] = values[i] ?? ''; });
    return row;
  });

  return { headers, rows };
}

// ── Core harness functions ────────────────────────────────────────────────────

/**
 * Navigates the Playwright page to a standalone harness.html (served via
 * page.route at http://localhost:3001/) and starts a fresh Power BI embed.
 *
 * Flow (mirrors AskMe CygnusPage.jsx "App Owns Data" approach):
 *   1. Fetch cluster-aware embedUrl from Power BI REST API (using user MSAL token)
 *   2. Generate a Power BI embed token via GenerateToken (user-owns-data)
 *   3. Serve harness/harness.html from disk via page.route — no real HTTP server
 *   4. Intercept globalservice cluster call and forward from Node.js with embed token
 *      (CORS blocks this from the browser; Node.js has no such restriction)
 *   5. Navigate to http://localhost:3001/ — clean page, no AMD / CSP interference
 *   6. Call window.__startEmbed(embedToken, embedUrl) — defined in harness.html
 *   7. Wait for window.__rendered === true or window.__lastError
 *
 * Prerequisites:
 *   - caller has already visited app.powerbi.com and extracted a valid user token
 *   - harness/harness.html exists at process.cwd()/harness/harness.html
 */
export async function loadHarnessPage(page: Page, token: string): Promise<void> {
  console.log(`[harness] Config report URL: ${process.env['CYGNUS_REPORT_URL'] ?? '(unset)'}`);

  // Step A: Fetch report metadata using the SP token.
  // GROUP_ID/REPORT_ID now point to the AskMe workspace; the SP has access there.
  // Fall back to the user token if the SP token fails.
  console.log('[harness] Fetching report metadata from Power BI REST API...');
  let embedUrl: string;
  let datasetId: string;
  let reportName: string;
  try {
    const spToken = await getPbiAccessToken();
    ({ embedUrl, datasetId, name: reportName } = await getReportMetadata(spToken));
    console.log('[harness] Report metadata fetched via SP token.');
  } catch {
    console.warn('[harness] SP metadata fetch failed — falling back to user token.');
    try {
      ({ embedUrl, datasetId, name: reportName } = await getReportMetadata(token));
    } catch (e) {
      console.warn(`[harness] User-token metadata fetch failed (${e}) — retrying with fresh browser token.`);
      const freshToken = await extractUserTokenFromBrowser(page);
      if (!freshToken) {
        throw e;
      }
      ({ embedUrl, datasetId, name: reportName } = await getReportMetadata(freshToken));
    }
  }
  console.log(`[harness] Report name: ${reportName}`);
  console.log(`[harness] Embed URL:  ${embedUrl}`);
  console.log(`[harness] Dataset ID: ${datasetId}`);

  // Guardrail: ensure the report metadata we fetched matches the configured report URL.
  // This makes hidden env overrides obvious and prevents embedding the wrong report silently.
  const configuredReportUrl = process.env['CYGNUS_REPORT_URL'] ?? '';
  const expectedReportId = configuredReportUrl.match(/\/reports\/([^\/?]+)/i)?.[1] ?? null;
  let actualReportId: string | null = null;
  try {
    actualReportId = new URL(embedUrl).searchParams.get('reportId');
  } catch {
    actualReportId = null;
  }
  if (expectedReportId && actualReportId && expectedReportId.toLowerCase() !== actualReportId.toLowerCase()) {
    throw new Error(
      `[harness] Report mismatch: expected reportId=${expectedReportId}, metadata returned reportId=${actualReportId}`,
    );
  }

  // Step B: Generate Power BI embed token (used only for globalservice proxy auth).
  // NOTE: We now use TokenType.Aad in Step F, so this embed token is NOT used for
  // the actual embed. It is still attempted for the globalservice cluster call proxy.
  // If generation fails for any reason, we continue with empty token (the globalservice
  // proxy will fall back to the decoded cluster URL from embedUrl config).
  console.log('[harness] Generating Power BI embed token (for cluster proxy)...');
  const userEmail = extractEmailFromToken(token);
  if (userEmail) {
    console.log(`[harness] User email: ${userEmail}`);
  }

  let embedToken = '';
  try {
    if (userEmail) {
      const result = await generateEmbedTokenWithSP(userEmail, datasetId);
      embedToken = result.token;
      console.log(`[harness] SP embed token acquired (expires: ${result.expiration})`);
    }
  } catch {
    console.log('[harness] SP token skipped — will use cluster fallback from embedUrl config.');
  }

  // Step C: Serve harness.html from disk via page.route intercepting an
  //   app.powerbi.com path. This makes the parent page HTTPS + app.powerbi.com,
  //   so the embed iframe (also app.powerbi.com) is same-origin.
  //   postMessage from parent to iframe bypasses origin checks, and WABI accepts
  //   requests from the iframe’s app.powerbi.com origin natively.
  const harnessHtmlPath = path.join(process.cwd(), 'harness', 'harness.html');
  const harnessHtml     = fs.readFileSync(harnessHtmlPath, 'utf8');
  await page.route(HARNESS_URL, async (route) => {
    await route.fulfill({ status: 200, contentType: 'text/html', body: harnessHtml });
  });

  // Step D: Intercept the browser-side globalservice cluster resolution call.
  //   powerbi-client calls api.powerbi.com/powerbi/globalservice/.../clusterdetails
  //   before loading the report. From the browser (localhost:3001 origin) this is
  //   blocked by CORS. We intercept via page.route() and proxy from Node.js.
  //   Use the embed token as auth (the user MSAL token is rejected by this endpoint).
  //   Fallback: derive the cluster origin from the embedUrl returned by the REST API —
  //   the API-returned embedUrl already encodes the correct cluster host for the tenant.
  console.log('[harness] Setting up globalservice cluster interceptor...');
  await page.route('**globalservice**', async (route) => {
    const reqUrl = route.request().url();
    console.log(`[harness] Intercepted cluster call: ${reqUrl}`);
    try {
      const res = await getClusterDetails(embedToken, reqUrl);
      if (res.status === 200) {
        console.log(`[harness] Cluster OK: ${res.body.slice(0, 200)}`);
        await route.fulfill({ status: 200, contentType: 'application/json', body: res.body });
      } else {
        throw new Error(`HTTP ${res.status}: ${res.body.slice(0, 100)}`);
      }
    } catch (proxyErr) {
      // Fallback: the REST API-returned embedUrl encodes the real cluster inside
      // its 'config' base64 query parameter, e.g.:
      //   ?config=eyJjbHVzdGVyVXJsIjoiaHR0cHMuLi4ifQ==
      //   → { "clusterUrl": "https://WABI-NORTH-EUROPE-D-PRIMARY-redirect.analysis.windows.net" }
      //
      // Using new URL(embedUrl).origin returns app.powerbi.com — WRONG. That
      // host serves an HTML page for API calls, causing "Unexpected token '<'"
      // and "Could not retrieve models and explorations."
      let realCluster: string = new URL(embedUrl).origin; // safe default
      try {
        const configParam = new URL(embedUrl).searchParams.get('config');
        if (configParam) {
          const decoded = JSON.parse(Buffer.from(configParam, 'base64').toString('utf8'));
          const cluster = decoded.clusterUrl ?? decoded.ClusterUri ?? decoded.FixedClusterUri;
          if (cluster) realCluster = cluster;
        }
      } catch { /* keep default */ }

      console.log(`[harness] Proxy failed (${proxyErr}) — using decoded cluster from embedUrl config: ${realCluster}`);
      await route.fulfill({
        status: 200,
        contentType: 'application/json',
        body: JSON.stringify({
          FixedClusterUri: realCluster,
          ClusterUri:      realCluster,
          clusterUrl:      realCluster,
        }),
      });
    }
  });

  // Step D-2: WABI proxy removed.
  //   Now that the harness is served at app.powerbi.com/__playwright_harness/,
  //   the embed iframe is same-origin (also app.powerbi.com). WABI accepts
  //   requests from app.powerbi.com natively — no Node.js proxy needed.

  // Step E: Navigate to the standalone harness page.
  //   This replaces the old "inject into app.powerbi.com" approach with a clean,
  //   isolated page — exactly like AskMe's React app running at its own origin.
  console.log('[harness] Navigating to standalone harness page...');
  await page.goto(HARNESS_URL);
  await page.waitForLoadState('domcontentloaded');
  // HTML is served — unroute the handler so refreshes don't re-trigger it
  await page.unroute(HARNESS_URL);

  // Step F: Start the Power BI embed using the user's AAD token (TokenType.Aad = 1).
  //   We are now on https://app.powerbi.com so the embed is same-origin.
  //   TokenType.Aad uses the user's own identity — workspace Members bypass RLS
  //   and see all data without needing an SP-generated embed token.
  //   Passing tokenType=1 explicitly so harness.html uses Aad instead of Embed.
  console.log('[harness] Calling window.__startEmbed(userToken, embedUrl, tokenType=Aad)...');
  await page.evaluate(({ t, u }: { t: string; u: string }) => {
    void (window as any).__startEmbed(t, u, 1); // 1 = TokenType.Aad
  }, { t: token, u: embedUrl });

  // Wait for rendered or error (up to 3 minutes — matches harness.html internal timeout)
  console.log('[harness] Waiting for report to render...');
  await page.waitForFunction(
    () => (window as any).__rendered === true || (window as any).__lastError !== null,
    { timeout: 180_000 },
  );

  // Clean up cluster interceptor
  await page.unroute('**globalservice**');

  const lastError = await page.evaluate(() => (window as any).__lastError as string | null);
  if (lastError) {
    throw new Error(`Power BI embed error: ${lastError}`);
  }

  console.log('[harness] Report rendered ✅');
}

/**
 * Returns all pages (tabs) in the embedded report.
 * Use the returned `name` field with setReportPage().
 */
export async function getReportPages(page: Page): Promise<ReportPage[]> {
  const pages = await page.evaluate((): Promise<ReportPage[]> => {
    return (window as any).__getPages();
  });
  return pages;
}

/**
 * Navigates the embedded report to the specified page by internal name.
 * Get page names from getReportPages() first.
 * Waits 4 seconds after navigation for the page to settle.
 */
export async function setReportPage(page: Page, pageName: string): Promise<void> {
  await page.evaluate((name: string) => {
    return (window as any).__setPage(name);
  }, pageName);
}

/**
 * Exports data from all visuals on the currently active report page.
 * Returns structured data including parsed CSV rows.
 *
 * Note: Some visual types (e.g. images, slicers, text boxes) may not support
 * exportData() and will return an errorMessage instead.
 */
export async function exportCurrentPageVisuals(page: Page): Promise<VisualExport[]> {
  const rawResults: RawVisualExport[] = await page.evaluate((): Promise<RawVisualExport[]> => {
    return (window as any).__exportPageVisuals();
  });

  return rawResults.map(raw => {
    const { headers, rows } = raw.csvData
      ? parseVisualCsv(raw.csvData)
      : { headers: [], rows: [] };

    return {
      visualName:   raw.name,
      visualTitle:  raw.title || '(no title)',
      visualType:   raw.type,
      csvData:      raw.csvData,
      headers,
      rows,
      rowCount:     rows.length,
      errorMessage: raw.errorMessage,
    };
  });
}

// ── Slicer helpers ──────────────────────────────────────────────────────────
//
// These power both the "Discover Filters" UI step and the actual filtered
// runs. The cascading/hierarchy behaviour is handled WITHOUT any hard-coded
// knowledge of which slicer depends on which: we always read a slicer's
// options fresh, in page order, right before we need them — so a dependent
// slicer (e.g. Department under BU/DU) is naturally read only AFTER its
// parents are set, and Power BI has already narrowed its options by then.

const SLICER_SETTLE_MS     = 4000; // wait for the page to redraw after a selection
const SLICER_READ_RETRY_MS = 1500; // pause before retrying an empty options read
const SLICER_READ_PACE_MS  = 400;  // pause between consecutive slicer reads during discovery

function targetLabel(targets: SlicerTarget[]): string {
  if (!targets || targets.length === 0) return '(unknown)';
  const one = (t: SlicerTarget) => t.hierarchy
    ? `${t.table ?? '?'}.${t.hierarchy}${t.hierarchyLevel ? ' › ' + t.hierarchyLevel : ''}`
    : `${t.table ?? '?'}.${t.column ?? '?'}`;
  return targets.map(one).join('  +  ');
}

/**
 * Lists the slicers on the currently active page WITHOUT their option lists
 * (fast — just bindings + current selection). Used as the first step of
 * discovery, and to check what's selected.
 */
export async function listPageSlicers(page: Page): Promise<RawSlicerInfo[]> {
  return page.evaluate((): Promise<RawSlicerInfo[]> => {
    return (window as any).__listPageSlicers();
  });
}

/**
 * Reads a single flat slicer's full list of available option values by
 * exporting the slicer visual's own data. Returns [] for tree slicers (whose
 * options are hierarchical and read differently) or on any export error.
 */
export async function readSlicerOptions(page: Page, visualName: string): Promise<string[]> {
  const attempt = async (): Promise<string[]> => {
    const raw = await page.evaluate((name: string) => {
      return (window as any).__exportSingleVisual(name);
    }, visualName);

    if (!raw || !raw.csvData) return [];
    const { rows } = parseVisualCsv(raw.csvData as string);
    // A slicer's export is a single-column list of its values. Take the first
    // column of each row, dedupe, drop blanks.
    const seen = new Set<string>();
    const options: string[] = [];
    for (const r of rows) {
      const first = Object.values(r)[0];
      const val = (first ?? '').toString().trim();
      if (val && !seen.has(val)) { seen.add(val); options.push(val); }
    }
    return options;
  };

  const first = await attempt();
  if (first.length > 0) return first;

  // Empty on the first try is ambiguous — it might genuinely have no data, or
  // it might be a transient throttling/timing hiccup from reading many
  // slicers back-to-back (observed in practice on pages with 15+ slicers:
  // exportData() succeeds with no error but returns nothing). One retry after
  // a short pause resolves that case cheaply; a page that's genuinely empty
  // just stays empty and costs one extra short wait.
  await page.waitForTimeout(SLICER_READ_RETRY_MS);
  return attempt();
}

/**
 * Applies a value selection to a flat slicer, then waits for the page to
 * redraw. Pass the Power BI visual `name` (from listPageSlicers), not the
 * title. Selecting multiple values acts as an OR (Power BI "In" operator).
 */
export async function setSlicerSelection(
  page: Page, visualName: string, values: string[],
  isHierarchy?: boolean, targets?: SlicerTarget[],
): Promise<void> {
  await page.evaluate(({ name, vals, hier, tgts }: { name: string; vals: string[]; hier?: boolean; tgts?: SlicerTarget[] }) => {
    return (window as any).__setSlicerSelection(name, vals, hier, tgts);
  }, { name: visualName, vals: values, hier: isHierarchy, tgts: targets });
  await page.waitForTimeout(SLICER_SETTLE_MS);
}

/**
 * Full discovery for the active page: lists every slicer and, for each flat
 * slicer, reads its available options. Tree/hierarchy slicers are reported
 * with isHierarchy=true and an empty options list (the UI shows them as a
 * drill-in picker rather than a flat dropdown; deeper levels are read
 * step-by-step as parent levels are chosen, by calling setSlicerSelection on
 * the parent then readSlicerOptions on the child).
 *
 * IMPORTANT ordering note: this reads options in the slicer order returned by
 * Power BI. If a page has cascading slicers and you want a dependent slicer's
 * *narrowed* options, set the parent selection(s) first (setSlicerSelection)
 * and then call readSlicerOptions on the child — do not rely on this bulk
 * call to pre-narrow, since at discovery time nothing is selected yet.
 */
export async function discoverPageSlicers(page: Page): Promise<DiscoveredSlicer[]> {
  const raw = await listPageSlicers(page);
  const out: DiscoveredSlicer[] = [];

  for (let i = 0; i < raw.length; i++) {
    const s = raw[i];
    const isHierarchy = s.kind === 'tree';
    let options: string[] = [];
    if (!s.errorMessage) {
      try {
        // Works for hierarchy slicers too: readSlicerOptions() takes the
        // FIRST column of each exported row and dedupes — for a hierarchy
        // export that first column IS the top-level value (confirmed from a
        // real BU/DU/DU-Detail export: rows come back as
        // ('BPS','BPS',null), ('Business Enablers','Business Enablers','CSR'),
        // etc. — column 1 dedupes to exactly the top-level set). This gives
        // the top-level picker real values instead of an empty list, at no
        // extra cost — same one read either way.
        options = await readSlicerOptions(page, s.name);
      } catch (e) {
        // Non-fatal: report the slicer with no options rather than aborting
        // discovery for the whole page.
        options = [];
      }
      // Small pace between reads — reading many slicers back-to-back with zero
      // delay is what produced empty (but error-free) results on dense pages
      // in practice. Skipped after the last slicer.
      if (i < raw.length - 1) await page.waitForTimeout(SLICER_READ_PACE_MS);
    }
    if (isHierarchy && options.length === 0 && !s.errorMessage) {
      // Not necessarily wrong — see the "column 1 dedupe" note above — but a
      // hierarchy slicer reporting zero top-level options usually means either
      // a genuinely empty field or a top-level row with a blank first column,
      // which this dedupe approach can't distinguish. Flag it so a scenario
      // author investigates rather than assuming "no options" is the truth.
      console.warn(`[harness]   ⚠ hierarchy slicer "${s.title}" (${s.name}) returned 0 top-level options — verify manually before assuming this field has none.`);
    }
    out.push({
      name:         s.name,
      title:        s.title,
      kind:         s.kind,
      targetLabel:  targetLabel(s.targets),
      targets:      s.targets,
      options,
      selected:     s.selected,
      isHierarchy,
      errorMessage: s.errorMessage,
    });
  }
  return out;
}

// ── Report-level (global) filtering ─────────────────────────────────────────
//
// Applies a filter once, at the report level, instead of once per page via a
// slicer visual. Only valid for flat/Basic-filter fields — Power BI's Hierarchy
// filter type is rejected outright by report.setFilters() (confirmed against
// the real SDK source), so hierarchy fields must keep using setSlicerSelection.
// The caller (report-parity.spec.ts) decides WHEN this is eligible; these are
// just the primitive calls (setGlobalFieldFilters is defined further below,
// alongside setPageFieldFilters — both take arrays for the same reason).

/** Clears all report-level filters (used between scenarios). */
export async function clearGlobalFilters(page: Page): Promise<void> {
  await page.evaluate(() => {
    return (window as any).__clearGlobalFilters();
  });
  await page.waitForTimeout(SLICER_SETTLE_MS);
}

// ── Global filter check + field-based per-page filtering ────────────────────

/** A report-level filter as read back by getGlobalFilters(). */
export interface GlobalFilterInfo {
  target: SlicerTarget;
  values: string[] | null;
  hierarchyData: unknown | null;
  filterType: number;
}

/**
 * Fast check for report-level filters ("Filters on all pages") — one call, no
 * page navigation, works regardless of how many pages the report has. Returns
 * [] for reports that repeat the same field as an independent slicer on every
 * page instead (Cygnus's pattern) — that's not detectable this way, only by
 * reading pages individually.
 */
export async function getGlobalFilters(page: Page): Promise<GlobalFilterInfo[]> {
  return page.evaluate((): Promise<GlobalFilterInfo[]> => {
    return (window as any).__getGlobalFilters();
  });
}

/**
 * Applies Basic filters to every page in the report in one call. Takes an
 * ARRAY — always pass every field intended for this scope together, never
 * call this once per field: report.setFilters() REPLACES the whole filter
 * set each time, so a second call would silently wipe out the first.
 */
export async function setGlobalFieldFilters(page: Page, specs: Array<{ target: SlicerTarget; values: string[] }>): Promise<void> {
  await page.evaluate((s: Array<{ target: SlicerTarget; values: string[] }>) => {
    return (window as any).__setGlobalFilters(s);
  }, specs);
  await page.waitForTimeout(SLICER_SETTLE_MS);
}

/**
 * Applies Basic filters to the CURRENTLY ACTIVE page via page.setFilters() —
 * no slicer visual required, just each field's table/column. Caller must have
 * already navigated to the target page (setReportPage) before calling this.
 * Takes an ARRAY for the same reason as setGlobalFieldFilters above — always
 * pass every field-based selection for a page together, never one at a time.
 */
export async function setPageFieldFilters(page: Page, specs: Array<{ target: SlicerTarget; values: string[] }>): Promise<void> {
  await page.evaluate((s: Array<{ target: SlicerTarget; values: string[] }>) => {
    return (window as any).__setPageFieldFilters(s);
  }, specs);
  await page.waitForTimeout(SLICER_SETTLE_MS);
}