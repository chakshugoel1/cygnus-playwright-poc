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
import { getReportMetadata, generateEmbedToken, generateEmbedTokenWithSP, getClusterDetails, proxyHttpsRequest, extractEmailFromToken, extractUserTokenFromBrowser, getPbiAccessToken, type EmbedTokenResult } from './pbi-api.helpers';

// ── Types ─────────────────────────────────────────────────────────────────────

export interface ReportPage {
  name:        string;   // internal Power BI page name (used for setPage)
  displayName: string;   // human-readable name shown in the report
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

/**
 * Convenience wrapper: navigate to a page by displayName and export all visuals.
 * Finds the page whose displayName matches (case-insensitive partial match).
 *
 * @param page         Playwright page with harness loaded
 * @param displayName  Report page display name (e.g. "Manager", "Employee")
 */
export async function exportVisualsForTab(
  page: Page,
  displayName: string,
): Promise<{ page: ReportPage | null; visuals: VisualExport[] }> {
  const pages = await getReportPages(page);

  // Try exact match first, then case-insensitive partial
  const target =
    pages.find(p => p.displayName === displayName) ??
    pages.find(p => p.displayName.toLowerCase().includes(displayName.toLowerCase()));

  if (!target) {
    console.warn(`[harness] Page "${displayName}" not found. Available: ${pages.map(p => p.displayName).join(', ')}`);
    return { page: null, visuals: [] };
  }

  console.log(`[harness] Navigating to page "${target.displayName}" (${target.name})`);
  await setReportPage(page, target.name);

  const visuals = await exportCurrentPageVisuals(page);
  const dataCount = visuals.filter(v => v.rowCount > 0).length;
  console.log(`[harness] Exported ${visuals.length} visuals — ${dataCount} with data`);

  return { page: target, visuals };
}
