/**
 * Power BI REST API helpers — token acquisition and DAX query execution.
 *
 * Prerequisites (plug-in when IT provides service principal):
 *   ~/.askme-poc-secrets/pbi-service-principal.json
 *   {
 *     "tenantId": "8b87af7d-8647-4dc7-8df4-5f69a2011bb5",
 *     "clientId":  "<from IT>",
 *     "clientSecret": "<from IT>"
 *   }
 *
 * Usage:
 *   import { runDaxQuery, getPbiAccessToken } from './pbi-api.helpers';
 *
 *   const rows = await runDaxQuery('EVALUATE TOPN(10, DIM_EMP)');
 *   // rows = [{ '[DIM_EMP].[EMP_SK]': 1, '[DIM_EMP].[FULL_NAME]': 'Alice' }, ...]
 */

import * as https from 'https';
import * as fs   from 'fs';
import * as path from 'path';
import type { Page } from '@playwright/test';
import { getPocConfig } from './poc-config.helpers';

// ── Constants ────────────────────────────────────────────────────────────────

const POC = getPocConfig();
const TENANT_ID  = POC.tenantId;
// wks_cygnus_dev dataset (matches GROUP_ID/REPORT_ID above)
const DATASET_ID = POC.datasetId;

// RLS role name configured in the Cygnus Power BI dataset (confirmed via AskMe source code)
const RLS_ROLE   = POC.rlsRole;

const SECRETS_DIR = path.join(
  process.env['USERPROFILE'] ?? process.env['HOME'] ?? '',
  '.askme-poc-secrets',
);
const SP_FILE = path.join(SECRETS_DIR, 'pbi-service-principal.json');

const TOKEN_ENDPOINT = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const EXECUTE_QUERIES_ENDPOINT =
  `https://api.powerbi.com/v1.0/myorg/datasets/${DATASET_ID}/executeQueries`;

// ── In-memory token cache ─────────────────────────────────────────────────────

interface TokenCache {
  token: string;
  expiresAt: number; // epoch ms
}

let _cache: TokenCache | null = null;

// ── Internal: load service principal from secrets file ───────────────────────

interface ServicePrincipal {
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

function loadServicePrincipal(): ServicePrincipal {
  if (!fs.existsSync(SP_FILE)) {
    throw new Error(
      `Service principal credentials not found.\n` +
      `Expected: ${SP_FILE}\n` +
      `Create this file with: { "tenantId": "...", "clientId": "...", "clientSecret": "..." }`,
    );
  }
  const raw = fs.readFileSync(SP_FILE, 'utf-8');
  const sp  = JSON.parse(raw) as ServicePrincipal;
  if (!sp.clientId || !sp.clientSecret) {
    throw new Error(`${SP_FILE} is missing clientId or clientSecret.`);
  }
  return sp;
}

// ── Internal: HTTPS POST helper (no external deps) ───────────────────────────

function httpsPost(
  url: string,
  body: string,
  headers: Record<string, string>,
): Promise<{ status: number; body: string }> {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const req = https.request(
      {
        hostname: urlObj.hostname,
        path:     urlObj.pathname + urlObj.search,
        method:   'POST',
        headers:  { ...headers, 'Content-Length': Buffer.byteLength(body) },
      },
      (res) => {
        let data = '';
        res.on('data', (chunk) => { data += chunk; });
        res.on('end',  () => resolve({ status: res.statusCode ?? 0, body: data }));
      },
    );
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

function httpsGet(
  url: string,
  headers: Record<string, string>,
): Promise<{ status: number; body: string }> {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const req = https.request(
      {
        hostname: urlObj.hostname,
        path:     urlObj.pathname + urlObj.search,
        method:   'GET',
        headers,
      },
      (res) => {
        let data = '';
        res.on('data', (chunk) => { data += chunk; });
        res.on('end',  () => resolve({ status: res.statusCode ?? 0, body: data }));
      },
    );
    req.on('error', reject);
    req.end();
  });
}

// ── Public: get a valid Bearer token ─────────────────────────────────────────

/**
 * Returns a valid Power BI Bearer token, using an in-memory cache.
 * Fetches a new token when the cache is empty or within 2 min of expiry.
 */
export async function getPbiAccessToken(): Promise<string> {
  const now = Date.now();
  if (_cache && _cache.expiresAt - now > 120_000) {
    return _cache.token;
  }

  const sp = loadServicePrincipal();

  // Use tenantId from the secrets file — NOT the hardcoded TENANT_ID constant.
  const tokenEndpoint = `https://login.microsoftonline.com/${sp.tenantId}/oauth2/v2.0/token`;
  console.log(`[pbi-api] SP token: tenantId=${sp.tenantId} clientId=${sp.clientId}`);

  const bodyParams = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     sp.clientId,
    client_secret: sp.clientSecret,
    scope:         'https://analysis.windows.net/powerbi/api/.default',
  });

  const res = await httpsPost(
    tokenEndpoint,
    bodyParams.toString(),
    { 'Content-Type': 'application/x-www-form-urlencoded' },
  );

  if (res.status !== 200) {
    throw new Error(`Token request failed — HTTP ${res.status}:\n${res.body}`);
  }

  const parsed = JSON.parse(res.body);
  if (!parsed.access_token) {
    throw new Error(`Token response missing access_token:\n${res.body}`);
  }

  _cache = {
    token:     parsed.access_token,
    expiresAt: now + (parsed.expires_in ?? 3600) * 1000,
  };

  return _cache.token;
}

// ── Public types ──────────────────────────────────────────────────────────────

export interface DaxRow {
  [columnName: string]: string | number | boolean | null;
}

export interface DaxQueryResult {
  tableName: string;
  daxExpression: string;
  rowCount: number;
  rows: DaxRow[];
  columnNames: string[];
  errorMessage: string | null;
}

// ── Public: run a DAX query against the dataset ───────────────────────────────

/**
 * Executes a DAX expression against the Cygnus dataset and returns rows.
 *
 * @param daxExpression - e.g. `EVALUATE TOPN(10, DIM_EMP, [EMP_SK], 1)`
 * @param label         - optional label for logging/reporting
 *
 * @example
 *   const result = await runDaxQuery('EVALUATE TOPN(5, DIM_EMP)', 'DIM_EMP');
 *   console.log(result.rows); // [{ '[DIM_EMP].[EMP_SK]': 1, ... }, ...]
 */
export async function runDaxQuery(
  daxExpression: string,
  label = 'query',
): Promise<DaxQueryResult> {
  const result: DaxQueryResult = {
    tableName:     label,
    daxExpression,
    rowCount:      0,
    rows:          [],
    columnNames:   [],
    errorMessage:  null,
  };

  try {
    const token = await getPbiAccessToken();

    const requestBody = JSON.stringify({
      queries: [{ query: daxExpression }],
      serializerSettings: { includeNulls: true },
    });

    const res = await httpsPost(
      EXECUTE_QUERIES_ENDPOINT,
      requestBody,
      {
        'Content-Type':  'application/json',
        'Authorization': `Bearer ${token}`,
      },
    );

    if (res.status !== 200) {
      result.errorMessage = `HTTP ${res.status}: ${res.body.slice(0, 300)}`;
      return result;
    }

    const parsed = JSON.parse(res.body);
    const table  = parsed?.results?.[0]?.tables?.[0];

    if (!table) {
      result.errorMessage = `Unexpected response structure: ${res.body.slice(0, 300)}`;
      return result;
    }

    const rows: DaxRow[] = table.rows ?? [];
    result.rows        = rows;
    result.rowCount    = rows.length;
    result.columnNames = rows.length > 0 ? Object.keys(rows[0]) : [];

  } catch (e: any) {
    result.errorMessage = String(e);
  }

  return result;
}

// ── Public: batch query multiple tables ───────────────────────────────────────

/**
 * Queries multiple tables in sequence. Returns one DaxQueryResult per table.
 * Stops early if the first query fails with an auth error.
 *
 * @param queries - array of { dax, label } objects
 */
export async function runBatchDaxQueries(
  queries: Array<{ dax: string; label: string }>,
): Promise<DaxQueryResult[]> {
  const results: DaxQueryResult[] = [];

  for (const q of queries) {
    const r = await runDaxQuery(q.dax, q.label);
    results.push(r);

    // Stop early on auth failure — no point running more
    if (r.errorMessage?.includes('401') || r.errorMessage?.includes('403')) {
      console.warn(`[pbi-api] Auth failure on "${q.label}" — stopping batch`);
      break;
    }
  }

  return results;
}

// ── Public: get the API-correct embed URL for the report ─────────────────────
// wks_cygnus_dev workspace — the user (chakshu.goel) is a Member here, so
// TokenType.Aad embeds bypass RLS and see all data.
//
// NOTE: groupId/reportId are read FRESH from getPocConfig() inside each function
// below (not captured once at module load). This lets the report-parity flow
// switch the active report (source ↔ target) at runtime by changing the env vars
// getPocConfig() reads. For the normal Cygnus run, no override is set, so the
// values resolve to exactly the same defaults as before — behaviour unchanged.

/**
 * Fetches the report metadata from Power BI REST API and returns the embed URL.
 * Use this instead of hardcoding the embed URL — the API-returned URL contains
 * the correct cluster routing for this tenant.
 *
 * @param userToken - Bearer token from extractUserTokenFromBrowser()
 * @returns embedUrl string (e.g. "https://app.powerbi.com/reportEmbed?...")
 */
export interface ReportMetadata {
  embedUrl:  string;
  datasetId: string;
  name:      string;
}

/**
 * Returns full report metadata including the embedUrl AND the datasetId that
 * actually backs this report.  Use the returned datasetId (not the constant)
 * when building GenerateToken identities — hardcoding the wrong dataset ID
 * causes WABI to return 403 on report load even when the token is issued with 200.
 */
export async function getReportMetadata(userToken: string): Promise<ReportMetadata> {
  const { groupId: GROUP_ID, reportId: REPORT_ID } = getPocConfig();
  const url = `https://api.powerbi.com/v1.0/myorg/groups/${GROUP_ID}/reports/${REPORT_ID}`;
  const res = await httpsGet(url, { 'Authorization': `Bearer ${userToken}` });

  if (res.status !== 200) {
    console.warn(`[pbi-api] getReportMetadata failed for URL: ${url}`);
    throw new Error(`Failed to get report metadata — HTTP ${res.status}: ${res.body.slice(0, 200)}`);
  }

  const data = JSON.parse(res.body);
  if (!data.embedUrl) {
    throw new Error(`Report metadata response has no embedUrl: ${res.body.slice(0, 200)}`);
  }

  return {
    embedUrl:  data.embedUrl  as string,
    datasetId: data.datasetId as string,
    name:      data.name      as string,
  };
}

/** @deprecated Use getReportMetadata which also returns the datasetId */
export async function getReportEmbedUrl(userToken: string): Promise<string> {
  const { embedUrl } = await getReportMetadata(userToken);
  return embedUrl;
}

// ── Browser user-token extraction (works with Delegated permissions) ──────────
/**
 * Extracts the signed-in user's Power BI AAD access token directly from the
 * Playwright browser's MSAL token cache (localStorage / sessionStorage).
 *
 * Why: Service principal (client_credentials) requires Application permissions.
 *      But a real user token works with Delegated permissions — which are
 *      already configured on the app registration.
 *
 * Prerequisites:
 *   - The Playwright page must be on app.powerbi.com (authenticated via cygnus.user.json)
 *   - Delegated 'Dataset.Read.All' permission must be configured and admin-consented
 *
 * Token validity: ~1 hour. The function checks expiry and returns null if expired.
 *
 * @param page - Playwright Page already navigated to app.powerbi.com
 * @returns Bearer token string, or null if not found / expired
 */
export async function extractUserTokenFromBrowser(page: Page): Promise<string | null> {
  return await page.evaluate((): string | null => {
    const PBI_AUDIENCE = 'analysis.windows.net';

    // MSAL v2 uses both localStorage and sessionStorage depending on config
    const storages: Storage[] = [];
    try { storages.push(localStorage); }    catch { /* not available */ }
    try { storages.push(sessionStorage); } catch { /* not available */ }

    for (const storage of storages) {
      try {
        const allKeys = Object.keys(storage);

        for (const key of allKeys) {
          // MSAL access token entries always contain 'accesstoken' in the key
          if (!key.toLowerCase().includes('accesstoken')) continue;

          try {
            const raw = storage.getItem(key);
            if (!raw) continue;

            const item = JSON.parse(raw) as Record<string, string | number>;

            // The 'target' field contains the OAuth scope — check for Power BI
            const target = String(item['target'] ?? item['scope'] ?? '').toLowerCase();
            if (!target.includes(PBI_AUDIENCE)) continue;

            const secret = item['secret'] as string | undefined;
            if (!secret) continue;

            // Check token is not expired (MSAL stores expiry as Unix seconds)
            const expiresOn = Number(item['expiresOn'] ?? item['extended_expires_on'] ?? 0);
            if (expiresOn > 0 && expiresOn * 1000 < Date.now()) continue;

            return secret;
          } catch { /* skip malformed entries */ }
        }
      } catch { /* storage iteration error */ }
    }

    return null;
  });
}

/**
 * Runs a DAX query using the signed-in user's token extracted from the browser.
 * Use this when Application permissions are unavailable (Delegated-only setup).
 *
 * @param page          - Playwright Page on app.powerbi.com
 * @param daxExpression - DAX query string
 * @param label         - label for reporting
 */
export async function runDaxQueryWithUserToken(
  page: Page,
  daxExpression: string,
  label = 'query',
): Promise<DaxQueryResult> {
  const result: DaxQueryResult = {
    tableName:    label,
    daxExpression,
    rowCount:     0,
    rows:         [],
    columnNames:  [],
    errorMessage: null,
  };

  const token = await extractUserTokenFromBrowser(page);
  if (!token) {
    result.errorMessage = 'Could not extract user token from browser MSAL cache. ' +
      'Ensure the page is on app.powerbi.com and the user is logged in.';
    return result;
  }

  const requestBody = JSON.stringify({
    queries: [{ query: daxExpression }],
    serializerSettings: { includeNulls: true },
  });

  const res = await httpsPost(
    EXECUTE_QUERIES_ENDPOINT,
    requestBody,
    {
      'Content-Type':  'application/json',
      'Authorization': `Bearer ${token}`,
    },
  );

  if (res.status !== 200) {
    result.errorMessage = `HTTP ${res.status}: ${res.body.slice(0, 300)}`;
    return result;
  }

  try {
    const parsed = JSON.parse(res.body);
    const table  = parsed?.results?.[0]?.tables?.[0];
    if (table) {
      result.rows        = table.rows ?? [];
      result.rowCount    = result.rows.length;
      result.columnNames = result.rows.length > 0 ? Object.keys(result.rows[0]) : [];
    } else {
      result.errorMessage = `Unexpected response structure: ${res.body.slice(0, 200)}`;
    }
  } catch (e: any) {
    result.errorMessage = `JSON parse error: ${e}`;
  }

  return result;
}

// ── Public: decode user email from a JWT access token (Node.js, no verify) ─────

/**
 * Decodes the JWT payload of the user's Power BI access token and extracts
 * the user-principal-name (email address).  No signature verification — we
 * only need the email to populate the RLS identity in GenerateToken.
 */
export function extractEmailFromToken(jwtToken: string): string | null {
  try {
    const payloadB64 = jwtToken.split('.')[1];
    if (!payloadB64) return null;
    const json = Buffer.from(
      payloadB64.replace(/-/g, '+').replace(/_/g, '/'),
      'base64',
    ).toString('utf8');
    const payload = JSON.parse(json) as Record<string, unknown>;
    // Azure AD access tokens carry the UPN in one of these three claims
    const email =
      (payload['upn'] as string | undefined) ??
      (payload['unique_name'] as string | undefined) ??
      (payload['preferred_username'] as string | undefined) ??
      null;
    return email;
  } catch {
    return null;
  }
}

// ── Public: generate a Power BI embed token (user-owns-data scenario) ─────────

export interface EmbedTokenResult {
  token:      string;
  tokenId:    string;
  expiration: string;
}

/**
 * Generates a Power BI embed token for the Cygnus report using the signed-in
 * user's browser token (user-owns-data embed scenario).
 *
 * Why this works when TokenType.Aad does not:
 *   - TokenType.Aad requires the AAD token to be issued for OUR registered app
 *     (appid = 5089a9ca). The browser token is issued for Microsoft's own PBI
 *     web app (appid = 871c010f), which the embed API rejects.
 *   - GenerateToken uses the user token to produce a Power BI-specific embed
 *     token that is not an AAD token. The returned token works with
 *     TokenType.Embed = 0, which has no appid restrictions.
 *
 * Prerequisites:
 *   - Delegated Report.Read.All permission (already granted & consented) ✅
 *   - The user must have a Pro or Premium Per User (PPU) licence, OR the
 *     workspace must be in a Premium/Fabric capacity.
 *   - The user must have at least Viewer access to the workspace.
 *
 * @param userToken - Bearer token from extractUserTokenFromBrowser()
 * @param userEmail - UPN from extractEmailFromToken(); when provided the RLS
 *                    identity is included so the Cygnus DynamicRoles filter works
 * @returns embed token, tokenId, and expiration timestamp
 */
export async function generateEmbedToken(
  userToken: string,
  userEmail?: string | null,
): Promise<EmbedTokenResult> {
  const { groupId: GROUP_ID, reportId: REPORT_ID } = getPocConfig();
  const url  = `https://api.powerbi.com/v1.0/myorg/groups/${GROUP_ID}/reports/${REPORT_ID}/GenerateToken`;

  // Include the RLS identity if we know the user's email.
  // The Cygnus dataset requires the DynamicRoles identity; without it the
  // WABI cluster returns 403 when loading report data even though the token
  // itself is generated successfully.
  const bodyObj: Record<string, unknown> = { accessLevel: 'View' };
  if (userEmail) {
    bodyObj['identities'] = [
      {
        username: userEmail,
        roles:    [RLS_ROLE],
        datasets: [DATASET_ID],
      },
    ];
  }
  const body = JSON.stringify(bodyObj);

  const res = await httpsPost(url, body, {
    'Content-Type':  'application/json',
    'Authorization': `Bearer ${userToken}`,
  });

  if (res.status !== 200) {
    throw new Error(
      `GenerateToken failed — HTTP ${res.status}: ${res.body.slice(0, 300)}`,
    );
  }

  const data = JSON.parse(res.body);
  if (!data.token) {
    throw new Error(`GenerateToken response missing token field: ${res.body.slice(0, 200)}`);
  }

  return {
    token:      data.token      as string,
    tokenId:    data.tokenId    as string,
    expiration: data.expiration as string,
  };
}

// ── Public: generate an embed token using the Service Principal (App-Owns-Data) ─

/**
 * Generates a Power BI embed token using the Service Principal — mirroring
 * AskMe's powerbi_functions.py approach exactly.
 *
 * WHY this is needed instead of generateEmbedToken (User-Owns-Data):
 *   Microsoft's API docs require the caller to have Admin, Member, or Contributor
 *   workspace role to specify effective identities (RLS overrides) in GenerateToken.
 *   A Viewer-role user cannot override identity, so the resulting token is rejected
 *   by the WABI cluster with 403. A Service Principal with workspace Admin/Member
 *   access can always specify effective identities — this is the AskMe pattern.
 *
 * Prerequisites:
 *   - ~/.askme-poc-secrets/pbi-service-principal.json exists with credentials
 *   - The SP must be an Admin or Member of the Cygnus workspace
 *
 * @param userEmail - UPN of the signed-in user, used as the RLS effective identity
 * @returns embed token, tokenId, and expiration timestamp
 */
export async function generateEmbedTokenWithSP(
  userEmail: string,
  reportDatasetId = DATASET_ID,
): Promise<EmbedTokenResult> {
  const { groupId: GROUP_ID, reportId: REPORT_ID } = getPocConfig();
  const spToken = await getPbiAccessToken();

  // ── Check what the dataset actually requires before building the identity ──
  // Power BI exposes two flags on the dataset object:
  //   isEffectiveIdentityRequired      → a username must be supplied
  //   isEffectiveIdentityRolesRequired → a role name must also be supplied
  // If the dev workspace has NO RLS roles configured, passing DynamicRoles
  // causes WABI to return 403 at embed time even when GenerateToken returns 200.
  // We read the flags and include only what the dataset actually demands.
  let identityRequired = true;
  let rolesRequired    = true;
  try {
    const dsUrl = `https://api.powerbi.com/v1.0/myorg/groups/${GROUP_ID}/datasets/${reportDatasetId}`;
    const dsRes = await httpsGet(dsUrl, { 'Authorization': `Bearer ${spToken}` });
    if (dsRes.status === 200) {
      const ds     = JSON.parse(dsRes.body);
      identityRequired = Boolean(ds.isEffectiveIdentityRequired);
      rolesRequired    = Boolean(ds.isEffectiveIdentityRolesRequired);
      console.log(
        `[pbi-api] Dataset RLS flags: identityRequired=${identityRequired}, rolesRequired=${rolesRequired}`,
      );
    } else {
      console.warn(`[pbi-api] Could not read dataset RLS flags (HTTP ${dsRes.status}) — assuming both required`);
    }
  } catch (e) {
    console.warn(`[pbi-api] Dataset RLS check failed (${e}) — assuming both required`);
  }

  // ── Diagnostic: log email-like columns from DIM_EMP to verify username format ──
  // The RLS DynamicRoles filter uses USERNAME() which is matched against a column
  // in the dataset (e.g. [EMP_EMAIL] or [UPN]).  If the format in the dataset
  // differs from the AAD UPN we pass (e.g. @soprasteria.in vs @soprasteria.com),
  // WABI returns 403 because the filter returns zero rows for this user.
  // This block logs the first employee row so the mismatch can be spotted.
  try {
    const daxRes = await httpsPost(
      `https://api.powerbi.com/v1.0/myorg/datasets/${reportDatasetId}/executeQueries`,
      JSON.stringify({ queries: [{ query: 'EVALUATE TOPN(1, DIM_EMP)' }], serializerSettings: { includeNulls: true } }),
      { 'Content-Type': 'application/json', 'Authorization': `Bearer ${spToken}` },
    );
    if (daxRes.status === 200) {
      const rows = JSON.parse(daxRes.body)?.results?.[0]?.tables?.[0]?.rows ?? [];
      if (rows.length > 0) {
        const firstRow = rows[0] as Record<string, unknown>;
        const emailCols = Object.entries(firstRow)
          .filter(([, v]) => String(v ?? '').includes('@'))
          .map(([k, v]) => `${k} = "${v}"`);
        console.log(`[pbi-api] DIM_EMP — email-like columns in first row: ${emailCols.join(', ') || '(none)'}`);
        console.log(`[pbi-api] DIM_EMP — all columns: ${Object.keys(firstRow).join(', ')}`);
        console.log(`[pbi-api] Identity username we are sending: "${userEmail}"`);
      }
    }
  } catch { /* diagnostic only — does not affect embed */ }

  const url     = `https://api.powerbi.com/v1.0/myorg/groups/${GROUP_ID}/reports/${REPORT_ID}/GenerateToken`;
  const bodyObj: Record<string, unknown> = { accessLevel: 'View' };

  if (identityRequired) {
    const identity: Record<string, unknown> = {
      username: userEmail,
      datasets: [reportDatasetId],
    };
    if (rolesRequired) {
      identity['roles'] = [RLS_ROLE];
    }
    bodyObj['identities'] = [identity];
  }

  const body = JSON.stringify(bodyObj);

  const res = await httpsPost(url, body, {
    'Content-Type':  'application/json',
    'Authorization': `Bearer ${spToken}`,
  });

  if (res.status !== 200) {
    throw new Error(
      `GenerateToken (SP) failed — HTTP ${res.status}: ${res.body.slice(0, 300)}`,
    );
  }

  const data = JSON.parse(res.body);
  if (!data.token) {
    throw new Error(`GenerateToken (SP) response missing token field: ${res.body.slice(0, 200)}`);
  }

  return {
    token:      data.token      as string,
    tokenId:    data.tokenId    as string,
    expiration: data.expiration as string,
  };
}

// ── Public: resolve the Power BI backend cluster for this tenant ──────────────

/**
 * Fetches the cluster routing details for this tenant from the Power BI
 * globalservice API.  Called from Node.js where the user token is accepted;
 * the same call fails with 403 when made from the browser because the embed
 * token / MSAL token is rejected by this endpoint in that context.
 *
 * The raw JSON response is returned verbatim so it can be replayed via
 * page.route() to the powerbi-client SDK running inside the embed iframe.
 *
 * @param embedToken - Power BI embed token from generateEmbedToken()
 * @param url        - exact URL intercepted by page.route (e.g. https://api.powerbi.com/powerbi/globalservice/v201606/clusterdetails)
 * @returns raw JSON string from the clusterdetails endpoint
 */
export async function getClusterDetails(embedToken: string, url: string): Promise<{ status: number; body: string }> {
  return await httpsGet(url, {
    'Authorization': `Bearer ${embedToken}`,
    'Accept':        'application/json',
  });
}

// ── Public: general-purpose HTTPS proxy (forwards any method from Node.js) ──────

/**
 * Forwards an HTTP/S request from Node.js, bypassing all browser CORS and
 * mixed-content restrictions.  Used by harness.helpers to proxy WABI cluster
 * calls that the browser cannot make from http://localhost:3001.
 */
export async function proxyHttpsRequest(
  url: string,
  method: string,
  body: string,
  headers: Record<string, string>,
): Promise<{ status: number; body: string; responseHeaders: Record<string, string | string[]> }> {
  return new Promise((resolve, reject) => {
    const urlObj     = new URL(url);
    const upperMethod = method.toUpperCase();
    const hasBody    = (upperMethod === 'POST' || upperMethod === 'PUT' || upperMethod === 'PATCH') && body;
    // Clone headers and strip hop-by-hop fields Node.js must not forward
    const reqHeaders: Record<string, string | number> = { ...headers };
    delete reqHeaders['host'];
    delete reqHeaders['connection'];
    delete reqHeaders['transfer-encoding'];
    if (hasBody) reqHeaders['Content-Length'] = Buffer.byteLength(body);

    const req = https.request(
      {
        hostname: urlObj.hostname,
        path:     urlObj.pathname + urlObj.search,
        method:   upperMethod,
        headers:  reqHeaders,
      },
      (res) => {
        let data = '';
        res.on('data', (chunk) => { data += chunk; });
        res.on('end',  () => resolve({
          status:          res.statusCode ?? 0,
          body:            data,
          responseHeaders: res.headers as Record<string, string | string[]>,
        }));
      },
    );
    req.on('error', reject);
    if (hasBody) req.write(body);
    req.end();
  });
}
