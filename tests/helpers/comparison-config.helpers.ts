/**
 * comparison-config.helpers.ts
 * ============================================================================
 *  SINGLE PLACE TO CONFIGURE REPORT-PARITY (migration) VALIDATION.
 * ============================================================================
 *
 * Purpose
 * -------
 * Validate that a Power BI report migrated from IMPORT mode to DIRECT LAKE mode
 * still shows the SAME data. For each pair:
 *
 *    source  = Import mode report   → exported data becomes the EXPECTED values
 *    target  = Direct Lake report   → exported data becomes the ACTUAL values
 *
 * The two exports are compared and a summary of the differences is written.
 *
 * PLUG-AND-PLAY: to validate a different report pair, add an entry to PAIRS (or
 * edit the Cygnus one) with the two reports' identifiers, then run:
 *
 *    npm run parity                 # export source + target, then compare (default)
 *    npm run parity:source          # export ONLY the Import-mode source → expected
 *    npm run parity:target          # export ONLY the Direct-Lake target → actual, then compare
 *
 * Choose which pair is active with the PAIR env var (defaults to the first):
 *    PAIR=Cygnus npm run parity
 *
 * Nothing else in the codebase needs editing to point at a new report pair —
 * the identifiers set here are applied to the existing export pipeline at runtime.
 */

// ── One report's identity ────────────────────────────────────────────────────
export interface ReportIdentity {
  /** Azure AD tenant GUID (usually the SAME for source and target). */
  tenantId: string;
  /** Power BI workspace ("group") GUID that contains the report. */
  groupId: string;
  /** Report GUID. */
  reportId: string;
  /** Dataset (semantic model) GUID backing the report. */
  datasetId: string;
  /**
   * Optional. Full report URL. If omitted it is derived from the IDs above.
   * Only set this if your tenant needs a non-standard URL.
   */
  reportUrl?: string;
  /**
   * Optional RLS role name. Leave unset unless the report's dataset requires a
   * specific effective-identity role (the embed uses the user's own identity,
   * so this is rarely needed).
   */
  rlsRole?: string;
}

// ── A source→target pair to validate ─────────────────────────────────────────
export interface ReportPair {
  /** Friendly name, used for output folders and the PAIR env selector. */
  name: string;
  /** Import-mode report — its exported data is the EXPECTED baseline. */
  source: ReportIdentity;
  /** Direct-Lake report — its exported data is the ACTUAL result. */
  target: ReportIdentity;
  /**
   * Optional. Restrict validation to these page (tab) display names.
   * If omitted, ALL pages found in the report are exported and compared.
   */
  pages?: string[];
  /**
   * Optional. Remap page display names that were RENAMED during migration:
   *   { "Old Source Name": "New Target Name" }
   * When set, the target's page is found by the mapped name but its sheet is
   * written under the SOURCE name, so the two workbooks line up for comparison.
   * Only needed when a tab's display name differs between the two reports.
   */
  pageMap?: Record<string, string>;
}

// ── Run mode ─────────────────────────────────────────────────────────────────
export type RunMode = 'both' | 'source' | 'target';

/* =============================================================================
 *  CONFIGURE YOUR REPORT PAIRS HERE
 * =============================================================================
 * The Cygnus pair below is seeded with the current Cygnus report as the SOURCE
 * (Import mode). Fill in the TARGET (Direct Lake) identifiers, then run.
 *
 * To validate a completely different report, copy the block and change the IDs.
 */
export const PAIRS: ReportPair[] = [
  {
    name: 'Power BI Report',
    // Import mode (EXPECTED) — current Cygnus report.
    source: {
      tenantId:  '',
      groupId:   '',
      reportId:  '',
      datasetId: '',
      rlsRole:   '',
    },
    // Direct Lake (ACTUAL) — FILL THESE IN with the migrated report's IDs.
    target: {
      tenantId:  '',
      groupId:   '',
      reportId:  '',
      datasetId: '',
      rlsRole:   '',
    },
     pages: ['CXO'],   // optional: only these tabs
    // pageMap: { 'DU Head': 'DU-Head' }, // optional: renamed tabs
  },
];

interface RuntimeParityConfig {
  pairName?: string;
  source?: ReportIdentity;
  target?: ReportIdentity;
  pages?: string[];
}

function readRuntimeParityConfig(): RuntimeParityConfig | null {
  const enabled = (process.env['CYGNUS_UI_RUNTIME_OVERRIDE'] ?? '').trim() === '1';
  if (!enabled) return null;

  const fs = require('fs') as typeof import('fs');
  const path = require('path') as typeof import('path');
  const configPath = (process.env['CYGNUS_UI_RUNTIME_CONFIG_PATH'] ?? '').trim();
  if (!configPath) return null;

  const resolved = path.isAbsolute(configPath)
    ? configPath
    : path.join(process.cwd(), configPath);

  if (!fs.existsSync(resolved)) {
    throw new Error(`Runtime parity config file not found: ${resolved}`);
  }

  const raw = fs.readFileSync(resolved, 'utf8');
  const parsed = JSON.parse(raw) as RuntimeParityConfig;

  if (!parsed.source || !parsed.target) {
    throw new Error('Runtime parity config is missing source/target identities.');
  }

  return parsed;
}

function getRuntimePair(): ReportPair | null {
  const runtime = readRuntimeParityConfig();
  if (!runtime) return null;
  const source = runtime.source as ReportIdentity;
  const target = runtime.target as ReportIdentity;

  return {
    name: runtime.pairName?.trim() || 'UI Runtime Pair',
    source,
    target,
    pages: runtime.pages && runtime.pages.length > 0 ? runtime.pages : undefined,
  };
}

// ── Placeholder detection (so a half-configured pair fails clearly, early) ────
const PLACEHOLDER_RE = /^REPLACE_WITH_/i;

/**
 * A field only counts as configured if it's non-empty AND not the literal
 * REPLACE_WITH_ placeholder text. An empty string previously passed this
 * check silently (it doesn't match /^REPLACE_WITH_/), so a blanked-out
 * PAIRS entry (e.g. scrubbed before pushing to a public repo) would slip
 * past this guard and fail later with a much more confusing error deep
 * inside the embed/API layer instead of here, up front.
 */
function fieldIsConfigured(value: string): boolean {
  return value.trim().length > 0 && !PLACEHOLDER_RE.test(value);
}

export function identityIsConfigured(id: ReportIdentity): boolean {
  return (
    fieldIsConfigured(id.tenantId) &&
    fieldIsConfigured(id.groupId) &&
    fieldIsConfigured(id.reportId) &&
    fieldIsConfigured(id.datasetId)
  );
}

// ── Selectors ────────────────────────────────────────────────────────────────

/** Returns the active pair, chosen by the PAIR env var or the first entry. */
export function getActivePair(): ReportPair {
  const runtimePair = getRuntimePair();
  if (runtimePair) return runtimePair;

  const wanted = (process.env['PAIR'] ?? '').trim();
  if (wanted) {
    const found = PAIRS.find(p => p.name.toLowerCase() === wanted.toLowerCase());
    if (!found) {
      const names = PAIRS.map(p => p.name).join(', ');
      throw new Error(`PAIR="${wanted}" not found in comparison-config. Available: ${names}`);
    }
    return found;
  }
  if (PAIRS.length === 0) {
    throw new Error('No report pairs configured in comparison-config.helpers.ts (PAIRS is empty).');
  }
  return PAIRS[0];
}

/** Returns the run mode from the MODE env var (default 'both'). */
export function getRunMode(): RunMode {
  const raw = (process.env['MODE'] ?? 'both').trim().toLowerCase();
  if (raw === 'source' || raw === 'source-only') return 'source';
  if (raw === 'target' || raw === 'target-only') return 'target';
  if (raw === 'both' || raw === '') return 'both';
  console.warn(`[parity-config] Unknown MODE="${raw}" — defaulting to 'both'.`);
  return 'both';
}

/** Builds a standard report URL from the identity's IDs. */
export function buildReportUrl(id: ReportIdentity): string {
  if (id.reportUrl && id.reportUrl.trim()) return id.reportUrl.trim();
  return (
    `https://app.powerbi.com/groups/${id.groupId}/reports/${id.reportId}` +
    `?ctid=${id.tenantId}&experience=power-bi`
  );
}

/**
 * Applies a report identity to the environment so the existing pipeline
 * (getPocConfig → getReportMetadata → embed) targets THIS report.
 *
 * It sets the HIGHEST-precedence variables getPocConfig() reads, so this wins
 * regardless of what a local ~/Power_BI_report_validation_credentials/.env contains. Because
 * getPocConfig() is not cached and pbi-api reads group/report IDs fresh per
 * call, the switch takes effect immediately for the next export stage.
 */
export function applyReportIdentity(id: ReportIdentity): void {
  const reportUrl = buildReportUrl(id);
  process.env['PowerBITenantId']    = id.tenantId;
  process.env['PowerBIWorkspaceId'] = id.groupId;
  process.env['PowerBIReportId']    = id.reportId;
  process.env['PowerBIDatasetId']   = id.datasetId;
  // CYGNUS_REPORT_URL is what the harness guardrail parses and what page.goto uses.
  process.env['CYGNUS_REPORT_URL']  = reportUrl;
  if (id.rlsRole && id.rlsRole.trim()) {
    process.env['CYGNUS_RLS_ROLE'] = id.rlsRole.trim();
  }
}
