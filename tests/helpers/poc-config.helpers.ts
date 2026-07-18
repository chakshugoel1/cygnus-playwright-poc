import * as path from 'path';
import * as dotenv from 'dotenv';

export interface PocConfig {
  tenantId: string;
  groupId: string;
  reportId: string;
  workspaceUrl: string;
  reportUrl: string;
  datasetId: string;
  rlsRole: string;
}

const DEFAULT_TENANT_ID = '8b87af7d-8647-4dc7-8df4-5f69a2011bb5';
const DEFAULT_GROUP_ID = '3f3f8c93-4762-459f-8b68-c36893aac01b';
const DEFAULT_REPORT_ID = '2ba857e0-e07f-4475-81c4-427cbd95cd9e';
const DEFAULT_DATASET_ID = 'fbb612a3-0d0f-43c8-b2a7-317d8d59d423';
const DEFAULT_RLS_ROLE = 'DynamicRoles';

const SECRETS_DIR = path.join(process.env['USERPROFILE'] ?? process.env['HOME'] ?? '', 'Power_BI_report_validation_credentials');
const ENV_FILE = path.join(SECRETS_DIR, '.env');

dotenv.config({ path: ENV_FILE, override: false });

function parseGroupAndReportIds(reportUrl: string): { groupId?: string; reportId?: string } {
  const match = reportUrl.match(/\/groups\/([^\/]+)\/reports\/([^\/?]+)/i);
  if (!match) return {};
  return { groupId: match[1], reportId: match[2] };
}

function buildWorkspaceUrl(groupId: string, tenantId: string): string {
  return `https://app.powerbi.com/groups/${groupId}/list?ctid=${tenantId}&chromeless=true&experience=power-bi`;
}

function buildReportUrl(groupId: string, reportId: string): string {
  return `https://app.powerbi.com/groups/${groupId}/reports/${reportId}?experience=power-bi`;
}

function firstNonEmpty(...values: Array<string | undefined>): string | undefined {
  for (const value of values) {
    const trimmed = (value ?? '').trim();
    if (trimmed) return trimmed;
  }
  return undefined;
}

export function getPocConfig(): PocConfig {
  const envReportUrl = (process.env['CYGNUS_REPORT_URL'] ?? '').trim();
  const idsFromReport = parseGroupAndReportIds(envReportUrl);

  const tenantId = firstNonEmpty(
    process.env['PowerBITenantId'],
    process.env['CYGNUS_TENANT_ID'],
    DEFAULT_TENANT_ID,
  ) as string;
  const groupId = firstNonEmpty(
    process.env['PowerBIWorkspaceId'],
    process.env['CYGNUS_GROUP_ID'],
    idsFromReport.groupId,
    DEFAULT_GROUP_ID,
  ) as string;
  const reportId = firstNonEmpty(
    process.env['PowerBIReportId'],
    process.env['CYGNUS_REPORT_ID'],
    idsFromReport.reportId,
    DEFAULT_REPORT_ID,
  ) as string;

  const workspaceUrl = (
    process.env['CYGNUS_WORKSPACE_URL'] ??
    buildWorkspaceUrl(groupId, tenantId)
  ).trim();

  const reportUrl = (envReportUrl || buildReportUrl(groupId, reportId)).trim();

  const datasetId = firstNonEmpty(
    process.env['CYGNUS_DATASET_ID'],
    process.env['PowerBIDatasetId'],
    DEFAULT_DATASET_ID,
  ) as string;

  const rlsRole = firstNonEmpty(
    process.env['CYGNUS_RLS_ROLE'],
    process.env['PowerBIRLSRole'],
    DEFAULT_RLS_ROLE,
  ) as string;

  return {
    tenantId,
    groupId,
    reportId,
    workspaceUrl,
    reportUrl,
    datasetId,
    rlsRole,
  };
}

export function getConfiguredTargetPages<T extends { name: string; displayName: string }>(
  defaultPages: T[],
): T[] {
  const raw = (process.env['CYGNUS_TARGET_PAGES_JSON'] ?? '').trim();
  if (!raw) return defaultPages;

  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return defaultPages;

    const pages = parsed
      .filter((p: any) => p && typeof p.name === 'string' && typeof p.displayName === 'string')
      .map((p: any) => ({ name: p.name, displayName: p.displayName })) as T[];

    return pages.length > 0 ? pages : defaultPages;
  } catch {
    return defaultPages;
  }
}
