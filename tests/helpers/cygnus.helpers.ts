import { type Page, expect } from '@playwright/test';
import { getPocConfig } from './poc-config.helpers';

// Workspace list page - used only for auth setup
const POC = getPocConfig();

export const CYGNUS_WORKSPACE_URL = POC.workspaceUrl;

// Direct URL to the configured report
export const CYGNUS_REPORT_URL = POC.reportUrl;

/**
 * Navigates to the report and waits until the Power BI canvas has rendered
 * at least one .textRun element (meaning the report is fully painted).
 */
export async function waitForReportLoad(page: Page): Promise<void> {
  console.log(`Navigating to report: ${CYGNUS_REPORT_URL}`);
  await page.goto(CYGNUS_REPORT_URL, { waitUntil: 'domcontentloaded', timeout: 60_000 });

  // Wait for .textRun elements (Power BI renders labels first)
  await expect
    .poll(
      async () => page.locator('.textRun').count(),
      { timeout: 120_000, message: 'Power BI report did not render any .textRun elements within 120s' },
    )
    .toBeGreaterThan(0);

  console.log(`Report loaded: ${page.url()}`);
}
