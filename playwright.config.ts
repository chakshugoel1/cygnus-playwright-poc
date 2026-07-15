import { defineConfig } from '@playwright/test';
import * as path from 'path';

// Credentials live OUTSIDE OneDrive — never synced to cloud.
// Stored at: %USERPROFILE%\.askme-poc-secrets\
const SECRETS_DIR = path.join(process.env['USERPROFILE'] ?? process.env['HOME'] ?? '', '.askme-poc-secrets');

/**
 * Playwright config for Cygnus (Power BI) automation.
 * Run with: npx playwright test
 */
export default defineConfig({
  timeout: 120_000,
  workers: 1,
  reportSlowTests: null,
  testIgnore: ['archive-unused-files/**'],
  reporter: [
    ['html', { outputFolder: 'playwright-report' }],
  ],
  use: {
    baseURL: 'https://app.powerbi.com',
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
    video: 'on',
    headless: false,
    launchOptions: {
      slowMo: 0,
      args: [
        '--window-position=0,0',
        '--window-size=1920,1080',
        '--force-device-scale-factor=1',  // bypass Windows DPI scaling
      ],
    },
    viewport: { width: 1920, height: 1080 },
  },
  projects: [
    {
      name: 'setup',
      testMatch: ['tests/setup/auth.cygnus.setup.ts'],
    },
    {
      // Standalone Excel comparison — no browser needed, configure via COMPARE_FILE_A / COMPARE_FILE_B env vars
      name: 'compare-excel',
      testMatch: ['tests/specs/excel-compare.spec.ts'],
    },
    {
      // Cygnus Power BI main run — report embed, visual data export, Excel generation
      name: 'main-run',
      testMatch: ['tests/specs/cygnus-main-run.spec.ts'],
      use: {
        storageState: path.join(SECRETS_DIR, '.auth', 'cygnus.user.json'),
        bypassCSP: true,  // Allows loading CDN scripts on app.powerbi.com
      },
      dependencies: ['setup'],
    },
    {
      // Report parity (migration) validator — exports the Import-mode source and
      // the Direct-Lake target, then compares the two data sets.
      // Configure the report pair in tests/helpers/comparison-config.helpers.ts
      name: 'report-parity',
      testMatch: ['tests/specs/report-parity.spec.ts'],
      use: {
        storageState: path.join(SECRETS_DIR, '.auth', 'cygnus.user.json'),
        bypassCSP: true,  // Allows loading CDN scripts on app.powerbi.com
      },
      dependencies: ['setup'],
    },
  ],
});
