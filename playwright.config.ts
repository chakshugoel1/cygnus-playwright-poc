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
  // Power BI embed + auth is genuinely flaky (token timing, transient WABI
  // hiccups) — one retry turns a spurious failure into a pass instead of
  // dumping a ~20-min run. This is ALSO what makes `trace: 'on-first-retry'`
  // below actually fire: with retries at the default 0, that trace mode never
  // triggers and you get no traces at all, including on the failures you most
  // want to debug. 2 retries under CI (slower/noisier), 1 locally.
  retries: process.env['CI'] ? 2 : 1,
  reportSlowTests: null,
  testIgnore: ['archive-unused-files/**'],
  reporter: [
    ['html', { outputFolder: 'playwright-report' }],
  ],
  use: {
    baseURL: 'https://app.powerbi.com',
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
    // 'on' recorded continuous video for the WHOLE run — for tests that embed
    // two reports and export every page (10-20+ min each) that's a large,
    // constant CPU/disk cost on every run, pass or fail. Keep it only when a
    // test actually fails, which is the only time the video is useful anyway.
    video: 'retain-on-failure',
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
      // Pure-logic unit tests — no browser, no auth, no network. Covers the
      // logic-dense helpers (filter planning, scenario validation, slug
      // building). Fast enough to run on every push in CI. Runs standalone:
      // `npm run test:unit`.
      name: 'unit',
      testMatch: ['tests/unit/**/*.spec.ts'],
      retries: 0, // deterministic — a failure here is real, never flake
    },
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
    {
      // Lists every slicer on a report (or one page) — visualName, binding,
      // and available options — so you can write a slicer scenario JSON.
      // See tests/specs/discover-slicers.spec.ts for env vars.
      name: 'discover-slicers',
      testMatch: ['tests/specs/discover-slicers.spec.ts'],
      use: {
        storageState: path.join(SECRETS_DIR, '.auth', 'cygnus.user.json'),
        bypassCSP: true,
      },
      dependencies: ['setup'],
    },
  ],
});
