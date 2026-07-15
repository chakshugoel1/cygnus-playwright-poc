import { test as setup, expect } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';
import * as dotenv from 'dotenv';
import { clickSubmit } from '../helpers/auth.helpers';

dotenv.config({ path: path.join(process.env['USERPROFILE'] ?? process.env['HOME'] ?? '', '.askme-poc-secrets', '.env') });

// Power BI auth can be slow on first load
setup.setTimeout(180_000);

const authDir  = path.join(process.env['USERPROFILE'] ?? process.env['HOME'] ?? '', '.askme-poc-secrets', '.auth');
const authFile = path.join(authDir, 'cygnus.user.json');

import { CYGNUS_WORKSPACE_URL } from '../helpers/cygnus.helpers';
const CYGNUS_URL = CYGNUS_WORKSPACE_URL;

setup('authenticate and save Cygnus session', async ({ page }) => {
  const username = process.env.E2E_USERNAME;
  const password = process.env.E2E_PASSWORD;

  if (!username || !password) {
    throw new Error('E2E_USERNAME and E2E_PASSWORD must be set in .env');
  }

  fs.mkdirSync(authDir, { recursive: true });

  // Navigate to Cygnus — Power BI may redirect to a login page
  await page.goto(CYGNUS_URL);

  // Wait until we land on either a login page or the Power BI app itself
  await page.waitForURL(
    (url) => url.href.includes('login.microsoftonline.com') || url.href.includes('powerbi.com'),
    { timeout: 30_000 }
  );

  // Allow any intermediate redirect pages (Sopra SSO loading screen, singleSignOn page) to settle
  await page.waitForTimeout(3_000);

  const currentUrl = page.url();

  // Detect any login page — covers both:
  //   login.microsoftonline.com  (Microsoft Entra)
  //   app.powerbi.com/singleSignOn  (Power BI native sign-in form)
  const isOnLoginPage =
    currentUrl.includes('login.microsoftonline.com') ||
    currentUrl.includes('singleSignOn');

  if (isOnLoginPage) {
    console.log(`➡️ Login page detected (${currentUrl.split('?')[0]}) — filling credentials`);

    // Email input — covers both Microsoft and Power BI forms
    const emailInput = page.locator(
      'input[name="loginfmt"], input[type="email"], input[placeholder*="someone@example.com" i], input[placeholder="Enter email"]'
    ).first();
    await emailInput.waitFor({ state: 'visible', timeout: 45_000 });

    await emailInput.fill(username);

    // Submit — covers both #idSIButton9 (Microsoft) and yellow Submit button (Power BI form)
    const microsoftNext = page.locator('#idSIButton9').first();
    const powerBiSubmit = page.getByRole('button', { name: /^submit$/i }).first();

    if (await microsoftNext.isVisible({ timeout: 3_000 }).catch(() => false)) {
      await clickSubmit(page);
    } else if (await powerBiSubmit.isVisible({ timeout: 3_000 }).catch(() => false)) {
      await powerBiSubmit.click();
    }

    // Wait for redirect to Microsoft after Power BI form submit
    await page.waitForURL(
      (url) => url.href.includes('login.microsoftonline.com') || url.href.includes('powerbi.com/groups'),
      { timeout: 30_000 }
    ).catch(() => {});
    await page.waitForTimeout(2_000);

    // Bad username guard
    const badUser = page.locator(`text=We couldn't find an account with that username.`).first();
    if (await badUser.isVisible({ timeout: 5_000 }).catch(() => false)) {
      throw new Error(`Invalid E2E_USERNAME: ${username}`);
    }

    // Handle second email prompt (Microsoft sometimes re-prompts)
    const secondEmail = page.locator('input[name="loginfmt"], input[type="email"]').first();
    const alreadyOnPassword = page.locator('input[name="passwd"], input[type="password"]').first();
    await page.waitForTimeout(2_000);
    if (
      (await secondEmail.isVisible().catch(() => false)) &&
      !(await alreadyOnPassword.isVisible().catch(() => false))
    ) {
      console.log('➡️ Second email prompt detected — filling again');
      await secondEmail.fill(username);
      await clickSubmit(page);
    }

    const passwordInput = page.locator('input[name="passwd"], input[type="password"]').first();
    const signInOptions = page.getByText(/sign-in options/i).first();
    const alreadyOnPowerBi = () =>
      page.url().includes('powerbi.com') &&
      !page.url().includes('login') &&
      !page.url().includes('singleSignOn');

    await expect
      .poll(
        async () =>
          alreadyOnPowerBi() ||
          (await passwordInput.isVisible().catch(() => false)) ||
          (await signInOptions.isVisible().catch(() => false)) ||
          /fido/i.test(page.url()),
        { timeout: 30_000 }
      )
      .toBeTruthy();

    if (alreadyOnPowerBi()) {
      console.log('✅ SSO completed automatically — already on Power BI, skipping password step');
    } else if (await passwordInput.isVisible().catch(() => false)) {
      console.log('✅ Standard password flow');
      await passwordInput.fill(password);
      await clickSubmit(page);
    } else {
      // Passwordless / FIDO — fall back to password option
      console.log('⚠️ Passwordless/FIDO flow');
      if (await signInOptions.isVisible().catch(() => false)) {
        await signInOptions.click();
        const usePasswordOption = page.getByText(/password/i).first();
        if (await usePasswordOption.isVisible({ timeout: 10_000 }).catch(() => false)) {
          await usePasswordOption.click();
          await passwordInput.waitFor({ state: 'visible', timeout: 15_000 });
          await passwordInput.fill(password);
          await clickSubmit(page);
        }
      }
    }

    // "Stay signed in?" prompt
    try {
      const staySignedIn = page.getByRole('button', { name: /^Yes$/i });
      if (await staySignedIn.isVisible({ timeout: 8_000 }).catch(() => false)) {
        await staySignedIn.click();
      }
    } catch {
      // not shown — acceptable
    }
  } else {
    console.log('✅ Already authenticated — no login page shown');
  }

  // Wait until Power BI report content is visible (Employee tab or report canvas)
  await expect
    .poll(
      async () => {
        const url = page.url();
        return url.includes('powerbi.com') && !url.includes('login.microsoftonline');
      },
      { timeout: 60_000 }
    )
    .toBeTruthy();

  await page.waitForLoadState('networkidle', { timeout: 30_000 }).catch(() => {});

  // Wait for MSAL to write an access token into localStorage/sessionStorage
  // before saving the auth state. On the SSO fast-path this can take several
  // seconds because Power BI triggers a silent token refresh asynchronously.
  console.log('⏳ Waiting for MSAL access token to appear in browser storage...');
  let msalReady = false;
  const msalStart = Date.now();
  while (!msalReady && Date.now() - msalStart < 25_000) {
    msalReady = await page.evaluate((): boolean => {
      const storages: Storage[] = [];
      try { storages.push(localStorage); } catch { /* ignore */ }
      try { storages.push(sessionStorage); } catch { /* ignore */ }
      for (const s of storages) {
        for (const key of Object.keys(s)) {
          if (key.toLowerCase().includes('accesstoken')) {
            const raw = s.getItem(key);
            if (!raw) continue;
            try {
              const item = JSON.parse(raw) as Record<string, unknown>;
              const target = String(item['target'] ?? item['scope'] ?? '').toLowerCase();
              if (target.includes('analysis.windows.net')) return true;
            } catch { /* ignore */ }
          }
        }
      }
      return false;
    });
    if (!msalReady) await page.waitForTimeout(2_000);
  }
  console.log(msalReady
    ? `✅ MSAL token found in storage (${Math.round((Date.now() - msalStart) / 1000)}s)`
    : '⚠️  MSAL token not found — saving state anyway (main:run will retry)');

  // Save auth state for reuse in test runs
  await page.context().storageState({ path: authFile });
  console.log(`✅ Cygnus auth state saved to: ${authFile}`);
});
