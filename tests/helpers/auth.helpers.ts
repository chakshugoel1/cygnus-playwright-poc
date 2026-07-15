import { expect, type Page } from '@playwright/test';
import * as fs from 'fs';

/**
 * Waits for the Microsoft Entra submit button (#idSIButton9) to be clickable, then clicks it.
 * Safe to call even if the popup is already advancing (e.g. FIDO redirect).
 */
export async function clickSubmit(popup: Page): Promise<void> {
  const submitBtn = popup.locator('#idSIButton9').first();
  try {
    await expect
      .poll(
        async () => {
          if (popup.isClosed()) return 'done';
          const visible = await submitBtn.isVisible().catch(() => false);
          const enabled = await submitBtn.isEnabled().catch(() => false);
          if (visible && enabled) return 'ready';
          if (/fido/i.test(popup.url())) return 'done';
          return 'waiting';
        },
        { timeout: 15_000 }
      )
      .not.toBe('waiting');

    if (
      !popup.isClosed() &&
      (await submitBtn.isVisible().catch(() => false)) &&
      (await submitBtn.isEnabled().catch(() => false))
    ) {
      await submitBtn.click();
    } else {
      console.log('ℹ️ Submit click skipped — auth flow already advanced');
    }
  } catch {
    console.log('ℹ️ Submit step skipped — auth likely continued automatically');
  }
}

/**
 * Saves Playwright browser auth state (cookies + localStorage) and MSAL sessionStorage
 * to the given file paths.
 */
export async function saveAuthState(
  page: Page,
  authFile: string,
  sessionFile: string
): Promise<void> {
  await page.context().storageState({ path: authFile });

  const sessionStorageData = await page.evaluate(() => {
    const data: Record<string, string> = {};
    for (let i = 0; i < window.sessionStorage.length; i++) {
      const key = window.sessionStorage.key(i);
      if (key) {
        const value = window.sessionStorage.getItem(key);
        if (value !== null) data[key] = value;
      }
    }
    return data;
  });

  fs.mkdirSync(require('path').dirname(sessionFile), { recursive: true });
  fs.writeFileSync(sessionFile, JSON.stringify(sessionStorageData, null, 2), 'utf-8');

  console.log(`✅ Saved authenticated state to: ${authFile}`);
  console.log(`✅ Saved sessionStorage to: ${sessionFile}`);
}
