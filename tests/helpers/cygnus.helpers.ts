import { type Page, type Locator, expect } from '@playwright/test';
import { getPocConfig } from './poc-config.helpers';

// â”€â”€ Failure codes â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
export type CygnusFailureCode =
  | 'PAGE_NOT_LOADED'
  | 'TAB_NOT_VISIBLE'
  | 'SECTION_NOT_VISIBLE'
  | 'LABEL_MISSING'
  | 'VALUE_MISSING'
  | 'VALUE_MISMATCH';

export interface FieldResult {
  section: string;
  label: string;
  expected: string;
  actual: string;
  status: 'PASS' | 'FAIL';
  failureCode?: CygnusFailureCode;
}

// â”€â”€ URLs â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Workspace list page â€” used only for auth setup
const POC = getPocConfig();

export const CYGNUS_WORKSPACE_URL = POC.workspaceUrl;

// Direct URL to the CygnusIN_gantt_chart report (Employee tab)
export const CYGNUS_REPORT_URL = POC.reportUrl;

// Kept for backwards compatibility
export const CYGNUS_URL = CYGNUS_WORKSPACE_URL;
export const CYGNUS_REPORT_NAME = 'CygnusIN_gantt_chart';

// â”€â”€ Known section headers on the Employee tab â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
export const SECTIONS = {
  MY_ORGANISATION: 'My Organization',
  DIAMOND_CLUB: 'DIAMOND CLUB',
  QUALIFICATIONS: 'My Qualifications & Skills',
  ASSETS: 'My Asset',
  MY_BADGES: 'My Badges',
} as const;

function escapeRegex(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

async function navigateToTab(page: Page, tabName: string, expectedMarker?: string): Promise<void> {
  await page.waitForSelector('.textRun', { timeout: 20_000 });

  const tabButton = page
    .locator('.visual-actionButton')
    .filter({ hasText: new RegExp(`^${escapeRegex(tabName)}$`) })
    .first();

  await tabButton.waitFor({ state: 'visible', timeout: 20_000 });
  await tabButton.click({ timeout: 10_000 });

  if (expectedMarker) {
    await expect(
      page.locator('.textRun, div.title', { hasText: expectedMarker }).first()
    ).toBeVisible({ timeout: 15_000 });
    return;
  }

  await page.waitForTimeout(5_000);
  await page.waitForSelector('.textRun', { timeout: 20_000 });
}

// â”€â”€ Navigate & wait for page to load â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
/**
 * Navigates to Cygnus and waits until the Power BI report canvas has rendered
 * at least one `.textRun` element (meaning the report is fully painted).
 */
export async function waitForReportLoad(page: Page): Promise<void> {
  // Navigate directly to the report URL â€” no workspace list click needed
  console.log(`⏳ Navigating to report: ${CYGNUS_REPORT_URL}`);
  await page.goto(CYGNUS_REPORT_URL, { waitUntil: 'domcontentloaded', timeout: 60_000 });

  // Wait for .textRun elements (Power BI renders labels first)
  await expect
    .poll(
      async () => page.locator('.textRun').count(),
      { timeout: 120_000, message: 'Power BI report did not render any .textRun elements within 120s' }
    )
    .toBeGreaterThan(0);

  console.log(`âœ… Report loaded: ${page.url()}`);
}

// â”€â”€ Tab navigation â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
/**
 * Clicks the Employee tab and waits for My Organization section to appear.
 */
export async function navigateToEmployeeTab(page: Page): Promise<void> {
  await navigateToTab(page, 'Employee', SECTIONS.MY_ORGANISATION);
}

/**
 * Clicks the Manager tab and waits for content to stabilise.
 *
 * The tab buttons (Employee / Manager / DU Head / CXO) are Power BI
 * "actionButton" visuals rendered with class "visual-actionButton" â€”
 * completely different from the ".textRun" elements used for report content.
 * Previous attempts using ".textRun" were searching the wrong element type.
 */
export async function navigateToManagerTab(page: Page): Promise<void> {
  await navigateToTab(page, 'Manager');
}

/**
 * Clicks the DU Head tab and waits for the report to re-render.
 */
export async function navigateToDUHeadTab(page: Page): Promise<void> {
  await navigateToTab(page, 'DU Head');
}

/**
 * Clicks the CXO tab and waits for the report to re-render.
 * Falls back to "Codir" if no button labelled "CXO" is found within 5 seconds â€”
 * some report versions show this tab as "Codir".
 */
export async function navigateToCXOTab(page: Page): Promise<void> {
  try {
    const btn = page
      .locator('.visual-actionButton')
      .filter({ hasText: new RegExp(`^${escapeRegex('CXO')}$`) })
      .first();
    await btn.waitFor({ state: 'visible', timeout: 5_000 });
    await navigateToTab(page, 'CXO');
  } catch {
    console.log('â„¹ï¸ "CXO" tab not found â€” trying "Codir" instead');
    await navigateToTab(page, 'Codir');
  }
}

// â”€â”€ Logged-in user profile (left sidebar) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
/**
 * Extracts the logged-in user's name and designation from the left sidebar.
 *
 * Strategy: scan ALL visible text elements in the sidebar x-band, then
 * split them into two groups â€” those ABOVE the first nav item (Profile/My Profile)
 * and those below. The name is the topmost text above the nav; designation is second.
 *
 * Covers .textRun and div.title since Power BI can render the profile card either way.
 */
export async function getLoggedInUser(page: Page): Promise<{ name: string | null; designation: string | null }> {
  // Scan both .textRun and div.title across the whole page
  const allEls = page.locator('.textRun, div.title');
  const count = await allEls.count();

  // Known nav labels â€” used to find the y-position where the nav section starts
  const NAV_LABELS = new Set([
    'profile', 'my profile', 'employee history', 'utilization', 'rewards',
    'skills and certification', 'leave & attendance', 'learning', 'ireflect ilearn',
    'v2', 'reimbursement', 'fulfillment', 'talent acquisition', 'travel', 'communities',
  ]);

  const sidebarItems: Array<{ text: string; x: number; y: number; isNav: boolean }> = [];

  for (let i = 0; i < count; i++) {
    const el = allEls.nth(i);
    const text = (await el.textContent())?.trim() ?? '';
    if (!text || text.length > 60) continue;
    const box = await el.boundingBox();
    if (!box) continue;

    // Sidebar x band: left portion of the page (up to ~320px for 1920px viewport)
    if (box.x > 320) continue;
    if (box.y < 30) continue;   // skip very top (logo area)
    if (box.y > 700) continue;  // skip below visible sidebar

    const isNav = NAV_LABELS.has(text.toLowerCase());
    sidebarItems.push({ text, x: Math.round(box.x), y: Math.round(box.y), isNav });
  }

  sidebarItems.sort((a, b) => a.y - b.y);

  // Find where the nav starts (first nav label y)
  const firstNavItem = sidebarItems.find(i => i.isNav);
  const navStartY = firstNavItem?.y ?? 999;

  // Profile card items = non-nav items that appear ABOVE the nav section
  const profileItems = sidebarItems.filter(i => !i.isNav && i.y < navStartY);

  return {
    name:        profileItems[0]?.text ?? null,
    designation: profileItems[1]?.text ?? null,
  };
}

// â”€â”€ Section visibility check â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
/**
 * Asserts that a section header is visible on the current page.
 * Returns a FieldResult-style failure object if not found, or null if OK.
 */
export async function checkSectionVisible(
  page: Page,
  sectionName: string
): Promise<{ visible: boolean; failureCode?: CygnusFailureCode }> {
  // DIAMOND CLUB header is a div.title; all other section headers are .textRun
  const sectionEl = page.locator('.textRun, div.title', { hasText: sectionName }).first();
  const visible = await sectionEl.isVisible({ timeout: 8_000 }).catch(() => false);
  return visible ? { visible: true } : { visible: false, failureCode: 'SECTION_NOT_VISIBLE' };
}

// â”€â”€ Diamond Club â€” dynamic card extraction â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
/**
 * Dynamically extracts ALL visible fields from the Diamond Club card.
 *
 * Approach: collect every text element (tspan + HTML) inside the card region,
 * sort into visual rows, then apply content-pattern matching to assign semantic
 * field labels. Nothing is hardcoded by index or position â€” labels are inferred
 * from the shape of the content:
 *
 *   "Name (12345/67890)"  â†’ Name + Employee ID   (parenthetical digit pattern)
 *   "Text : (Grade)"      â†’ Designation           (" : (" pattern)
 *   Second line after name, before designation â†’ Job Stream
 *   "Current Project: X"  â†’ Current Project       (prefix match, SVG tspan)
 *   "Total Exp:" label row + next row value       â†’ Total Experience
 *   "SSI Exp:" label row  + next row value        â†’ SSI Experience
 */
export async function getDiamondClubFields(page: Page): Promise<Record<string, string>> {
  const results: Record<string, string> = {};

  // â”€â”€ Locate the Diamond Club card â”€â”€
  const headerEl = page.locator('.textRun, div.title', { hasText: /DIAMOND CLUB/i }).first();
  if (!await headerEl.isVisible({ timeout: 8_000 }).catch(() => false)) return results;
  const headerBox = await headerEl.boundingBox();
  if (!headerBox) return results;

  // Generous card boundary to capture all content regardless of zoom/DPI
  const C = {
    left:   headerBox.x - 30,
    top:    headerBox.y - 10,
    right:  headerBox.x + 540,
    bottom: headerBox.y + 400,
  };
  const inCard = (x: number, y: number) =>
    x >= C.left && x <= C.right && y >= C.top && y <= C.bottom;

  // â”€â”€ Collect all text items, deduplicated by text+position â”€â”€
  type Item = { text: string; x: number; y: number };
  const seen  = new Set<string>();
  const allItems: Item[] = [];

  const push = (text: string, bx: number, by: number) => {
    // Round to 5px grid for deduplication (SVG and HTML may render same text Â±3px apart)
    const key = `${text}|${Math.round(bx / 5) * 5}|${Math.round(by / 5) * 5}`;
    if (!seen.has(key)) {
      seen.add(key);
      allItems.push({ text, x: Math.round(bx), y: Math.round(by) });
    }
  };

  // SVG tspan elements
  const tspanEls = page.locator('tspan');
  const tc = await tspanEls.count();
  for (let i = 0; i < tc; i++) {
    const text = (await tspanEls.nth(i).textContent())?.trim() ?? '';
    if (!text) continue;
    const box = await tspanEls.nth(i).boundingBox();
    if (box && inCard(box.x, box.y)) push(text, box.x, box.y);
  }

  // HTML elements
  const htmlEls = page.locator('div.title, .textRun');
  const hc = await htmlEls.count();
  for (let i = 0; i < hc; i++) {
    const text = (await htmlEls.nth(i).textContent())?.trim() ?? '';
    if (!text || /^DIAMOND CLUB$/i.test(text)) continue;
    const box = await htmlEls.nth(i).boundingBox();
    if (box && inCard(box.x, box.y)) push(text, box.x, box.y);
  }

  // Sort top-to-bottom, left-to-right
  allItems.sort((a, b) => a.y - b.y || a.x - b.x);

  // â”€â”€ Group into visual rows (items within 10px y of each other = same row) â”€â”€
  const rows: Item[][] = [];
  for (const item of allItems) {
    const last = rows[rows.length - 1];
    if (last && Math.abs(item.y - last[0].y) <= 10) {
      last.push(item);
    } else {
      rows.push([item]);
    }
  }

  // â”€â”€ Parse each row with pattern matching â”€â”€
  let nameFound = false;

  for (let ri = 0; ri < rows.length; ri++) {
    const row     = rows[ri].sort((a, b) => a.x - b.x);
    const rowText = row.map(r => r.text).join(' ').trim();
    if (!rowText) continue;

    // â”€â”€ Current Project â”€â”€
    // SVG tspan: "Current Project: Int_CygnusIN_N1" (label+value in one string)
    // OR: two separate tspans on the same row
    const cpMatch = rowText.match(/current\s+project\s*:\s*(.+)/i);
    if (cpMatch && cpMatch[1].trim()) {
      results['Current Project'] = cpMatch[1].trim();
      continue;
    }
    if (/current\s+project/i.test(rowText) && row.length >= 2) {
      const labelItem = row.find(r => /current\s+project/i.test(r.text));
      const others    = row.filter(r => r !== labelItem);
      if (labelItem && others.length > 0) {
        results['Current Project'] = others.map(o => o.text).join(' ').trim();
        continue;
      }
    }

    // â”€â”€ Total Exp + SSI Exp labels â€” values on the NEXT row â”€â”€
    // Layout: [Total Exp:]  [SSI Exp :]   â† this row (no "Years" text)
    //         [0 Yrs 2 Mo]  [0 Yrs -2 Mo] â† next row
    if (/total\s+exp/i.test(rowText) && !/years/i.test(rowText)) {
      const nextRow   = ri + 1 < rows.length ? rows[ri + 1].sort((a, b) => a.x - b.x) : [];
      const totalLbl  = row.find(r => /total\s+exp/i.test(r.text));
      const ssiLbl    = row.find(r => /ssi\s+exp/i.test(r.text));

      const closestX  = (label: Item) => nextRow.reduce((best, cur) =>
        Math.abs(cur.x - label.x) < Math.abs(best.x - label.x) ? cur : best
      );

      if (totalLbl && nextRow.length > 0) {
        results['Total Experience'] = closestX(totalLbl).text;
      }
      if (ssiLbl && nextRow.length > 1) {
        const ssiVal = closestX(ssiLbl);
        if (ssiVal.text !== results['Total Experience']) {
          results['SSI Experience'] = ssiVal.text;
        }
      }
      continue;
    }

    // â”€â”€ Total Exp + SSI Exp inline (label and value on same row) â”€â”€
    if (/total\s+exp/i.test(rowText) && /years/i.test(rowText)) {
      const totalMatch = rowText.match(/total\s+exp\s*[:\s]+(.+?)(?:\s+ssi\s+exp|$)/i);
      if (totalMatch) results['Total Experience'] = totalMatch[1].trim();
      const ssiMatch = rowText.match(/ssi\s+exp\s*[:\s]+(.+)/i);
      if (ssiMatch) results['SSI Experience'] = ssiMatch[1].trim();
      continue;
    }

    // â”€â”€ Name + Employee ID â”€â”€
    // Pattern: "Firstname Lastname (empId/otherId)"
    const nameIdMatch = rowText.match(/^(.+?)\s+\((\d+\/\d+)\)/);
    if (nameIdMatch && !nameFound) {
      results['Name']        = nameIdMatch[1].trim();
      results['Employee ID'] = nameIdMatch[2].trim();
      nameFound = true;
      continue;
    }

    // â”€â”€ Designation â”€â”€
    // Pattern: "Engineer Trainee : (1'B)" â€” contains " : ("
    if (/\s*:\s*\(/.test(rowText) && !results['Designation']) {
      results['Designation'] = rowText;
      continue;
    }

    // â”€â”€ Job Stream â”€â”€
    // The line that appears AFTER Name but BEFORE Designation has no label on screen.
    // Detected positionally within the card â€” it's the only unlabelled line between them.
    if (nameFound && !results['Designation'] && !results['Job Stream']) {
      if (!Object.values(results).includes(rowText) && rowText.length > 2) {
        results['Job Stream'] = rowText;
      }
      continue;
    }

    // Skip anything already captured as a value
    if (Object.values(results).includes(rowText)) continue;
  }

  return results;
}


/**
 * Finds the value text associated with a given label on the Power BI report canvas.
 *
 * DOM discovery revealed three layouts:
 *  - My Org fields: label in .textRun, value in div.title (same row, to the right)
 *  - Department field: label in .textRun, value in .pivotTableCellWrap.main-cell (same row, to right)
 *  - Diamond Club: label in .textRun, value in div.title (directly BELOW label, same x column)
 *
 * Three-pass strategy:
 *  Pass 1 â€” scan div.title: same row (Â±20px), to the right â†’ catches My Org values
 *  Pass 2 â€” scan .pivotTableCellWrap.main-cell: same row, to right â†’ catches Department
 *  Pass 3 â€” scan div.title: 5â€“35px below, same x (Â±50px) â†’ catches Diamond Club values
 */
export async function getFieldValue(page: Page, labelText: string): Promise<string | null> {
  const labelEl = page.locator('.textRun', { hasText: labelText }).first();
  const labelVisible = await labelEl.isVisible({ timeout: 8_000 }).catch(() => false);

  // Pass 0: label is in an SVG tspan (e.g. Diamond Club "Current Project: <value>")
  // This runs when the label is NOT a .textRun element at all.
  if (!labelVisible) {
    const tspanEls = page.locator('tspan');
    const tCount = await tspanEls.count();
    for (let i = 0; i < tCount; i++) {
      const el = tspanEls.nth(i);
      const text = (await el.textContent())?.trim() ?? '';
      if (!text) continue;

      // Case A: entire "Label: Value" is in one tspan â€” e.g. "Current Project: 3SIP BU_0025_PUNEOF"
      if (text.toLowerCase().startsWith(labelText.toLowerCase())) {
        const colonIdx = text.indexOf(':');
        if (colonIdx !== -1 && colonIdx < text.length - 1) {
          const val = text.slice(colonIdx + 1).trim();
          if (val) return val;
        }
      }

      // Case B: label tspan sits next to a separate value tspan on the same row
      if (text.replace(':', '').trim().toLowerCase() === labelText.toLowerCase()) {
        const labelBox = await el.boundingBox();
        if (!labelBox) continue;
        let closest: string | null = null;
        let minDist = Infinity;
        for (let j = 0; j < tCount; j++) {
          if (j === i) continue;
          const sibling = tspanEls.nth(j);
          const sibText = (await sibling.textContent())?.trim() ?? '';
          if (!sibText) continue;
          const sibBox = await sibling.boundingBox();
          if (!sibBox) continue;
          if (Math.abs(sibBox.y - labelBox.y) > 15) continue;
          if (sibBox.x <= labelBox.x) continue;
          const dist = sibBox.x - (labelBox.x + labelBox.width);
          if (dist < minDist) { minDist = dist; closest = sibText; }
        }
        if (closest) return closest;
      }
    }
    return null;
  }

  const labelBox = await labelEl.boundingBox();
  if (!labelBox) return null;

  // Helper: find closest element to the RIGHT on the same row (Â±15px)
  const findToRight = async (selector: string): Promise<string | null> => {
    const els = page.locator(selector);
    const count = await els.count();
    let closest: string | null = null;
    let minDist = Infinity;
    for (let i = 0; i < count; i++) {
      const el = els.nth(i);
      const text = (await el.textContent())?.trim() ?? '';
      if (!text || text === labelText) continue;
      const box = await el.boundingBox();
      if (!box) continue;
      if (Math.abs(box.y - labelBox.y) > 15) continue;
      if (box.x <= labelBox.x + labelBox.width) continue;
      const dist = box.x - (labelBox.x + labelBox.width);
      if (dist < minDist) { minDist = dist; closest = text; }
    }
    return closest;
  };

  // Pass 1: div.title to the right on same row (My Org field values)
  const pass1 = await findToRight('div.title');
  if (pass1) return pass1;

  // Pass 2: pivotTable main-cell to the right on same row (Department field)
  const pass2 = await findToRight('.pivotTableCellWrap.cell-interactive.main-cell');
  if (pass2) return pass2;

  // Pass 3: div.title directly below label (Diamond Club: value sits under label)
  const titleEls = page.locator('div.title');
  const titleCount = await titleEls.count();
  let closestBelow: string | null = null;
  let minBelowDist = Infinity;
  for (let i = 0; i < titleCount; i++) {
    const el = titleEls.nth(i);
    const text = (await el.textContent())?.trim() ?? '';
    if (!text || text === labelText) continue;
    const box = await el.boundingBox();
    if (!box) continue;
    const verticalDiff = box.y - labelBox.y;
    const horizontalDiff = Math.abs(box.x - labelBox.x);
    if (verticalDiff >= 5 && verticalDiff <= 60 && horizontalDiff <= 200) {
      if (horizontalDiff < minBelowDist) { minBelowDist = horizontalDiff; closestBelow = text; }
    }
  }
  if (closestBelow) return closestBelow;

  // Pass 4: SVG tspan elements (Diamond Club card uses SVG for some values, e.g. Current Project)
  const tspanEls = page.locator('tspan');
  const tspanCount = await tspanEls.count();
  let closestTspan: string | null = null;
  let minTspanDist = Infinity;
  for (let i = 0; i < tspanCount; i++) {
    const el = tspanEls.nth(i);
    const text = (await el.textContent())?.trim() ?? '';
    if (!text || text === labelText) continue;
    const box = await el.boundingBox();
    if (!box) continue;
    // Same row (Â±25px) and to the right, OR directly below (5â€“50px) and same column (Â±100px)
    const vertDiff = box.y - labelBox.y;
    const horizDiff = box.x - (labelBox.x + labelBox.width);
    const sameRow = Math.abs(vertDiff) <= 25 && horizDiff > 0;
    const below = vertDiff >= 5 && vertDiff <= 50 && Math.abs(box.x - labelBox.x) <= 100;
    if (sameRow) {
      if (horizDiff < minTspanDist) { minTspanDist = horizDiff; closestTspan = text; }
    } else if (below && closestTspan === null) {
      closestTspan = text;
    }
  }
  return closestTspan;
}

// â”€â”€ Get all list items under a section â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Pivot table column headers to exclude from list items
const PIVOT_COLUMN_HEADERS = new Set([
  'Qualification', 'Skill', 'iReflect Skill',
  'ASSET_MODEL', 'Department name',
]);

/**
 * Collects list items under a section header.
 *
 * DOM discovery revealed:
 *  - Qualifications/Skills are in .pivotTableCellWrap.cell-interactive.main-cell
 *  - Assets are in .pivotTableCellWrap.cell-interactive.prefix-cell
 *  - Section headers (My Qualifications & Skills, My Asset) can share the same y row,
 *    so x-range filtering (Â±350px of header x) is used to separate adjacent sections.
 */
export async function getSectionItems(page: Page, sectionHeader: string): Promise<string[]> {
  // Get section header position from .textRun
  const textRuns = page.locator('.textRun');
  const trCount = await textRuns.count();
  const textRunItems: Array<{ text: string; x: number; y: number }> = [];
  for (let i = 0; i < trCount; i++) {
    const el = textRuns.nth(i);
    const text = (await el.textContent())?.trim() ?? '';
    if (!text) continue;
    const box = await el.boundingBox();
    if (!box) continue;
    textRunItems.push({ text, x: Math.round(box.x), y: Math.round(box.y) });
  }

  const headerItem = textRunItems.find((item) => item.text === sectionHeader);
  if (!headerItem) return [];

  // Find the y of the next .textRun section header below this one
  const knownSectionHeaders = new Set(Object.values(SECTIONS));
  const nextSectionY = textRunItems
    .filter((item) => knownSectionHeaders.has(item.text as any) && item.y > headerItem.y)
    .sort((a, b) => a.y - b.y)[0]?.y ?? Infinity;

  // Find x boundary: if another section header is on the same row (Â±15px),
  // split at the midpoint â€” prevents cross-contamination between side-by-side sections
  const sameSectionToRight = textRunItems
    .filter((item) =>
      knownSectionHeaders.has(item.text as any) &&
      Math.abs(item.y - headerItem.y) <= 15 &&
      item.x > headerItem.x
    )
    .sort((a, b) => a.x - b.x)[0];

  const xRightBoundary = sameSectionToRight
    ? (headerItem.x + sameSectionToRight.x) / 2
    : headerItem.x + 400;

  const sameSectionToLeft = textRunItems
    .filter((item) =>
      knownSectionHeaders.has(item.text as any) &&
      Math.abs(item.y - headerItem.y) <= 15 &&
      item.x < headerItem.x
    )
    .sort((a, b) => b.x - a.x)[0];

  const xLeftBoundary = sameSectionToLeft
    ? (sameSectionToLeft.x + headerItem.x) / 2
    : headerItem.x - 50;

  // Collect pivot table items
  const pivotEls = page.locator(
    '.pivotTableCellWrap.cell-interactive.main-cell, .pivotTableCellWrap.cell-interactive.prefix-cell'
  );
  const pvCount = await pivotEls.count();
  const result: string[] = [];

  for (let i = 0; i < pvCount; i++) {
    const el = pivotEls.nth(i);
    const text = (await el.textContent())?.trim() ?? '';
    if (!text) continue;
    if (PIVOT_COLUMN_HEADERS.has(text)) continue;
    const box = await el.boundingBox();
    if (!box) continue;
    // Must be below section header and above next section
    if (box.y <= headerItem.y) continue;
    if (box.y >= nextSectionY) continue;
    // Must be within the dynamic x-range for this section
    if (box.x < xLeftBoundary || box.x > xRightBoundary) continue;
    result.push(text);
  }

  return result;
}

// â”€â”€ Validate a single field â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
/**
 * Validates one label-value pair against an expected value.
 * Returns a FieldResult with status PASS or FAIL and the actual value found.
 */
export async function validateField(
  page: Page,
  section: string,
  label: string,
  expected: string
): Promise<FieldResult> {
  const labelEl = page.locator('.textRun', { hasText: label }).first();
  const labelVisible = await labelEl.isVisible({ timeout: 5_000 }).catch(() => false);

  if (!labelVisible) {
    return {
      section,
      label,
      expected,
      actual: '',
      status: 'FAIL',
      failureCode: 'LABEL_MISSING',
    };
  }

  const actual = await getFieldValue(page, label);

  if (actual === null || actual === '') {
    return {
      section,
      label,
      expected,
      actual: actual ?? '',
      status: 'FAIL',
      failureCode: 'VALUE_MISSING',
    };
  }

  // Case-insensitive trim comparison
  const match = actual.trim().toLowerCase() === expected.trim().toLowerCase();

  return {
    section,
    label,
    expected,
    actual,
    status: match ? 'PASS' : 'FAIL',
    failureCode: match ? undefined : 'VALUE_MISMATCH',
  };
}

// â”€â”€ Validate a list-based section â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
/**
 * Validates list items under a section (e.g. Skills, Assets).
 * Each expectedItem must appear somewhere in the actual list (substring match).
 */
export async function validateSectionList(
  page: Page,
  section: string,
  expectedItems: string[]
): Promise<FieldResult[]> {
  const actualItems = await getSectionItems(page, section);
  const results: FieldResult[] = [];

  for (const expected of expectedItems) {
    const found = actualItems.some((actual) =>
      actual.toLowerCase().includes(expected.toLowerCase())
    );

    results.push({
      section,
      label: expected,
      expected,
      actual: found ? expected : '',
      status: found ? 'PASS' : 'FAIL',
      failureCode: found ? undefined : 'VALUE_MISSING',
    });
  }

  return results;
}

// â”€â”€ Visual container extraction â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

export interface VisualData {
  visualIndex: number;
  bounds: { x: number; y: number; w: number; h: number };
  type: 'kpi' | 'metric' | 'chart' | 'table' | 'text' | 'unknown';
  title: string | null;
  kpiLabel: string | null;
  kpiValue: string | null;
  metricLabel: string | null;
  metricValue: string | null;
  chartLabels: string[];
  chartValues: string[];
  tableHeaders: string[];
  tableRows: string[][];
  textItems: string[];
  rawAriaLabel: string | null;
  /** true when data was obtained via the Power BI "Show as table" toggle button */
  tableToggleUsed: boolean;
}

/** Regex to parse "KPI Label value." from img aria-label attributes */
const KPI_ARIA_PATTERN = /^(.+?)\s+([\d,+%().NA\-]+)\.?\s*$/;

// Selectors for the Power BI "Show as table" toggle button (revealed on hover)
const TABLE_TOGGLE_SELECTORS = [
  'button[aria-label="Show as a table"]',
  'button[aria-label*="Show as a table"]',
  'button[title="Show as a table"]',
  'button[aria-label="Switch to table"]',
].join(', ');

// Selectors for toggling back from table to chart view
const CHART_TOGGLE_SELECTORS = [
  'button[aria-label="Show visual"]',
  'button[aria-label="Switch to chart"]',
  'button[aria-label*="Switch to chart"]',
  'button[aria-label*="Show visual"]',
  'button[title="Show visual"]',
].join(', ');

/**
 * Attempts to hover over a visual container, click its "Show as table" toggle button,
 * read the resulting pivot table data, and toggle back to chart view.
 * Returns table headers + rows, or null if no toggle was found.
 */
async function tryTableToggleForContainer(
  page: Page,
  cx: number, cy: number, cw: number, ch: number,
): Promise<{ headers: string[]; rows: string[][] } | null> {
  // Hover at the center of the visual to reveal the header controls
  await page.mouse.move(cx + cw / 2, cy + ch / 2);
  await page.waitForTimeout(700);

  // Look for the toggle button â€” Power BI may render it in a portal outside the container
  const toggleBtn = page.locator(TABLE_TOGGLE_SELECTORS).first();
  const found = await toggleBtn.isVisible({ timeout: 1_500 }).catch(() => false);
  if (!found) return null;

  // Click to switch to table view
  await toggleBtn.click({ timeout: 5_000 });
  await page.waitForTimeout(2_000);

  // Collect pivot table cells within the container's bounding box
  const pivotEls = page.locator('.pivotTableCellWrap');
  const pvCount  = await pivotEls.count();
  type Cell = { text: string; x: number; y: number };
  const cells: Cell[] = [];
  for (let i = 0; i < pvCount; i++) {
    const text = (await pivotEls.nth(i).textContent())?.trim() ?? '';
    if (!text) continue;
    const box = await pivotEls.nth(i).boundingBox();
    if (!box) continue;
    if (box.x >= cx && box.x <= cx + cw && box.y >= cy && box.y <= cy + ch) {
      cells.push({ text, x: Math.round(box.x), y: Math.round(box.y) });
    }
  }

  if (cells.length === 0) return null;

  // Group cells into rows by Y position (Â±6px tolerance)
  cells.sort((a, b) => a.y - b.y || a.x - b.x);
  const rowGroups: Cell[][] = [];
  for (const cell of cells) {
    const last = rowGroups[rowGroups.length - 1];
    if (last && Math.abs(cell.y - last[0].y) <= 6) {
      last.push(cell);
    } else {
      rowGroups.push([cell]);
    }
  }

  const headers = (rowGroups[0] ?? []).sort((a, b) => a.x - b.x).map(c => c.text);
  const rows    = rowGroups.slice(1).map(row => row.sort((a, b) => a.x - b.x).map(c => c.text));

  // Try to toggle back to chart view
  await page.mouse.move(cx + cw / 2, cy + ch / 2);
  await page.waitForTimeout(400);
  const backBtn   = page.locator(CHART_TOGGLE_SELECTORS).first();
  const backFound = await backBtn.isVisible({ timeout: 1_000 }).catch(() => false);
  if (backFound) await backBtn.click({ timeout: 3_000 }).catch(() => {});
  await page.waitForTimeout(500);

  return { headers, rows };
}

/**
 * Scans all img[aria-label] elements on the page and parses their label/value.
 * Useful as a global fallback for KPI cards that Power BI renders as images.
 */
export async function extractImgKpiCards(
  page: Page
): Promise<Array<{ label: string; value: string | null; raw: string }>> {
  const imgs  = page.locator('img[aria-label]');
  const count = await imgs.count();
  const out: Array<{ label: string; value: string | null; raw: string }> = [];
  for (let i = 0; i < count; i++) {
    const raw = (await imgs.nth(i).getAttribute('aria-label'))?.trim() ?? '';
    if (!raw) continue;
    const m = raw.match(KPI_ARIA_PATTERN);
    if (m) {
      out.push({ label: m[1].trim(), value: m[2].trim(), raw });
    } else {
      out.push({ label: raw, value: null, raw });
    }
  }
  return out;
}

/**
 * Finds all Power BI visual containers and extracts structured data from each.
 *
 * For each visual container this function:
 *  1. Detects the container type (KPI card, inline metric, chart, table, text)
 *  2. For chart-type containers, hovers to reveal the "Show as table" toggle button,
 *     clicks it to get fully structured tabular data, then toggles back to chart view
 *  3. Falls back to raw tspan / textRun collection for containers without the toggle
 *
 * Container detection filters out the full-page wrapper (> 1700 Ã— 900 px) so that
 * only individual visual containers are processed.
 */
export async function extractAllVisuals(page: Page): Promise<VisualData[]> {
  const result: VisualData[] = [];

  // â”€â”€ Step 1: Find individual container bounding boxes via JS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  // This avoids the "full-page wrapper" problem: the JS filter skips containers
  // whose dimensions cover the whole viewport.
  type ContainerBox = { x: number; y: number; w: number; h: number };
  const containerBoxes: ContainerBox[] = await page.evaluate(() => {
    const seen = new Set<string>();
    const out: Array<{ x: number; y: number; w: number; h: number }> = [];
    const selectors = [
      '.visual-container-component',
      '[class*="visualContainerHost"]',
      '[class*="visualContainer"]',
    ];
    for (const sel of selectors) {
      document.querySelectorAll(sel).forEach(el => {
        const r = el.getBoundingClientRect();
        // Skip full-page wrappers and tiny elements
        if (r.width > 1700 || r.height > 900) return;
        if (r.width < 60  || r.height < 40)  return;
        const key = `${Math.round(r.left)},${Math.round(r.top)}`;
        if (seen.has(key)) return;
        // Only include if the container holds visible content
        const hasContent = !!(
          el.querySelector('tspan, .textRun, img[aria-label], .pivotTableCellWrap, div.title')
        );
        if (!hasContent) return;
        seen.add(key);
        out.push({
          x: Math.round(r.left), y: Math.round(r.top),
          w: Math.round(r.width), h: Math.round(r.height),
        });
      });
    }
    return out;
  });

  // â”€â”€ Step 2: Pre-collect page-level elements once â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  type PosEl = { text: string; x: number; y: number };

  const allImgKpi: Array<{ raw: string; label: string; value: string | null; x: number; y: number }> = [];
  const imgEls  = page.locator('img[aria-label]');
  const imgCount = await imgEls.count();
  for (let i = 0; i < imgCount; i++) {
    const raw = (await imgEls.nth(i).getAttribute('aria-label'))?.trim() ?? '';
    if (!raw) continue;
    const box = await imgEls.nth(i).boundingBox();
    if (!box) continue;
    const m = raw.match(KPI_ARIA_PATTERN);
    allImgKpi.push({ raw, label: m ? m[1].trim() : raw, value: m ? m[2].trim() : null, x: box.x, y: box.y });
  }

  const allDivTitle: PosEl[] = [];
  const divTitleEls = page.locator('div.title');
  const dtCount = await divTitleEls.count();
  for (let i = 0; i < dtCount; i++) {
    const text = (await divTitleEls.nth(i).textContent())?.trim() ?? '';
    if (!text) continue;
    const box = await divTitleEls.nth(i).boundingBox();
    if (box) allDivTitle.push({ text, x: box.x, y: box.y });
  }

  const allTextRun: PosEl[] = [];
  const textRunEls = page.locator('.textRun');
  const trCount = await textRunEls.count();
  for (let i = 0; i < trCount; i++) {
    const text = (await textRunEls.nth(i).textContent())?.trim() ?? '';
    if (!text) continue;
    const box = await textRunEls.nth(i).boundingBox();
    if (box) allTextRun.push({ text, x: box.x, y: box.y });
  }

  const allTspan: PosEl[] = [];
  const tspanEls = page.locator('tspan');
  const tsCount  = await tspanEls.count();
  for (let i = 0; i < tsCount; i++) {
    const text = (await tspanEls.nth(i).textContent())?.trim() ?? '';
    if (!text || text.length > 120) continue;
    const box = await tspanEls.nth(i).boundingBox();
    if (box) allTspan.push({ text, x: box.x, y: box.y });
  }

  const inBox = (ex: number, ey: number, b: ContainerBox) =>
    ex >= b.x && ex <= b.x + b.w && ey >= b.y && ey <= b.y + b.h;

  const NUMERIC_RE = /^[\d,.+%\-]+$/;

  // â”€â”€ Step 3: Process each container â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  for (const box of containerBoxes) {
    const visual: VisualData = {
      visualIndex:     result.length,
      bounds:          { x: box.x, y: box.y, w: box.w, h: box.h },
      type:            'unknown',
      title:           null,
      kpiLabel:        null, kpiValue:    null, rawAriaLabel: null,
      metricLabel:     null, metricValue: null,
      chartLabels:     [], chartValues: [],
      tableHeaders:    [], tableRows:   [],
      textItems:       [],
      tableToggleUsed: false,
    };

    // KPI card (img[aria-label])
    for (const img of allImgKpi) {
      if (inBox(img.x, img.y, box)) {
        visual.kpiLabel    = img.label;
        visual.kpiValue    = img.value;
        visual.rawAriaLabel = img.raw;
        visual.type  = 'kpi';
        visual.title = img.label;
        break;
      }
    }

    // Inline metric (div.title with " : ")
    if (visual.type === 'unknown') {
      for (const dt of allDivTitle) {
        if (!inBox(dt.x, dt.y, box)) continue;
        const ci = dt.text.indexOf(' : ');
        if (ci === -1) continue;
        const lbl = dt.text.slice(0, ci).trim();
        const val = dt.text.slice(ci + 3).trim();
        if (lbl && val && !lbl.includes(' = ')) {
          visual.metricLabel  = lbl;
          visual.metricValue  = val;
          visual.type  = 'metric';
          visual.title = lbl;
          break;
        }
      }
    }

    // .textRun title items (first one = visual title candidate)
    for (const tr of allTextRun) {
      if (inBox(tr.x, tr.y, box)) visual.textItems.push(tr.text);
    }
    if (!visual.title && visual.textItems.length > 0) visual.title = visual.textItems[0];

    // For chart/unknown containers: try the table toggle first â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    if (visual.type === 'unknown') {
      const tableData = await tryTableToggleForContainer(page, box.x, box.y, box.w, box.h);
      if (tableData && (tableData.headers.length > 0 || tableData.rows.length > 0)) {
        visual.tableHeaders    = tableData.headers;
        visual.tableRows       = tableData.rows;
        visual.tableToggleUsed = true;
        visual.type            = 'table';
      } else {
        // Fallback: collect tspan elements within the container's bounds
        for (const ts of allTspan) {
          if (!inBox(ts.x, ts.y, box)) continue;
          if (NUMERIC_RE.test(ts.text)) visual.chartValues.push(ts.text);
          else                          visual.chartLabels.push(ts.text);
        }
        if (visual.chartValues.length >= 3 || visual.chartLabels.length >= 2) {
          visual.type = 'chart';
        } else if (visual.textItems.length > 0) {
          visual.type = 'text';
        }
      }
    }

    result.push(visual);
  }

  // â”€â”€ Step 4: If no individual containers were found, fall back to global scans
  if (result.length === 0) {
    const kpiCards = await extractImgKpiCards(page);
    kpiCards.forEach((kpi, i) => {
      if (kpi.value === null) return;
      result.push({
        visualIndex: i, bounds: { x: 0, y: 0, w: 0, h: 0 },
        type: 'kpi', title: kpi.label,
        kpiLabel: kpi.label, kpiValue: kpi.value, rawAriaLabel: kpi.raw,
        metricLabel: null, metricValue: null,
        chartLabels: [], chartValues: [], tableHeaders: [], tableRows: [],
        textItems: [], tableToggleUsed: false,
      });
    });
    allDivTitle.forEach((dt, i) => {
      const ci = dt.text.indexOf(' : ');
      if (ci === -1) return;
      const lbl = dt.text.slice(0, ci).trim();
      const val = dt.text.slice(ci + 3).trim();
      if (!lbl.includes(' = ')) {
        result.push({
          visualIndex: kpiCards.length + i, bounds: { x: dt.x, y: dt.y, w: 0, h: 0 },
          type: 'metric', title: lbl,
          kpiLabel: null, kpiValue: null, rawAriaLabel: null,
          metricLabel: lbl, metricValue: val,
          chartLabels: [], chartValues: [], tableHeaders: [], tableRows: [],
          textItems: [], tableToggleUsed: false,
        });
      }
    });
  }

  return result;
}

// â”€â”€ Manager tab â€” semantic metric extraction â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
/**
 * Extracts labelled KPI metrics from the Manager tab.
 *
 * Pattern 1 â€” Inline (div.title / tspan containing "Label : Value"):
 *   "My Team : 5877", "Employees Serving Notice Period : 174", etc.
 *
 * Pattern 2 â€” img[aria-label] KPI cards:
 *   img aria-label="Headcount Today 5878." â†’ { "Headcount Today": "5878" }
 *
 * Pattern 3 â€” Container-scoped .textRun label + tspan value:
 *   Uses extractAllVisuals() to scope pairing within the same visual container,
 *   fixing the earlier 250px-tolerance bug that grabbed bar-chart values incorrectly.
 *
 * Returns a flat Record<label, value>.
 */
export async function getManagerMetrics(page: Page): Promise<Record<string, string>> {
  const results: Record<string, string> = {};

  // â”€â”€ Pattern 1a: img[aria-label] KPI cards â”€â”€
  const kpiCards = await extractImgKpiCards(page);
  for (const { label, value } of kpiCards) {
    if (value !== null && label) results[label] = value;
  }

  // â”€â”€ Pattern 1b: div.title elements with " : " separator â”€â”€
  const divTitleEls = page.locator('div.title');
  const dtCount = await divTitleEls.count();
  for (let i = 0; i < dtCount; i++) {
    const text = (await divTitleEls.nth(i).textContent())?.trim() ?? '';
    if (!text) continue;
    const colonIdx = text.indexOf(' : ');
    if (colonIdx === -1) continue;
    const label = text.slice(0, colonIdx).trim();
    const value = text.slice(colonIdx + 3).trim();
    if (label && value && !label.includes(' = ') && !results[label]) results[label] = value;
  }

  // â”€â”€ Pattern 1c: tspan elements with " : " separator â”€â”€
  const tspanEls = page.locator('tspan');
  const tsCount = await tspanEls.count();
  for (let i = 0; i < tsCount; i++) {
    const text = (await tspanEls.nth(i).textContent())?.trim() ?? '';
    if (!text) continue;
    const colonIdx = text.indexOf(' : ');
    if (colonIdx === -1) continue;
    const label = text.slice(0, colonIdx).trim();
    const value = text.slice(colonIdx + 3).trim();
    if (label && value && !label.includes(' = ') && !results[label]) results[label] = value;
  }

  // â”€â”€ Pattern 2: Container-scoped .textRun label + tspan value â”€â”€
  // Use extractAllVisuals to scope matching within the same visual bounding box,
  // preventing cross-visual proximity errors.
  const visuals = await extractAllVisuals(page);
  for (const v of visuals) {
    if (v.textItems.length === 0 || v.chartValues.length === 0) continue;
    for (const label of v.textItems) {
      if (!label || results[label]) continue;
      // Pick the single closest numeric value within this container
      if (v.chartValues.length === 1) {
        results[label] = v.chartValues[0];
      } else if (v.chartValues.length > 1) {
        // Only record when there's an unambiguous value (e.g. single KPI value visual)
        // Skip chart containers with many values â€” those are bar/donut charts
        if (v.chartValues.length <= 3) {
          results[label] = v.chartValues[0];
        }
      }
    }
  }

  return results;
}
