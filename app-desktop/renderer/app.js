const FIELD_PLACEHOLDERS = {
  tenantId: 'Enter tenant ID',
  groupId: 'Enter group/workspace ID',
  reportId: 'Enter report ID',
  datasetId: 'Enter dataset ID',
  rlsRole: 'Enter role name, if needed',
};

const SOURCE_KEYS = [
  ['tenantId', 'Tenant ID'],
  ['groupId', 'Group ID / Workspace ID'],
  ['reportId', 'Report ID'],
  ['datasetId', 'Dataset ID'],
  ['rlsRole', 'RLS Role (optional)'],
];

const TARGET_KEYS = [
  ['tenantId', 'Tenant ID'],
  ['groupId', 'Group ID / Workspace ID'],
  ['reportId', 'Report ID'],
  ['datasetId', 'Dataset ID'],
  ['rlsRole', 'RLS Role (optional)'],
];

function createRows(containerId, prefix, items) {
  const container = document.getElementById(containerId);
  items.forEach(([key, label]) => {
    const row = document.createElement('div');
    row.className = 'row';

    const fieldWrap = document.createElement('div');
    fieldWrap.className = 'field-wrap';

    const lbl = document.createElement('label');
    lbl.textContent = label;
    lbl.setAttribute('for', `${prefix}_${key}`);

    const input = document.createElement('input');
    input.id = `${prefix}_${key}`;
    input.placeholder = FIELD_PLACEHOLDERS[key] || '';

    fieldWrap.appendChild(lbl);

    row.appendChild(fieldWrap);
    row.appendChild(input);
    container.appendChild(row);
  });
}

createRows('source-fields', 'source', SOURCE_KEYS);
createRows('target-fields', 'target', TARGET_KEYS);

const logEl = document.getElementById('log');
const statusEl = document.getElementById('status');
const buttons = [
  document.getElementById('btnSetup'),
  document.getElementById('btnParity'),
  document.getElementById('btnBoth'),
  document.getElementById('btnDiscover'),
];

function appendLog(line) {
  logEl.textContent += `${line}\n`;
  logEl.scrollTop = logEl.scrollHeight;
}

// Discover Filters can't complete without a saved login session: it runs
// with --no-deps (skips the interactive 'setup' auth step, deliberately, so
// discovery doesn't force a fresh sign-in every time), and cross-report
// discovery specifically runs its browser windows off-screen for speed - so
// if a Microsoft sign-in page ever appeared, there would be no way to see or
// interact with it. Gate the button on a known-good session instead of
// letting that run start and fail confusingly.
let _isBusy = false;
let _hasSession = false;

function updateButtonStates() {
  buttons.forEach(b => { b.disabled = _isBusy; });
  const discoverBtn = document.getElementById('btnDiscover');
  discoverBtn.disabled = _isBusy || !_hasSession;
  const note = document.getElementById('discoverAuthNote');
  if (note) note.style.display = (!_isBusy && !_hasSession) ? 'block' : 'none';
}

function setBusy(isBusy) {
  _isBusy = isBusy;
  updateButtonStates();
}

async function refreshAuthGate() {
  try {
    const status = await window.cygnusDesktop.getAuthStatus();
    _hasSession = !!(status && status.hasSession);
  } catch {
    _hasSession = false;
  }
  updateButtonStates();
}
refreshAuthGate();

function readIdentity(prefix) {
  const obj = {};
  ['tenantId', 'groupId', 'reportId', 'datasetId', 'rlsRole'].forEach((k) => {
    obj[k] = document.getElementById(`${prefix}_${k}`).value.trim();
  });
  if (!obj.rlsRole) delete obj.rlsRole;
  return obj;
}

function getPagesInScope() {
  const pagesCsv = document.getElementById('pagesCsv').value.trim();
  return pagesCsv ? pagesCsv.split(',').map(s => s.trim()).filter(Boolean) : [];
}

function buildConfig() {
  const source = readIdentity('source');
  const target = readIdentity('target');
  const pairName = document.getElementById('pairName').value.trim() || 'Cygnus';
  const pages = getSelectedRunPages();
  const slicerScenario = buildScenarioFromPicks(pages);
  const lenientTextCompare = document.getElementById('lenientTextCompare')?.checked === true;

  return {
    pairName,
    source,
    target,
    pages: pages.length ? pages : undefined,
    slicerScenario,
    lenientTextCompare,
  };
}

function validateForParity(cfg) {
  const required = ['tenantId', 'groupId', 'reportId', 'datasetId'];
  const missing = [];
  required.forEach((k) => {
    if (!cfg.source[k]) missing.push(`source.${k}`);
    if (!cfg.target[k]) missing.push(`target.${k}`);
  });
  return missing;
}

const IDENTITY_REQUIRED_KEYS = ['tenantId', 'groupId', 'reportId', 'datasetId'];

/**
 * Discovery only ever reads ONE report identity, but it must now be supplied
 * explicitly from the UI. We accept either a fully-populated Source or a
 * fully-populated Target identity; we do NOT fall back to the hardcoded pair.
 */
function resolveDiscoverIdentity() {
  const source = readIdentity('source');
  const target = readIdentity('target');
  const sourceTouched = IDENTITY_REQUIRED_KEYS.some((k) => source[k]);
  const targetTouched = IDENTITY_REQUIRED_KEYS.some((k) => target[k]);

  if (!sourceTouched && !targetTouched) {
    return {
      identity: null,
      side: null,
      missing: IDENTITY_REQUIRED_KEYS.map((k) => `source.${k}`),
      note: 'fill in all Source report fields (or fill in all Target report fields instead)',
    };
  }

  const side = sourceTouched ? 'source' : 'target';
  const candidate = sourceTouched ? source : target;
  const missing = IDENTITY_REQUIRED_KEYS.filter((k) => !candidate[k]).map((k) => `${side}.${k}`);
  return {
    identity: missing.length === 0 ? candidate : null,
    side,
    missing,
    note: side === 'source'
      ? 'fill in all Source report fields before discovering'
      : 'fill in all Target report fields before discovering',
  };
}

/** True when Source AND Target are both fully filled in - the trigger for
 *  cross-report discovery instead of the single-side flow. */
function bothIdentitiesFilled() {
  const source = readIdentity('source');
  const target = readIdentity('target');
  return IDENTITY_REQUIRED_KEYS.every((k) => source[k]) && IDENTITY_REQUIRED_KEYS.every((k) => target[k]);
}

// Mirrors tests/helpers/report-export.helpers.ts's pageNameKey exactly (same
// reasoning as main.js's copy: this is plain browser JS, no TypeScript
// compilation step, so it can't import the .ts version directly).
function pageNameKeyJs(name) {
  return String(name || '').trim().toLowerCase().replace(/[\s_-]+/g, '');
}

function dedupePageNames(names) {
  const seen = new Set();
  const out = [];
  for (const n of names) {
    const key = pageNameKeyJs(n);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push(n);
  }
  return out;
}

window.cygnusDesktop.onLog((line) => appendLog(line));
window.cygnusDesktop.onState((state) => {
  // Display only - button enablement is driven by onBusy below, which spans
  // the WHOLE flow (e.g. all of Auth+Parity), not just whichever individual
  // step this text currently names. See beginFlow/endFlow in main.js.
  statusEl.textContent = `State: ${state}`;
});
window.cygnusDesktop.onBusy((isBusy) => setBusy(isBusy));

document.getElementById('btnSetup').addEventListener('click', async () => {
  appendLog('--- Running setup only ---');
  await window.cygnusDesktop.runSetup();
  await refreshAuthGate();
});

document.getElementById('btnParity').addEventListener('click', async () => {
  const cfg = buildConfig();
  const missing = validateForParity(cfg);
  if (missing.length > 0) {
    appendLog(`Validation failed: missing ${missing.join(', ')}`);
    return;
  }

  // A saved session-less Parity launch is guaranteed to fail 5+ minutes in,
  // deep inside token acquisition - exactly the confusing "ran but nothing
  // worked" experience the installer flow was already fixed to avoid. Check
  // first and transparently fall back to Auth + Parity instead of letting it
  // fail.
  const authStatus = await window.cygnusDesktop.getAuthStatus();

  const filterCount = Object.values(cfg.slicerScenario.pages).reduce((n, sels) => n + sels.length, 0);
  if (!authStatus.hasSession) {
    appendLog('--- No saved login session found - running Authentication first, then Parity ---');
  }
  appendLog(filterCount > 0
    ? `--- Running parity (${filterCount} filter selection(s) applied) ---`
    : '--- Running parity (unfiltered) ---');
  if (cfg.lenientTextCompare) {
    appendLog('--- Case insensitive compare enabled (case/space/underscore/hyphen differences ignored) ---');
  }
  document.getElementById('parityResults').innerHTML = '';

  const res = authStatus.hasSession
    ? await window.cygnusDesktop.runParity(cfg)
    : await window.cygnusDesktop.runSetupAndParity(cfg);
  if (res.ok) renderParityResults(res.result);
  await refreshAuthGate();
});

document.getElementById('btnBoth').addEventListener('click', async () => {
  const cfg = buildConfig();
  const missing = validateForParity(cfg);
  if (missing.length > 0) {
    appendLog(`Validation failed: missing ${missing.join(', ')}`);
    return;
  }
  appendLog('--- Running setup + parity ---');
  document.getElementById('parityResults').innerHTML = '';
  const res = await window.cygnusDesktop.runSetupAndParity(cfg);
  if (res.ok) renderParityResults(res.result);
  await refreshAuthGate();
});

// ── Filter discovery + picking ───────────────────────────────────────────────
//
// _filterPicks: Map<fieldKey, {
//   title, isHierarchy, targetLabel, targets, options,
//   byPage: { [pageName]: { visualName } },   // which pages this field was actually found on
//   value: string,                            // currently chosen value ('' = none picked)
// }>
window._filterPicks = new Map();
let _lastGlobalFilters = [];
let _discoveredPages = [];
let _selectedRunPages = new Set();

function fieldKey(title, targetLabel) {
  return `${title}::${targetLabel}`;
}

// Groups raw per-page discovery results into one entry per unique field,
// tracking exactly which pages it was actually found on (this is what lets
// the UI correctly disable a hierarchy field for pages where it doesn't
// have its own slicer visual, while leaving flat fields enabled everywhere).
function groupDiscoveredFields(byPage) {
  const groups = new Map();
  for (const [pageName, slicers] of Object.entries(byPage)) {
    for (const s of slicers) {
      if (s.errorMessage) continue; // not pickable if we couldn't read it
      const key = fieldKey(s.title, s.targetLabel);
      let g = groups.get(key);
      if (!g) {
        g = {
          title: s.title, isHierarchy: s.isHierarchy, targetLabel: s.targetLabel,
          targets: s.targets, options: s.options || [], byPage: {},
        };
        groups.set(key, g);
      }
      g.byPage[pageName] = { visualName: s.name };
      if (s.options && s.options.length > g.options.length) g.options = s.options;
    }
  }
  return groups;
}

function getOrCreatePick(picksMap, key, group) {
  let pick = picksMap.get(key);
  if (!pick) {
    pick = { ...group, value: '' };
    picksMap.set(key, pick);
  } else {
    // Refresh discovery data (byPage/options/targets) on re-discovery, keep the user's picks.
    pick.isHierarchy = group.isHierarchy;
    pick.targetLabel = group.targetLabel;
    pick.targets = group.targets;
    pick.byPage = group.byPage;
    pick.sideTargets = group.sideTargets;
    if (group.options.length > pick.options.length) pick.options = group.options;
  }
  return pick;
}

function getPageUniverse() {
  if (Array.isArray(_discoveredPages) && _discoveredPages.length > 0) {
    return _discoveredPages;
  }
  return getPagesInScope();
}

function getSelectedRunPages() {
  const universe = getPageUniverse();
  if (universe.length === 0) return [];

  const selected = universe.filter(p => _selectedRunPages.has(p));
  // Safety: if selection hasn't been initialized yet, default to all pages.
  return selected.length > 0 ? selected : universe;
}

function getDiscoverPagesCsv() {
  const pagesCsv = document.getElementById('pagesCsv').value.trim();
  if (pagesCsv) return pagesCsv;

  // If user left Pages CSV empty but picked pages in the shared selector,
  // use those pages as discovery scope.
  const selected = getPageUniverse().filter(p => _selectedRunPages.has(p));
  return selected.length > 0 ? selected.join(',') : '';
}

function buildScenarioFromPicks(pageScope) {
  // Fixed name (not a timestamp) — this is what makes each run overwrite the
  // same expected.xlsx/actual.xlsx/parity-summary.xlsx instead of creating a
  // new playwright-report-parity/Cygnus/ui-<timestamp>/ folder every time.
  const scenario = { name: 'ui-filters', pages: {} };
  for (const pageName of pageScope) scenario.pages[pageName] = [];
  if (pageScope.length === 0) return scenario;

  const addPicksFromMap = (picksMap, side) => {
    for (const pick of picksMap.values()) {
      if (!pick.value) continue;
      for (const pageName of pageScope) {
        if (pick.sideTargets && !side) {
          // Shared picker for a cross-report "identical" match: one VALUE
          // picked once, but each side's underlying table.column binding can
          // differ (e.g. a Direct Lake migration renamed the table) even
          // though the on-screen slicer TITLE is identical - matching by
          // title never guaranteed matching bindings. Emitting the source
          // side's targets for both sides is exactly what produced the
          // FILTER FIELD MISMATCH abort (and, before that check existed, a
          // silent one-sided-only filter). Each side must use its OWN
          // discovered targets.
          for (const s of ['source', 'target']) {
            const sd = pick.sideTargets[s];
            if (!sd) continue;
            const sel = { title: pick.title, values: [pick.value], targets: sd.targets, side: s };
            if (pick.isHierarchy) {
              const onThisPage = sd.byPage[pageName];
              if (!onThisPage) continue;
              sel.isHierarchy = true;
              sel.visualName = onThisPage.visualName;
            }
            scenario.pages[pageName].push(sel);
          }
          continue;
        }
        const sel = { title: pick.title, values: [pick.value], targets: pick.targets };
        if (side) sel.side = side; // undefined = applies to both sides, as always
        if (pick.isHierarchy) {
          const onThisPage = pick.byPage[pageName];
          if (!onThisPage) continue; // hierarchy can only be set where its visual exists
          sel.isHierarchy = true;
          sel.visualName = onThisPage.visualName;
        }
        scenario.pages[pageName].push(sel);
      }
    }
  };

  if (window._discoverMode === 'split') {
    // Cross-report discovery found the two reports' filters do NOT match -
    // each side's picks were made against that side's own fields, so they
    // must only ever apply on that side (see selectionsForSide in
    // slicer-config.helpers.ts, which is what actually enforces this at
    // apply time).
    addPicksFromMap(window._filterPicksSource || new Map(), 'source');
    addPicksFromMap(window._filterPicksTarget || new Map(), 'target');
  } else {
    addPicksFromMap(window._filterPicks, undefined);
  }
  return scenario;
}

function renderPageScopeSelector(pageNames) {
  const host = document.getElementById('pageScopeResult');
  host.innerHTML = '';

  if (!Array.isArray(pageNames) || pageNames.length === 0) {
    host.innerHTML = '<div class="placeholder">Run Discover once to load report pages. Selected pages here will be used for all slicers.</div>';
    return;
  }

  const validSelection = pageNames.filter(p => _selectedRunPages.has(p));
  _selectedRunPages = validSelection.length > 0 ? new Set(validSelection) : new Set(pageNames);

  const wrap = document.createElement('div');
  wrap.className = 'discover-page';
  const h3 = document.createElement('h3');
  h3.textContent = 'Selected pages (applies to all slicers below)';
  wrap.appendChild(h3);

  const pagesWrap = document.createElement('div');
  pagesWrap.style.marginTop = '6px';
  pagesWrap.style.display = 'flex';
  pagesWrap.style.flexWrap = 'wrap';
  pagesWrap.style.gap = '10px';

  const allLabel = document.createElement('label');
  allLabel.style.fontSize = '12px';
  allLabel.style.display = 'flex';
  allLabel.style.alignItems = 'center';
  allLabel.style.gap = '4px';
  allLabel.style.fontWeight = '600';

  const allCb = document.createElement('input');
  allCb.type = 'checkbox';

  const syncAllCheckboxState = () => {
    const selected = pageNames.filter(p => _selectedRunPages.has(p)).length;
    allCb.checked = selected === pageNames.length;
    allCb.indeterminate = selected > 0 && selected < pageNames.length;
  };

  const syncPageCheckboxes = () => {
    for (const cb of pagesWrap.querySelectorAll('input[data-page]')) {
      const page = cb.getAttribute('data-page');
      cb.checked = page ? _selectedRunPages.has(page) : false;
    }
  };

  allCb.addEventListener('change', () => {
    _selectedRunPages = allCb.checked ? new Set(pageNames) : new Set();
    syncAllCheckboxState();
    syncPageCheckboxes();
  });

  allLabel.appendChild(allCb);
  allLabel.appendChild(document.createTextNode('All pages'));
  pagesWrap.appendChild(allLabel);

  for (const pageName of pageNames) {
    const label = document.createElement('label');
    label.style.fontSize = '12px';
    label.style.display = 'flex';
    label.style.alignItems = 'center';
    label.style.gap = '4px';

    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.setAttribute('data-page', pageName);
    cb.checked = _selectedRunPages.has(pageName);
    cb.addEventListener('change', () => {
      if (cb.checked) _selectedRunPages.add(pageName);
      else _selectedRunPages.delete(pageName);
      syncAllCheckboxState();
    });

    label.appendChild(cb);
    label.appendChild(document.createTextNode(pageName));
    pagesWrap.appendChild(label);
  }

  syncAllCheckboxState();
  wrap.appendChild(pagesWrap);
  host.appendChild(wrap);
}

function renderGlobalFilters(globalFilters) {
  const el = document.getElementById('globalFiltersResult');
  el.innerHTML = '';
  if (!globalFilters || globalFilters.length === 0) {
    el.innerHTML = '<div class="note">No report-level filters found (or not checked in this discovery mode) — this report likely repeats fields as per-page slicers instead. Use page discovery below.</div>';
    return;
  }
  const pageDiv = document.createElement('div');
  pageDiv.className = 'discover-page';
  const h3 = document.createElement('h3');
  h3.textContent = `Global filters (${globalFilters.length}) — already applied to every page`;
  pageDiv.appendChild(h3);
  for (const f of globalFilters) {
    const row = document.createElement('div');
    row.className = 'discover-slicer';
    const t = f.target || {};
    const label = t.table && t.column ? `${t.table}.${t.column}` : JSON.stringify(t);
    const name = document.createElement('span');
    name.className = 'name';
    name.textContent = label;
    const meta = document.createElement('span');
    meta.className = 'meta';
    meta.textContent = f.values ? f.values.join(', ') : '(hierarchy selection)';
    row.appendChild(name);
    row.appendChild(meta);
    pageDiv.appendChild(row);
  }
  el.appendChild(pageDiv);
}

// ── Parity results (in-page summary of parity-result.json) ──────────────────
//
// A more detailed, always-visible alternative to opening parity-summary.xlsx
// by hand - NOT a replacement for it. Full per-visual detail (sample rows,
// etc.) stays exclusive to the xlsx; this view is page-level, plus a link to
// open each of the three generated files for anyone who wants to go deeper.

const VERDICT_BADGE = {
  pass:          { cls: 'badge lg',        text: '✅ PASS' },
  fail:          { cls: 'badge lg fail',   text: '❌ FAIL' },
  not_comparable: { cls: 'badge lg review', text: '⚠️ NOT COMPARABLE' },
};

const PAGE_STATUS_BADGE = {
  'identical':       { cls: 'badge',        text: 'Identical' },
  'header-only':     { cls: 'badge',        text: 'Labels differ only' },
  'different':       { cls: 'badge fail',   text: 'Different' },
  'only-in-source':  { cls: 'badge review', text: 'Missing in Target' },
  'only-in-target':  { cls: 'badge review', text: 'Only in Target' },
};

async function openOutputFile(filePath, label) {
  const res = await window.cygnusDesktop.openOutputFile(filePath);
  if (!res.ok) appendLog(`--- Could not open ${label}: ${res.error} ---`);
}

function renderParityResults(result) {
  const el = document.getElementById('parityResults');
  el.innerHTML = '';

  if (!result) {
    el.innerHTML = '<div class="note">No parity summary to show yet — either this was a source-only run (nothing was compared), or the run did not get far enough to produce one. Check the log above.</div>';
    return;
  }

  const heading = document.createElement('div');
  heading.className = 'section-title';
  heading.style.marginTop = 'var(--space-3)';
  heading.textContent = 'Parity results';
  el.appendChild(heading);

  for (const sc of result.scenarios || []) {
    const pageDiv = document.createElement('div');
    pageDiv.className = 'discover-page';

    const h3 = document.createElement('h3');
    h3.textContent = sc.name && sc.name !== '(unfiltered)' ? `Scenario: ${sc.name}` : 'Result';
    pageDiv.appendChild(h3);

    const verdictRow = document.createElement('div');
    verdictRow.className = 'parity-verdict';
    const verdict = VERDICT_BADGE[sc.ui.verdict] || VERDICT_BADGE.fail;
    const verdictBadge = document.createElement('span');
    verdictBadge.className = verdict.cls;
    verdictBadge.textContent = verdict.text;
    verdictRow.appendChild(verdictBadge);
    if (sc.ui.differingSheets > 0) {
      const note = document.createElement('span');
      note.className = 'note';
      note.textContent = `${sc.ui.differingSheets} page(s) with real differences`;
      verdictRow.appendChild(note);
    }
    pageDiv.appendChild(verdictRow);

    for (const line of sc.ui.narrative || []) {
      const p = document.createElement('div');
      p.className = 'note';
      p.style.marginBottom = '4px';
      p.textContent = line;
      pageDiv.appendChild(p);
    }

    for (const page of sc.ui.pages || []) {
      const row = document.createElement('div');
      row.className = 'discover-slicer';
      const name = document.createElement('span');
      name.className = 'name';
      name.textContent = page.name;
      const meta = document.createElement('span');
      meta.className = 'meta';
      const statusBadge = PAGE_STATUS_BADGE[page.status] || PAGE_STATUS_BADGE.different;
      const badgeEl = document.createElement('span');
      badgeEl.className = statusBadge.cls;
      badgeEl.textContent = statusBadge.text;
      meta.appendChild(document.createTextNode(`${page.rowsExpected} / ${page.rowsActual} rows `));
      meta.appendChild(badgeEl);
      row.appendChild(name);
      row.appendChild(meta);
      pageDiv.appendChild(row);
    }

    if ((sc.ui.pagesOnlyInSource || []).length || (sc.ui.pagesOnlyInTarget || []).length) {
      const p = document.createElement('div');
      p.className = 'note';
      p.style.marginTop = '6px';
      const parts = [];
      if (sc.ui.pagesOnlyInSource.length) parts.push(`only in Source: ${sc.ui.pagesOnlyInSource.join(', ')}`);
      if (sc.ui.pagesOnlyInTarget.length) parts.push(`only in Target: ${sc.ui.pagesOnlyInTarget.join(', ')}`);
      p.textContent = `Pages ${parts.join('; ')}`;
      pageDiv.appendChild(p);
    }

    const links = document.createElement('div');
    links.className = 'parity-file-links';
    const fileButtons = [
      ['Open Expected', sc.expectedFile],
      ['Open Actual', sc.actualFile],
      ['Open Parity Summary', sc.summaryFile],
    ];
    for (const [label, filePath] of fileButtons) {
      const btn = document.createElement('button');
      btn.className = 'secondary';
      btn.type = 'button';
      btn.textContent = label;
      btn.addEventListener('click', () => openOutputFile(filePath, label));
      links.appendChild(btn);
    }
    pageDiv.appendChild(links);

    el.appendChild(pageDiv);
  }
}

function renderFieldPicker(key, pick) {
  const wrap = document.createElement('div');
  wrap.className = 'discover-page';

  const h3 = document.createElement('h3');
  h3.textContent = pick.title;
  if (pick.isHierarchy) {
    const badge = document.createElement('span');
    badge.className = 'badge hierarchy';
    badge.textContent = 'hierarchy — top level only';
    h3.appendChild(badge);
  }
  wrap.appendChild(h3);

  const metaLine = document.createElement('div');
  metaLine.className = 'note';
  metaLine.textContent = pick.targetLabel;
  wrap.appendChild(metaLine);

  // Value picker
  const select = document.createElement('select');
  select.style.marginTop = '6px';
  select.style.width = '100%';
  select.style.padding = '6px';
  select.style.fontSize = '12px';
  const noneOpt = document.createElement('option');
  noneOpt.value = '';
  noneOpt.textContent = pick.options.length > 0 ? '— select a value —' : '(no values available)';
  select.appendChild(noneOpt);
  for (const v of pick.options) {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    if (v === pick.value) opt.selected = true;
    select.appendChild(opt);
  }
  select.addEventListener('change', () => { pick.value = select.value; });
  wrap.appendChild(select);

  const note = document.createElement('div');
  note.className = 'note';
  note.style.marginTop = '6px';
  note.textContent = 'This value applies to the selected pages above.';
  wrap.appendChild(note);

  return wrap;
}

function renderDiscoverResults(data) {
  window._discoverMode = 'shared';
  const globalFilters = data.globalFilters || [];
  const byPage = data.pages || {};
  const allPages = Array.isArray(data.allPages) ? data.allPages : [];
  _lastGlobalFilters = globalFilters;
  _discoveredPages = allPages.length > 0 ? allPages : Object.keys(byPage);

  renderGlobalFilters(globalFilters);
  renderPageScopeSelector(getPageUniverse());

  const groups = groupDiscoveredFields(byPage);

  const container = document.getElementById('discoverResults');
  container.innerHTML = '';

  if (groups.size === 0) {
    container.innerHTML = '<div class="note">No pages were crawled for slicer visuals yet. Add one or more pages and Discover again to build pickable filters.</div>';
    return;
  }

  const heading = document.createElement('div');
  heading.className = 'section-title';
  heading.style.marginTop = '4px';
  heading.textContent = 'Pick values to apply on your selected pages:';
  container.appendChild(heading);

  for (const [key, group] of groups) {
    const pick = getOrCreatePick(window._filterPicks, key, group);
    container.appendChild(renderFieldPicker(key, pick));
  }
}

/**
 * Cross-report mode: source and target discovered the SAME filter titles
 * (matchDetails.identical), so the user only picks a value once - one
 * shared control per field, not two columns. BUT "same title" does not mean
 * "same underlying table.column": a Direct Lake migration can rename the
 * table backing a field while the on-screen slicer title stays identical
 * (observed in production - a field titled the same on both reports bound
 * to 'Germany - Extract'[Recruitment Area] on the source and to
 * 'V_FACT_APPLICATION_GERMANY'[RECRUITMENT_AREA] on the target). Reusing
 * just one side's targets for both reports is exactly what produced a
 * FILTER FIELD MISMATCH abort (previously, before that check existed, a
 * silently one-sided filter). So each merged pick keeps BOTH sides'
 * targets/visualNames (sideTargets), and buildScenarioFromPicks emits two
 * side-tagged selections from the one shared value - see there.
 */
function renderMatchedDiscoverResults(result) {
  window._discoverMode = 'shared';
  _lastGlobalFilters = [];
  const sourcePage = result.sourcePage;
  const targetPage = result.targetPage;

  renderGlobalFilters(null);
  renderPageScopeSelector(getPageUniverse());

  const sourceGroups = groupDiscoveredFields(sourcePage ? { [sourcePage]: result.sourceFields || [] } : {});
  const targetGroups = groupDiscoveredFields(targetPage ? { [targetPage]: result.targetFields || [] } : {});

  // Pair up by normalized title - matchDetails.identical already guarantees
  // every source field has exactly one title+hierarchy-ness counterpart in
  // target and vice versa, AND that no title resolves to more than one
  // distinct field binding on either side (see duplicateTitleConflicts in
  // matchDiscoveredFields, cross-report-match.helpers.ts) - so this pairing
  // should never need to choose between more than one candidate. The
  // grouping-by-title below still defends against it happening anyway: it
  // is exactly what produced a real bug previously (two "Bewerbungseingang
  // Cluster" slicers on different fields both getting silently paired to
  // the SAME target field). If a title ever has more than one distinct
  // variant on either side here, skip it rather than guess.
  const sourceGroupsByTitleKey = new Map();
  for (const g of sourceGroups.values()) {
    const tk = pageNameKeyJs(g.title);
    const list = sourceGroupsByTitleKey.get(tk);
    if (list) list.push(g); else sourceGroupsByTitleKey.set(tk, [g]);
  }
  const targetGroupsByTitleKey = new Map();
  for (const g of targetGroups.values()) {
    const tk = pageNameKeyJs(g.title);
    const list = targetGroupsByTitleKey.get(tk);
    if (list) list.push(g); else targetGroupsByTitleKey.set(tk, [g]);
  }

  const container = document.getElementById('discoverResults');
  container.innerHTML = '';

  const merged = [];
  for (const [titleKey, srcGroupsForTitle] of sourceGroupsByTitleKey) {
    const tgtGroupsForTitle = targetGroupsByTitleKey.get(titleKey);
    if (!tgtGroupsForTitle) continue; // should not happen when result.identical is true
    if (srcGroupsForTitle.length > 1 || tgtGroupsForTitle.length > 1) {
      console.warn(`[cygnus] Shared picker: "${srcGroupsForTitle[0].title}" has more than one distinct field binding on at least one side - skipping rather than guessing which pairs with which. This should not happen when identical=true.`);
      continue;
    }
    const sourceGroup = srcGroupsForTitle[0];
    const targetGroup = tgtGroupsForTitle[0];
    merged.push({ key: fieldKey(sourceGroup.title, sourceGroup.targetLabel), sourceGroup, targetGroup });
  }

  if (merged.length === 0) {
    container.innerHTML = '<div class="note">No pages were crawled for slicer visuals yet. Add one or more pages and Discover again to build pickable filters.</div>';
    return;
  }

  const heading = document.createElement('div');
  heading.className = 'section-title';
  heading.style.marginTop = '4px';
  heading.textContent = 'Pick values to apply on your selected pages (applied separately to each report\'s own fields):';
  container.appendChild(heading);

  for (const { key, sourceGroup, targetGroup } of merged) {
    const sameLabel = sourceGroup.targetLabel === targetGroup.targetLabel;
    const group = {
      title: sourceGroup.title,
      isHierarchy: sourceGroup.isHierarchy,
      targetLabel: sameLabel
        ? sourceGroup.targetLabel
        : `${sourceGroup.targetLabel}  (target report: ${targetGroup.targetLabel})`,
      options: sourceGroup.options.length >= targetGroup.options.length ? sourceGroup.options : targetGroup.options,
      sideTargets: {
        source: { targets: sourceGroup.targets, byPage: sourceGroup.byPage },
        target: { targets: targetGroup.targets, byPage: targetGroup.byPage },
      },
    };
    const pick = getOrCreatePick(window._filterPicks, key, group);
    container.appendChild(renderFieldPicker(key, pick));
  }
}

/**
 * Cross-report mode: source and target discovered DIFFERENT filters (or a
 * flat-vs-hierarchy mismatch on a same-named field), so a single shared
 * picker would be wrong - render two independent columns instead, with
 * their own pick state, and tag each pick with which side it came from so
 * buildScenarioFromPicks only applies it there.
 */
function renderSplitDiscoverResults(result) {
  window._discoverMode = 'split';
  if (!window._filterPicksSource) window._filterPicksSource = new Map();
  if (!window._filterPicksTarget) window._filterPicksTarget = new Map();

  const container = document.getElementById('discoverResults');
  container.innerHTML = '';

  const d = result.matchDetails || {};
  const reasons = [];
  if (d.onlyInSource && d.onlyInSource.length) reasons.push(`only in Source: ${d.onlyInSource.join(', ')}`);
  if (d.onlyInTarget && d.onlyInTarget.length) reasons.push(`only in Target: ${d.onlyInTarget.join(', ')}`);
  if (d.hierarchyMismatch && d.hierarchyMismatch.length) reasons.push(`type differs (flat vs hierarchy): ${d.hierarchyMismatch.join(', ')}`);
  if (d.duplicateTitleConflicts && d.duplicateTitleConflicts.length) reasons.push(`same title, different field on at least one side: ${d.duplicateTitleConflicts.join(', ')}`);

  const banner = document.createElement('div');
  banner.className = 'mismatch-banner';
  banner.textContent = 'Filters are not identical between the two reports' +
    (reasons.length ? ` (${reasons.join('; ')})` : ' (no common filters were found)') +
    ' — select values separately for each report below.';
  container.appendChild(banner);

  const columns = document.createElement('div');
  columns.className = 'split-columns';

  const buildColumn = (label, pageName, fields, picksMap) => {
    const col = document.createElement('div');
    col.className = 'split-column';
    const h = document.createElement('div');
    h.className = 'section-title';
    h.textContent = pageName ? `${label} filters (page: ${pageName})` : `${label} filters`;
    col.appendChild(h);

    const groups = groupDiscoveredFields(pageName ? { [pageName]: (fields || []) } : {});
    if (groups.size === 0) {
      const p = document.createElement('div');
      p.className = 'note';
      p.textContent = 'No slicers found.';
      col.appendChild(p);
      return col;
    }
    for (const [key, group] of groups) {
      const pick = getOrCreatePick(picksMap, key, group);
      col.appendChild(renderFieldPicker(key, pick));
    }
    return col;
  };

  columns.appendChild(buildColumn('Source', result.sourcePage, result.sourceFields, window._filterPicksSource));
  columns.appendChild(buildColumn('Target', result.targetPage, result.targetFields, window._filterPicksTarget));
  container.appendChild(columns);
}

document.getElementById('btnDiscover').addEventListener('click', async () => {
  // Defense in depth: the button is already disabled without a session, but
  // re-check at click time in case the UI state is ever stale - starting
  // this run with no session is a guaranteed, confusing failure (see
  // updateButtonStates's comment for why).
  await refreshAuthGate();
  if (!_hasSession) {
    appendLog('--- Cannot discover filters: no saved login session. Run Authentication first. ---');
    return;
  }

  const pairName = document.getElementById('pairName').value.trim() || 'Cygnus';
  const skipGlobalCheck = document.getElementById('discoverSkipGlobalCheck')?.checked === true;

  if (bothIdentitiesFilled()) {
    const source = readIdentity('source');
    const target = readIdentity('target');
    appendLog('--- Discovering filters on BOTH reports (comparing source vs target) ---');
    if (skipGlobalCheck) {
      appendLog('--- Skipping report-level filter check (uncheck the box above to re-enable) ---');
    }
    document.getElementById('globalFiltersResult').textContent = '';
    document.getElementById('pageScopeResult').textContent = 'Loading page list…';
    document.getElementById('discoverResults').textContent = 'Discovering source, then target… two fast first-match scans, one after the other.';

    const res = await window.cygnusDesktop.discoverCrossReport(pairName, source, target, skipGlobalCheck);
    if (!res.ok) {
      document.getElementById('pageScopeResult').textContent = '';
      document.getElementById('discoverResults').textContent = `Discovery failed: ${res.error}`;
      return;
    }

    const result = res.result;
    _discoveredPages = dedupePageNames([...(result.sourceAllPages || []), ...(result.targetAllPages || [])]);
    renderPageScopeSelector(getPageUniverse());

    if (result.identical) {
      appendLog(`--- Filters match between reports (page "${result.sourcePage}") — one shared picker ---`);
      renderMatchedDiscoverResults(result);
    } else {
      appendLog('--- Filters do NOT match between reports — select values separately below ---');
      renderSplitDiscoverResults(result);
    }
    return;
  }

  const pagesCsv = getDiscoverPagesCsv();
  const pagesCsvInput = document.getElementById('pagesCsv').value.trim();

  const { identity, side, missing, note } = resolveDiscoverIdentity();
  if (missing.length > 0) {
    appendLog(`Validation failed: missing ${missing.join(', ')}${note ? ` — ${note}` : ''}`);
    return;
  }

  appendLog(pagesCsv
    ? `--- Discovering filters: ${skipGlobalCheck ? 'pages' : 'global check + pages'} [${pagesCsv}] (applies nothing) ---`
    : `--- Discovering filters: ${skipGlobalCheck ? '(global check skipped)' : 'global check only'} (applies nothing) ---`);
  if (!pagesCsvInput && pagesCsv) {
    appendLog('--- Discovery scope source: selected pages from top page selector ---');
  }
  appendLog(`--- Using ${side === 'target' ? 'Target' : 'Source'} report entered above ---`);
  document.getElementById('globalFiltersResult').textContent = skipGlobalCheck ? '' : 'Checking…';
  document.getElementById('pageScopeResult').textContent = 'Loading page list…';
  document.getElementById('discoverResults').textContent = pagesCsv ? 'Crawling pages… this can take a while on a page-heavy report.' : '';
  const res = await window.cygnusDesktop.discoverSlicers(pairName, pagesCsv, identity, side, skipGlobalCheck);
  if (!res.ok) {
    document.getElementById('globalFiltersResult').textContent = '';
    document.getElementById('discoverResults').textContent = `Discovery failed: ${res.error}`;
    return;
  }
  renderDiscoverResults(res.result);
});