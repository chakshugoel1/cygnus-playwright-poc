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

function setBusy(isBusy) {
  buttons.forEach(b => { b.disabled = isBusy; });
}

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
  const applyGlobalFlatFilters = document.getElementById('applyGlobalFlatFilters')?.checked === true;
  const lenientTextCompare = document.getElementById('lenientTextCompare')?.checked === true;

  return {
    pairName,
    source,
    target,
    pages: pages.length ? pages : undefined,
    slicerScenario,
    applyGlobalFlatFilters,
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

window.cygnusDesktop.onLog((line) => appendLog(line));
window.cygnusDesktop.onState((state) => {
  statusEl.textContent = `State: ${state}`;
  setBusy(state !== 'Idle');
});

document.getElementById('btnSetup').addEventListener('click', async () => {
  appendLog('--- Running setup only ---');
  await window.cygnusDesktop.runSetup();
});

document.getElementById('btnParity').addEventListener('click', async () => {
  const cfg = buildConfig();
  const missing = validateForParity(cfg);
  if (missing.length > 0) {
    appendLog(`Validation failed: missing ${missing.join(', ')}`);
    return;
  }
  const filterCount = Object.values(cfg.slicerScenario.pages).reduce((n, sels) => n + sels.length, 0);
  appendLog(filterCount > 0
    ? `--- Running parity only (${filterCount} filter selection(s) applied) ---`
    : '--- Running parity only (unfiltered) ---');
  if (cfg.applyGlobalFlatFilters) {
    appendLog('--- Flat field filters: apply once at report level (all pages) ---');
  }
  if (cfg.lenientTextCompare) {
    appendLog('--- Lenient text compare enabled (case/space/underscore/hyphen differences ignored) ---');
  }
  await window.cygnusDesktop.runParity(cfg);
});

document.getElementById('btnBoth').addEventListener('click', async () => {
  const cfg = buildConfig();
  const missing = validateForParity(cfg);
  if (missing.length > 0) {
    appendLog(`Validation failed: missing ${missing.join(', ')}`);
    return;
  }
  appendLog('--- Running setup + parity ---');
  await window.cygnusDesktop.runSetupAndParity(cfg);
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

function getOrCreatePick(key, group) {
  let pick = window._filterPicks.get(key);
  if (!pick) {
    pick = { ...group, value: '' };
    window._filterPicks.set(key, pick);
  } else {
    // Refresh discovery data (byPage/options/targets) on re-discovery, keep the user's picks.
    pick.isHierarchy = group.isHierarchy;
    pick.targetLabel = group.targetLabel;
    pick.targets = group.targets;
    pick.byPage = group.byPage;
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

  for (const pick of window._filterPicks.values()) {
    if (!pick.value || pageScope.length === 0) continue;
    for (const pageName of pageScope) {
      const sel = { title: pick.title, values: [pick.value], targets: pick.targets };
      if (pick.isHierarchy) {
        const onThisPage = pick.byPage[pageName];
        if (!onThisPage) continue; // hierarchy can only be set where its visual exists
        sel.isHierarchy = true;
        sel.visualName = onThisPage.visualName;
      }
      scenario.pages[pageName].push(sel);
    }
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
    el.innerHTML = '<div class="note">No report-level filters found — this report likely repeats fields as per-page slicers instead. Use page discovery below.</div>';
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
    const pick = getOrCreatePick(key, group);
    container.appendChild(renderFieldPicker(key, pick));
  }
}

document.getElementById('btnDiscover').addEventListener('click', async () => {
  const pairName = document.getElementById('pairName').value.trim() || 'Cygnus';
  const pagesCsv = getDiscoverPagesCsv();
  const pagesCsvInput = document.getElementById('pagesCsv').value.trim();

  const { identity, side, missing, note } = resolveDiscoverIdentity();
  if (missing.length > 0) {
    appendLog(`Validation failed: missing ${missing.join(', ')}${note ? ` — ${note}` : ''}`);
    return;
  }

  appendLog(pagesCsv
    ? `--- Discovering filters: global check + pages [${pagesCsv}] (applies nothing) ---`
    : '--- Discovering filters: global check only (applies nothing) ---');
  if (!pagesCsvInput && pagesCsv) {
    appendLog('--- Discovery scope source: selected pages from top page selector ---');
  }
  appendLog(`--- Using ${side === 'target' ? 'Target' : 'Source'} report entered above ---`);
  document.getElementById('globalFiltersResult').textContent = 'Checking…';
  document.getElementById('pageScopeResult').textContent = 'Loading page list…';
  document.getElementById('discoverResults').textContent = pagesCsv ? 'Crawling pages… this can take a while on a page-heavy report.' : '';
  const res = await window.cygnusDesktop.discoverSlicers(pairName, pagesCsv, identity, side);
  if (!res.ok) {
    document.getElementById('globalFiltersResult').textContent = '';
    document.getElementById('discoverResults').textContent = `Discovery failed: ${res.error}`;
    return;
  }
  renderDiscoverResults(res.result);
});