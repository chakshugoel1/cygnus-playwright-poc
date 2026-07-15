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

function buildConfig() {
  const source = readIdentity('source');
  const target = readIdentity('target');
  const pairName = document.getElementById('pairName').value.trim() || 'Cygnus';
  const pagesCsv = document.getElementById('pagesCsv').value.trim();
  const pages = pagesCsv ? pagesCsv.split(',').map(s => s.trim()).filter(Boolean) : undefined;

  return { pairName, source, target, pages };
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
  appendLog('--- Running parity only ---');
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
