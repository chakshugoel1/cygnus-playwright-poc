const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn, execSync } = require('child_process');

function findWorkspaceRoot() {
  const explicitRoot = (process.env.CYGNUS_WORKSPACE_ROOT ?? '').trim();
  if (explicitRoot) {
    return path.resolve(explicitRoot);
  }

  let current = app.isPackaged ? path.dirname(process.execPath) : path.resolve(__dirname, '..');
  while (true) {
    const hasWorkspaceMarkers =
      fs.existsSync(path.join(current, 'package.json')) &&
      fs.existsSync(path.join(current, 'playwright.config.ts')) &&
      fs.existsSync(path.join(current, 'tests'));

    if (hasWorkspaceMarkers) {
      return current;
    }

    const parent = path.dirname(current);
    if (parent === current) break;
    current = parent;
  }

  return path.resolve(__dirname, '..');
}

const WORKSPACE_ROOT = findWorkspaceRoot();
const RUNTIME_DIR = path.join(WORKSPACE_ROOT, '.runtime');
const RUNTIME_FILE = path.join(RUNTIME_DIR, 'parity-runtime-config.json');
// Separate file from RUNTIME_FILE (parity's) so a Discover run never clobbers
// an in-progress/just-used Parity config, or vice versa.
const DISCOVER_RUNTIME_FILE = path.join(RUNTIME_DIR, 'discover-runtime-config.json');

let mainWindow = null;
let isRunning = false;

function emitLog(line) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('runner-log', line);
  }
}

function emitState(state) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('runner-state', state);
  }
}

function ensureRuntimeDir() {
  fs.mkdirSync(RUNTIME_DIR, { recursive: true });
}

function writeRuntimeConfig(config) {
  ensureRuntimeDir();
  const payload = {
    pairName: config.pairName,
    source: config.source,
    target: config.target,
    pages: config.pages,
  };
  fs.writeFileSync(RUNTIME_FILE, JSON.stringify(payload, null, 2), 'utf8');
  return RUNTIME_FILE;
}

function getRunCommand(scriptName, extraEnv) {
  const env = {
    ...process.env,
    ...extraEnv,
  };

  const npmCmdPath = resolveNpmCmd();
  // Quoted because the resolved path almost always contains spaces
  // ("Program Files"). This full string is handed to the shell as ONE line
  // (via shell: true below) — do not split this into a separate args array
  // for spawn(), that re-escapes the quotes already in this string and
  // produces a literal `\"...\"` that cmd.exe can't parse.
  const command = `"${npmCmdPath}" run ${scriptName}`;

  return { command, env };
}

// Finds npm.cmd on THIS machine instead of assuming Node.js lives at
// C:\Program Files\nodejs (only true if the installer defaults were used).
// Order: 1) whatever `where` resolves via this process's own PATH — correct
// on any machine where Node works normally from a terminal; 2) a short list
// of well-known install locations, for the case where Electron (especially
// when double-clicked or launched from a shortcut rather than a terminal)
// inherited a stale/incomplete PATH; 3) throw a clear, actionable error
// instead of spawning a command that's guaranteed to fail with a cryptic
// "npm is not recognized" deep inside cmd.exe.
let _npmCmdPath = null;
function resolveNpmCmd() {
  if (_npmCmdPath) return _npmCmdPath;

  try {
    const out = execSync('where npm.cmd', { encoding: 'utf8', windowsHide: true });
    const first = out.split(/\r?\n/).map((s) => s.trim()).find(Boolean);
    if (first && fs.existsSync(first)) {
      _npmCmdPath = first;
      return _npmCmdPath;
    }
  } catch {
    // `where` found nothing on PATH in this process — fall through to the
    // well-known-location check below.
  }

  const wellKnownDirs = [
    process.env['ProgramFiles']       && path.join(process.env['ProgramFiles'], 'nodejs'),
    process.env['ProgramFiles(x86)']  && path.join(process.env['ProgramFiles(x86)'], 'nodejs'),
    process.env['LOCALAPPDATA']       && path.join(process.env['LOCALAPPDATA'], 'Programs', 'nodejs'),
    process.env['APPDATA']            && path.join(process.env['APPDATA'], 'npm'),
  ].filter(Boolean);

  for (const dir of wellKnownDirs) {
    const candidate = path.join(dir, 'npm.cmd');
    if (fs.existsSync(candidate)) {
      _npmCmdPath = candidate;
      return _npmCmdPath;
    }
  }

  throw new Error(
    'Could not find npm on this machine (checked PATH and the usual Node.js ' +
    'install locations). Make sure Node.js is installed, then close and ' +
    'reopen this app.',
  );
}

function runScript(scriptName, extraEnv = {}, stateLabel = scriptName) {
  return new Promise((resolve, reject) => {
    if (isRunning) {
      reject(new Error('A run is already in progress.'));
      return;
    }

    let run;
    try {
      run = getRunCommand(scriptName, extraEnv);
    } catch (err) {
      emitLog(`[desktop-runner] ${err.message}`);
      reject(err);
      return;
    }

    isRunning = true;
    emitState(`Running: ${stateLabel}`);

    const child = spawn(run.command, {
      cwd: WORKSPACE_ROOT,
      env: run.env,
      windowsHide: true,
      shell: true,
    });

    child.stdout.on('data', (d) => emitLog(d.toString().replace(/\r?\n$/, '')));
    child.stderr.on('data', (d) => emitLog(d.toString().replace(/\r?\n$/, '')));

    child.on('error', (err) => {
      isRunning = false;
      emitState('Idle');
      reject(err);
    });

    child.on('close', (code) => {
      isRunning = false;
      emitState('Idle');
      if (code === 0) {
        emitLog(`[desktop-runner] ${scriptName} completed successfully.`);
        resolve();
      } else {
        reject(new Error(`${scriptName} failed with exit code ${code}`));
      }
    });
  });
}

async function runSetup() {
  emitLog('[desktop-runner] Starting setup via npm run test:setup');
  await runScript('test:setup', {}, 'test:setup');
}

async function runParity(config) {
  const runtimeConfigPath = writeRuntimeConfig(config);
  emitLog(`[desktop-runner] Runtime parity config written: ${runtimeConfigPath}`);

  const extraEnv = {
    CYGNUS_UI_RUNTIME_OVERRIDE: '1',
    CYGNUS_UI_RUNTIME_CONFIG_PATH: runtimeConfigPath,
  };

  if (config.slicerScenario && Object.keys(config.slicerScenario.pages || {}).length > 0) {
    ensureRuntimeDir();
    const scenarioPath = path.join(RUNTIME_DIR, 'slicer-scenario.json');
    fs.writeFileSync(scenarioPath, JSON.stringify({ scenarios: [config.slicerScenario] }, null, 2), 'utf8');
    emitLog(`[desktop-runner] Filter scenario written: ${scenarioPath}`);
    extraEnv.CYGNUS_SLICER_OVERRIDE = '1';
    extraEnv.CYGNUS_SLICER_CONFIG_PATH = scenarioPath;
  }

  if (config.applyGlobalFlatFilters) {
    extraEnv.CYGNUS_SLICER_GLOBAL_FILTERS = '1';
    emitLog('[desktop-runner] Flat field filters will be applied at report level (all pages).');
  }

  if (config.lenientTextCompare) {
    extraEnv.CYGNUS_COMPARE_TEXT_LENIENT = '1';
    emitLog('[desktop-runner] Lenient text compare enabled.');
  } else {
    extraEnv.CYGNUS_COMPARE_TEXT_LENIENT = '0';
  }

  emitLog('[desktop-runner] Starting parity via npm run parity');
  await runScript('parity', extraEnv, 'parity');
}

async function runDiscoverSlicers(pairName, pagesCsv, identity, side) {
  emitLog('[desktop-runner] Starting discovery via npm run discover:slicers');
  const extraEnv = {};

  if (identity) {
    // Mirrors runParity()'s override mechanism: point comparison-config's
    // getActivePair() at the report entered in the UI instead of the
    // hardcoded PAIRS[0] entry in comparison-config.helpers.ts. The runtime
    // loader requires BOTH source and target to be present (see
    // readRuntimeParityConfig) even though discovery only ever reads ONE side
    // — the other side is a harmless placeholder here, discovery never
    // touches it.
    ensureRuntimeDir();
    const payload = {
      pairName: (pairName || 'Cygnus').trim(),
      source: identity,
      target: identity,
    };
    fs.writeFileSync(DISCOVER_RUNTIME_FILE, JSON.stringify(payload, null, 2), 'utf8');
    extraEnv.CYGNUS_UI_RUNTIME_OVERRIDE = '1';
    extraEnv.CYGNUS_UI_RUNTIME_CONFIG_PATH = DISCOVER_RUNTIME_FILE;
    extraEnv.DISCOVER_SIDE = side === 'target' ? 'target' : 'source';
    emitLog(`[desktop-runner]   report: using ${side === 'target' ? 'Target' : 'Source'} fields entered in the UI (tenant ${identity.tenantId})`);
  } else {
    emitLog('[desktop-runner]   report: using the pair hardcoded in comparison-config.helpers.ts (no identity entered in the UI)');
  }

  const pages = (pagesCsv || '').trim();
  if (pages) {
    extraEnv.DISCOVER_PAGES = pages;
    emitLog(`[desktop-runner]   scope: pages [${pages}]`);
  } else {
    emitLog('[desktop-runner]   scope: global check only (no pages named — add pages to also crawl them)');
  }
  await runScript('discover:slicers', extraEnv, 'discover:slicers');

  const resolvedPairName = (pairName || 'Cygnus').trim();
  const resultPath = path.join(WORKSPACE_ROOT, 'playwright-report-parity', resolvedPairName, 'discovered-slicers.json');
  if (!fs.existsSync(resultPath)) {
    throw new Error(`Discovery finished but no results file was found at ${resultPath}. Check the pair name matches your comparison-config.helpers.ts.`);
  }
  const raw = fs.readFileSync(resultPath, 'utf8');
  return JSON.parse(raw);
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 860,
    backgroundColor: '#f7f9fc',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'));
}

ipcMain.handle('run-setup', async () => {
  try {
    await runSetup();
    return { ok: true };
  } catch (e) {
    emitLog(`[desktop-runner] Setup error: ${e.message}`);
    return { ok: false, error: e.message };
  }
});

ipcMain.handle('run-parity', async (_event, config) => {
  try {
    await runParity(config);
    return { ok: true };
  } catch (e) {
    emitLog(`[desktop-runner] Parity error: ${e.message}`);
    return { ok: false, error: e.message };
  }
});

ipcMain.handle('run-discover-slicers', async (_event, pairName, pagesCsv, identity, side) => {
  try {
    const result = await runDiscoverSlicers(pairName, pagesCsv, identity, side);
    return { ok: true, result };
  } catch (e) {
    emitLog(`[desktop-runner] Discover error: ${e.message}`);
    return { ok: false, error: e.message };
  }
});

ipcMain.handle('run-setup-and-parity', async (_event, config) => {
  try {
    await runSetup();
    await runParity(config);
    return { ok: true };
  } catch (e) {
    emitLog(`[desktop-runner] Combined flow error: ${e.message}`);
    return { ok: false, error: e.message };
  }
});

app.whenReady().then(() => {
  app.setName('Cygnus Desktop Runner');
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});