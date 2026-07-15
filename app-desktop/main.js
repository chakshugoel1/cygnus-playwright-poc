const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

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

function getWindowsNpmCommand(scriptName, extraEnv) {
  const env = {
    ...process.env,
    ...extraEnv,
  };

  const command = [
    'set PATH=C:\\Progra~1\\nodejs;%PATH%',
    `C:\\Progra~1\\nodejs\\npm.cmd run ${scriptName}`,
  ].join(' && ');

  return {
    cmd: process.env.ComSpec || 'cmd.exe',
    args: ['/c', command],
    env,
  };
}

function runScript(scriptName, extraEnv = {}, stateLabel = scriptName) {
  return new Promise((resolve, reject) => {
    if (isRunning) {
      reject(new Error('A run is already in progress.'));
      return;
    }
    isRunning = true;
    emitState(`Running: ${stateLabel}`);

    const run = getWindowsNpmCommand(scriptName, extraEnv);
    const child = spawn(run.cmd, run.args, {
      cwd: WORKSPACE_ROOT,
      env: run.env,
      windowsHide: true,
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
  emitLog('[desktop-runner] Starting parity via npm run parity');

  await runScript('parity', {
    CYGNUS_UI_RUNTIME_OVERRIDE: '1',
    CYGNUS_UI_RUNTIME_CONFIG_PATH: runtimeConfigPath,
  }, 'parity');
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
