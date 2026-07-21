const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('cygnusDesktop', {
  getAuthStatus: () => ipcRenderer.invoke('get-auth-status'),
  runSetup: () => ipcRenderer.invoke('run-setup'),
  runParity: (config) => ipcRenderer.invoke('run-parity', config),
  runSetupAndParity: (config) => ipcRenderer.invoke('run-setup-and-parity', config),
  discoverSlicers: (pairName, pagesCsv, identity, side, skipGlobalCheck) => ipcRenderer.invoke('run-discover-slicers', pairName, pagesCsv, identity, side, skipGlobalCheck),
  discoverCrossReport: (pairName, sourceIdentity, targetIdentity, skipGlobalCheck) => ipcRenderer.invoke('run-discover-cross-report', pairName, sourceIdentity, targetIdentity, skipGlobalCheck),
  openOutputFile: (filePath) => ipcRenderer.invoke('open-output-file', filePath),
  showOutputInFolder: (filePath) => ipcRenderer.invoke('show-output-in-folder', filePath),
  onLog: (handler) => {
    const wrapped = (_event, line) => handler(line);
    ipcRenderer.on('runner-log', wrapped);
    return () => ipcRenderer.removeListener('runner-log', wrapped);
  },
  onState: (handler) => {
    const wrapped = (_event, state) => handler(state);
    ipcRenderer.on('runner-state', wrapped);
    return () => ipcRenderer.removeListener('runner-state', wrapped);
  },
  onBusy: (handler) => {
    const wrapped = (_event, isBusy) => handler(isBusy);
    ipcRenderer.on('runner-busy', wrapped);
    return () => ipcRenderer.removeListener('runner-busy', wrapped);
  },
});