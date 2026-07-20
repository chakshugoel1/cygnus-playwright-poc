const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('cygnusDesktop', {
  getAuthStatus: () => ipcRenderer.invoke('get-auth-status'),
  runSetup: () => ipcRenderer.invoke('run-setup'),
  runParity: (config) => ipcRenderer.invoke('run-parity', config),
  runSetupAndParity: (config) => ipcRenderer.invoke('run-setup-and-parity', config),
  discoverSlicers: (pairName, pagesCsv, identity, side) => ipcRenderer.invoke('run-discover-slicers', pairName, pagesCsv, identity, side),
  discoverCrossReport: (pairName, sourceIdentity, targetIdentity) => ipcRenderer.invoke('run-discover-cross-report', pairName, sourceIdentity, targetIdentity),
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
});