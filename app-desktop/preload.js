const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('cygnusDesktop', {
  runSetup: () => ipcRenderer.invoke('run-setup'),
  runParity: (config) => ipcRenderer.invoke('run-parity', config),
  runSetupAndParity: (config) => ipcRenderer.invoke('run-setup-and-parity', config),
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
