const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('api', {
  minimize:     () => ipcRenderer.invoke('wo-minimize'),
  maximize:     () => ipcRenderer.invoke('wo-maximize'),
  close:        () => ipcRenderer.invoke('wo-close'),
  openExternal: (url) => ipcRenderer.invoke('open-external', url),
  openWorkOrder: (workOrderId, orgUrl, title) => ipcRenderer.invoke('open-workorder', { workOrderId, orgUrl, title }),
  // Route Web API calls through the main window's warm Dynamics session.
  xrmExec: (script) => ipcRenderer.invoke('xrm-exec', script),
  xrmReady: () => ipcRenderer.invoke('xrm-ready'),
});
