const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('api', {
  getSettings: ()  => ipcRenderer.invoke('get-settings'),
  saveSettings: (s) => ipcRenderer.invoke('save-settings', s),
  minimize:     ()           => ipcRenderer.invoke('minimize'),
  maximize:     ()           => ipcRenderer.invoke('maximize'),
  close:        ()           => ipcRenderer.invoke('close'),
  openRecord:    (url, title)              => ipcRenderer.invoke('open-record',    { url, title }),
  openWorkOrder: (workOrderId, orgUrl, title) => ipcRenderer.invoke('open-workorder', { workOrderId, orgUrl, title }),
  openWorkOrderDirect: (workOrderId, orgUrl, title) => ipcRenderer.invoke('open-workorder-direct', { workOrderId, orgUrl, title }),
  openContact: (contactId, orgUrl, title) => ipcRenderer.invoke('open-contact', { contactId, orgUrl, title }),
  openTeamMember: (name, orgUrl, title) => ipcRenderer.invoke('open-team-member', { name, orgUrl, title }),
  openExternal: (url) => ipcRenderer.invoke('open-external', url),
  getVersion: () => ipcRenderer.invoke('get-version'),
  checkForUpdate: () => ipcRenderer.invoke('updater-check'),
  applyUpdate:    (asset) => ipcRenderer.invoke('updater-apply', asset),
  onUpdateProgress: (cb) => ipcRenderer.on('updater-progress', (_, p) => cb(p)),
  // Register the main window's Dynamics webview as the shared session for child windows.
  registerApiWebview: (id) => ipcRenderer.invoke('register-api-webview', id),
});
