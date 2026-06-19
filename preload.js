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
});
