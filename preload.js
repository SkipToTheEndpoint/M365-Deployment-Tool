const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  openFileDialog: () => ipcRenderer.invoke('dialog:openFile'),
  openFolderDialog: () => ipcRenderer.invoke('dialog:openFolder'),
  getDefaultPaths: () => ipcRenderer.invoke('get:defaultPaths'),
  openFolder: (folderPath) => ipcRenderer.invoke('open:folder', folderPath),
  listXMLFiles: () => ipcRenderer.invoke('list:xmlFiles'),
  executeDeployment: (payload) => ipcRenderer.invoke('execute:deployment', payload),
  stopDeployment: () => ipcRenderer.invoke('stop:deployment'),
  onLog: (callback) => ipcRenderer.on('log', (_event, data) => callback(data)),
  onDetails: (callback) => ipcRenderer.on('details', (_event, data) => callback(data)),
  onStatus: (callback) => ipcRenderer.on('status', (_event, data) => callback(data)),
  onProgress: (callback) => ipcRenderer.on('progress', (_event, data) => callback(data)),
  onComplete: (callback) => ipcRenderer.on('complete', (_event, data) => callback(data)),
});
