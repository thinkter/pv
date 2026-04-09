const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('pptxViewer', {
  openPresentation: () => ipcRenderer.invoke('presentation:open'),
  readPresentation: (filePath) => ipcRenderer.invoke('presentation:read', filePath),
  getLaunchPresentation: () => ipcRenderer.invoke('presentation:get-launch-file'),
});
