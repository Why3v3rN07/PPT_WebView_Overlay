const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  testOverlay:      (data) => ipcRenderer.send('test-overlay', data),
  closeOverlays:    ()     => ipcRenderer.send('close-overlays'),
  startMonitoring:  ()     => ipcRenderer.send('start-monitoring'),
  stopMonitoring:   ()     => ipcRenderer.send('stop-monitoring'),
  getSetting:       (key)        => ipcRenderer.invoke('get-setting', key),
  setSetting:       (key, value) => ipcRenderer.invoke('set-setting', key, value),
});