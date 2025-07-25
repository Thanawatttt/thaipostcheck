const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('thaipost', {
  trackParcel: (trackingNumber) => ipcRenderer.invoke('track-parcel', trackingNumber)
}); 