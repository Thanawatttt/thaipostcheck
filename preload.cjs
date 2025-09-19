const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('thaipost', {
  trackParcel: (trackingNumber) => ipcRenderer.invoke('track-parcel', trackingNumber),
  savePDF: (htmlContent) => ipcRenderer.invoke('save-pdf', htmlContent),
  captureAllParcels: (parcelCount) => ipcRenderer.invoke('capture-all-parcels', parcelCount),
  importExcel: () => ipcRenderer.invoke('import-excel'),
  importPDF: () => ipcRenderer.invoke('import-pdf'),
  getSheetNames: () => ipcRenderer.invoke('get-sheet-names'),
  selectSheet: (sheetName) => ipcRenderer.invoke('select-sheet', sheetName),
  cancelSheetSelection: () => ipcRenderer.invoke('cancel-sheet-selection')
});

contextBridge.exposeInMainWorld('electronAPI', {
  ipcRenderer: ipcRenderer
});