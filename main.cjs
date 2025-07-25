const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const { trackParcel } = require('./thailandpost.cjs');

function createWindow() {
  const win = new BrowserWindow({
    width: 900,
    height: 700,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      nodeIntegration: false,
      contextIsolation: true,
    }
  });
  win.loadFile('renderer/index.html');
}

app.whenReady().then(createWindow);

ipcMain.handle('track-parcel', async (event, trackingNumber) => {
  return await trackParcel(trackingNumber);
}); 