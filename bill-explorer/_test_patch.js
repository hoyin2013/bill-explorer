const e = require('electron');
console.log('type:', typeof e);
if (typeof e === 'object') {
  console.log('app:', typeof e.app, 'ipcMain:', typeof e.ipcMain, 'BrowserWindow:', typeof e.BrowserWindow);
}
process.exit(0);
