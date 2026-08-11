const e = require('electron');
console.log('type:', typeof e);
console.log('app:', typeof e.app);
console.log('ipcMain:', typeof e.ipcMain);
console.log('BrowserWindow:', typeof e.BrowserWindow);
console.log('dialog:', typeof e.dialog);
process.exit(0);
