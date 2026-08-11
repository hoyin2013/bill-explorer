const e = require('electron');
console.log('type:', typeof e);
if (typeof e === 'object') {
  console.log('app:', typeof e.app);
  console.log('ipcMain:', typeof e.ipcMain);
} else {
  console.log('value:', String(e).slice(0,60));
}
process.exit(0);
