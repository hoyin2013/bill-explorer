// Try to get electron API through the browser_init builtin
const bi = require('electron/js2c/browser_init');
console.log('browser_init type:', typeof bi);
console.log('browser_init name:', typeof bi === 'function' ? bi.name : 'not fn');
try {
  if (typeof bi === 'function') {
    const result = bi();
    console.log('browser_init() result type:', typeof result);
    if (typeof result === 'object' && result) {
      console.log('keys:', Object.keys(result).slice(0, 15).join(', '));
      console.log('app:', typeof result.app);
      console.log('ipcMain:', typeof result.ipcMain);
      console.log('BrowserWindow:', typeof result.BrowserWindow);
    }
  }
} catch(e) {
  console.log('browser_init() err:', e.message);
}
process.exit(0);
