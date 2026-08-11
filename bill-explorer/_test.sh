const bi = require('electron/js2c/browser_init');
console.log('bi type:', typeof bi);
if (typeof bi === 'function') {
  try {
    const r = bi();
    console.log('result type:', typeof r);
    if (typeof r === 'object' && r) {
      console.log('keys:', Object.keys(r).slice(0,15).join(', '));
      console.log('app:', typeof r.app, 'ipcMain:', typeof r.ipcMain);
    }
  } catch(e) { console.log('bi() err:', e.message); }
}
process.exit(0);
