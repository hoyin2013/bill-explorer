console.log('A. node:', process.versions.node, 'electron:', process.versions.electron);
try {
  const bi = require('electron/js2c/browser_init');
  console.log('B. browser_init type:', typeof bi);
  if (typeof bi === 'function') {
    const r = bi();
    console.log('C. bi() result:', typeof r, r && Object.keys(r).slice(0,10).join(', '));
    if (r) console.log('   app:', typeof r.app, 'ipcMain:', typeof r.ipcMain);
  }
} catch(e) { console.log('B. browser_init err:', e.message); }
process.exit(0);
