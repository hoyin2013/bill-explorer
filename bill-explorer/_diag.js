console.log('A. node:', process.versions.node, 'electron:', process.versions.electron);
console.log('B. cwd:', process.cwd());
console.log('C. builtins:', require('module').builtinModules.filter(m => m.includes('electron')).join(' | '));
try {
  const bi = require('electron/js2c/browser_init');
  console.log('D. browser_init type:', typeof bi);
  if (typeof bi === 'function') {
    const r = bi();
    console.log('E. bi() result:', typeof r, r && Object.keys(r).slice(0,10).join(', '));
    if (r) console.log('   app:', typeof r.app, 'ipcMain:', typeof r.ipcMain);
  }
} catch(e) { console.log('D. browser_init err:', e.message, 'code:', e.code || '?'); }
process.exit(0);
