console.log('1. execPath:', process.execPath);
console.log('2. cwd:', process.cwd());
console.log('3. builtins:', require('module').builtinModules.filter(m => m.includes('electron')).join(' | '));
console.log('4. electron via pkg:', typeof require('electron'), String(require('electron')).slice(0,60));
try { const bi = require('electron/js2c/browser_init'); console.log('5. browser_init type:', typeof bi); if (typeof bi === 'function') { const r = bi(); console.log('6. browser_init() result:', typeof r, r && Object.keys(r).slice(0,8).join(', ')); if (r) console.log('   app:', typeof r.app, 'ipcMain:', typeof r.ipcMain); } } catch(e) { console.log('5. browser_init err:', e.message); }
process.exit(0);