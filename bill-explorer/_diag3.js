console.log('A. node:', process.versions.node, 'electron:', process.versions.electron);
try { const e = require('electron'); console.log('B. require("electron"):', typeof e, String(e).slice(0,60)); } catch(err) { console.log('B. err:', err.message, 'code:', err.code); }
try { const bi = require('electron/js2c/browser_init'); console.log('C. browser_init type:', typeof bi); if (typeof bi === 'function') { try { const r = bi(); console.log('D. bi() result:', typeof r, r && Object.keys(r).slice(0,10).join(', ')); if (r) console.log('   app:', typeof r.app, 'ipcMain:', typeof r.ipcMain); } catch(e2) { console.log('D. bi() err:', e2.message); } } } catch(err) { console.log('C. err:', err.message); }
process.exit(0);
