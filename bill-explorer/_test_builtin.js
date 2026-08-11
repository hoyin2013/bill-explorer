try { const e = require('electron'); console.log('type:', typeof e); if (typeof e === 'object') { console.log('app:', typeof e.app, 'ipcMain:', typeof e.ipcMain); } else { console.log('value:', String(e).slice(0,60)); } }
catch(err) { console.log('err:', err.message, 'code:', err.code); }
process.exit(0);
