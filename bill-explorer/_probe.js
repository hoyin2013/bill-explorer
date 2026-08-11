const { app } = require('electron');
console.log('app:', typeof app, app && app.getName ? app.getName() : 'N/A');
process.exit(0);
