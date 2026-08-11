console.log('require("electron"):', typeof require('electron'));
console.log('require("electron/main"):', typeof require('electron/main'));
console.log('require("electron/common"):', typeof require('electron/common'));
const { app } = require('electron/main');
console.log('app:', typeof app, app && app.getName ? app.getName() : 'N/A');
process.exit(0);
