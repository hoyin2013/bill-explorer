// Try every possible way to get electron API
console.log('A. require("electron"):', typeof require('electron'));
console.log('B. require("electron") value:', String(require('electron')).slice(0, 80));
