const fs = require('fs');

const content = fs.readFileSync('MainApp_RecursosHumanosHTML.html', 'utf8');
const lines = content.split('\n');
const MARKER = '// ════════════════════════ MÓDULOS — rebuild_main.js ════════════════════════ //';

const antes = lines.slice(0, 866).join('\n');
const despues = lines.slice(lines.length - 1).join('\n');
const nuevo = antes + '\n' + MARKER + '\n' + despues;
fs.writeFileSync('MainApp_RecursosHumanosHTML.html', nuevo, 'utf8');
console.log('MainApp actualizado. Líneas: ' + nuevo.split('\n').length);
